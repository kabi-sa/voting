# Power Automate Receiver — Complete Build Specification

The Microsoft alternative to the Google Apps Script receiver. Votes are written directly into
**KABi-EOM-Master.xlsx** on IC's OneDrive (table `tbl_Votes` — the workbook's monthly sheets and
dashboard recalculate automatically). Requires a **Power Automate Premium** license for the IC
account only (HTTP trigger).

> **Contract:** the voting page talks to ONE endpoint URL and expects the exact API below —
> identical to the Apps Script receiver. Implement all actions; the last section explains what
> breaks if you build only a subset.

---

## 1. One flow: "EOM — Receiver"

**Trigger:** *When an HTTP request is received* — Method: **Any**. After first save, copy the
**HTTP URL** → paste into `SUBMIT_ENDPOINT` in `index.html`.

**Secret:** add a Compose named `SECRET` holding the IC access code (e.g. a long random string).
It lives only inside the flow definition — never in the page. IC types this code at the console gate.

**First action — route by method:** Condition on `triggerOutputs()['method']` = `GET` → branch A,
else branch B (POST).

**Every Response action in this flow must include these headers** (otherwise the browser blocks
the page from reading the reply):
```
Access-Control-Allow-Origin: *
Content-Type: application/json
```

---

## 2. Branch A — GET actions (query `action`)

Read the action: `triggerOutputs()['queries']?['action']` (default `state`). Switch:

| action | Auth | Steps | Response body |
|---|---|---|---|
| `state` | none | List rows `tbl_Cycles`, Filter `Status eq 'Open'`. If one row → its CycleID; else take the latest row (closed). | `{"ok":true,"cycle":"2026-09","open":true}` |
| `auth` | `queries.code` equals `outputs('SECRET')` | — | `{"ok":true}` or `{"ok":false,"error":"unauthorized"}` |
| `roster` | none (voters need it) | List rows `tbl_Employees` where `EmploymentStatus eq 'Active'`; Select → `{id: EmployeeID, name: FullName, email: KABiEmail, jt: JobTitle}` | `{"ok":true,"roster":[ ... ]}` |
| `votes` | code = SECRET | List rows `tbl_Votes`, Filter `CycleID eq '<queries.cycle>' and Status eq 'Accepted'`; Select → payload shape below | `{"ok":true,"votes":[ ... ]}` |

`votes` item shape (what the console expects):
```json
{ "cycle":"2026-09","voterName":"…","voterEmail":"…","submittedAt":"2026-09-14T10:22",
  "selections":{"EOM":{"id":"EMP003","name":"…","why":"…"},"BUILD":{…},"INNOVATE":{…},"KABITAL":{…},"GROW":{…}} }
```

---

## 3. Branch B — POST actions (JSON body)

Parse JSON with this schema (add `action`, `code`, `emp`, `id`, `open` as optional properties):
```json
{ "type":"object","properties":{
  "action":{"type":"string"},"code":{"type":"string"},
  "cycle":{"type":"string"},"voterName":{"type":"string"},"voterEmail":{"type":"string"},
  "submittedAt":{"type":"string"},
  "selections":{"type":"object"},
  "emp":{"type":"object"},"id":{"type":"string"},"open":{"type":"boolean"}}}
```
Switch on `body('Parse_JSON')?['action']` (empty/absent = `vote`):

### `vote` (default) — validation chain, then write
Any failure → Response 200 with `{"ok":false,"error":"<key>"}` using exactly these keys
(the page maps them to user messages): `ineligible`, `duplicate`, `self-vote`,
`values must be different colleagues`, `closed`.

1. **Domain:** `endsWith(toLower(voterEmail),'@kabi.ai')` → else `ineligible`
2. **Roster:** List rows `tbl_Employees`, Filter `KABiEmail eq '<email>'` → empty = `ineligible`; keep `EmployeeID` as `voterId`
3. **Cycle open:** cycle row Status `Open` and `utcNow() < ClosesAt` → else `closed`
4. **Duplicate:** List rows `tbl_Votes`, Filter `CycleID eq '<cycle>' and VoterEmployeeID eq '<voterId>' and Status eq 'Accepted'` → non-empty = `duplicate`
5. **Self-vote:** `voterId` equals any of the five `selections.*.id` → `self-vote`
6. **Distinct values:** the four ids BUILD/INNOVATE/KABITAL/GROW must all differ → else `values must be different colleagues`
7. **Add a row into a table** → `tbl_Votes`: VoteID `concat('V-',workflow()['run']['name'])`,
   CycleID, VoterEmployeeID, SubmittedAt `utcNow()`, the five candidate IDs, the five Why columns,
   Status `Accepted`, AcceptedKey `concat(cycle,'|',voterId)`
8. Response `{"ok":true}`

### IC actions (all require `body.code` = SECRET, else `unauthorized`)
- `setState`: if `open=false` → Update a row `tbl_Cycles` (Key CycleID) Status=`Closed`; if `open=true` on existing cycle → Status=`Open`; if `cycle` is new → Add a row (Status `Open`, OpensAt `utcNow()`, ClosesAt = end of that month) → `{"ok":true}`.
  *Note:* the master's per-month report tab is still a 10-second manual duplicate (or the SeedCycle Office Script) — votes and dashboards work regardless.
- `addEmp`: validate email domain + uniqueness in `tbl_Employees` → Add a row (EmployeeID, FullName, KABiEmail, JobTitle, EmploymentStatus `Active`, EligibleDefault `Yes`) → `{"ok":true}`
- `delEmp`: Delete a row `tbl_Employees` (Key Column EmployeeID = `body.id`) → `{"ok":true}`

---

## 4. Minimum viable vs. full console

| You implement | Employees can vote | IC console in live mode |
|---|---|---|
| `vote` only | ✅ fully | ❌ can't even log in (`auth` fails) |
| `vote` + `auth` + `state` | ✅ | Login + cycle display work; no live votes/roster mgmt |
| **All actions (this doc)** | ✅ | ✅ everything: live results, who voted, new cycle, open/close, add/remove employees |

Recommendation: build the full set — it is one Switch with small branches, ~half a day.

---

## 5. Test with curl before wiring the page
```bash
curl "<URL>?action=state"
curl "<URL>?action=auth&code=WRONG"        # → {"ok":false,...}
curl "<URL>?action=roster"
curl -X POST <URL> -H "Content-Type: text/plain" -d '{"cycle":"2026-09","voterName":"Test","voterEmail":"x@kabi.ai","selections":{...}}'
# repeat the same POST → {"ok":false,"error":"duplicate"}
```
Then run the 5-minute pre-launch checklist in README.md against the deployed page.
