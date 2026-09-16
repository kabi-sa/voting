/**
 * KABi Employee of the Month — ALL-IN-ONE (page + receiver) on Google Apps Script
 * --------------------------------------------------------------------------------
 * ONE deployment serves BOTH the voting page and the vote receiver.
 * Setup: sheets.new → Extensions → Apps Script →
 *   1) paste this into Code.gs
 *   2) File + → HTML → name it exactly "Index" → paste Index.html into it
 *   3) set SECRET below (the IC console access code)
 *   4) Deploy → New deployment → Web app → Execute as: Me · Access: Anyone
 * The Web app URL IS the voting link — send it to employees. Nothing else to configure.
 * ---------------------------------------------------------------
 * What it does:
 *  - Receives each vote from the GitHub Pages page and appends it as a row
 *    in a KABi-styled sheet named per month ("Sep 2026", "Oct 2026", ...).
 *  - Validates server-side: @kabi.ai email, roster membership, no duplicate
 *    vote per person per month, no self-vote, voting open.
 *  - Serves the roster and the current cycle state to the page, and lets the
 *    IC console (with the IC access code) add/remove employees, start a new
 *    monthly cycle, open/close voting, and pull live votes.
 *
 * Setup (once):
 *  1. sheets.new  →  Extensions → Apps Script  →  paste this file, Save.
 *  2. Set SECRET below to the SAME value as IC_PASSCODE in index.html.
 *  3. Deploy → New deployment → Web app:
 *       Execute as: Me        Who has access: Anyone
 *     Copy the Web app URL → paste into SUBMIT_ENDPOINT in index.html.
 *  4. In the spreadsheet, fill the "Roster" sheet (auto-created on first run):
 *       ID | Name | Email | JobTitle      (emails must be @kabi.ai)
 *  Monthly Excel: File → Download → Microsoft Excel — every month is its own sheet.
 */

const SECRET = "KABI-IC-2026";            // must equal IC_PASSCODE in index.html
const DOMAIN = "@kabi.ai";

const CATS = [
  { code: "EOM",      name: "Employee of the Month",      color: "#3D3185", tint: "#D8D6EA" },
  { code: "BUILD",    name: "We Build KABians",           color: "#1E5EA2", tint: "#D2DFEC" },
  { code: "INNOVATE", name: "We Innovate to Transform",   color: "#1F8C9C", tint: "#D2E8EB" },
  { code: "KABITAL",  name: "We Put Human KABital First", color: "#675EAA", tint: "#E1DEEE" },
  { code: "GROW",     name: "We Grow Together",           color: "#2E8A66", tint: "#D5E8E0" },
];
const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];

/* ---------------- helpers ---------------- */
const out = (obj) => ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
const props = () => PropertiesService.getScriptProperties();

function getState() {
  const p = props();
  let cycle = p.getProperty("cycle");
  if (!cycle) { cycle = Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM"); p.setProperty("cycle", cycle); p.setProperty("open", "true"); }
  return { cycle: cycle, open: p.getProperty("open") !== "false" };
}
function monthTag(cycle) { return MONTHS[parseInt(cycle.slice(5), 10) - 1].slice(0, 3) + " " + cycle.slice(0, 4); }
function label(cycle)    { return MONTHS[parseInt(cycle.slice(5), 10) - 1] + " " + cycle.slice(0, 4); }

function rosterSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Roster");
  if (!sh) {
    sh = ss.insertSheet("Roster", 0);
    sh.getRange(1, 1, 1, 4).setValues([["ID", "Name", "Email", "JobTitle"]])
      .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#216AB1");
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 4, 160);
  }
  return sh;
}
function getRoster() {
  const sh = rosterSheet();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => r[1] && r[2])
    .map(r => ({ id: String(r[0] || ""), name: String(r[1]), email: String(r[2]).toLowerCase().trim(), jt: String(r[3] || "") }));
}

function monthSheet(cycle) {
  const ss = SpreadsheetApp.getActive();
  const name = monthTag(cycle);
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  const nCols = 3 + CATS.length * 2;
  // title + confidentiality strip
  sh.getRange(1, 1, 1, nCols).merge().setValue("KABi — Employee of the Month · Votes · " + label(cycle))
    .setBackground("#216AB1").setFontColor("#FFFFFF").setFontSize(13).setFontWeight("bold").setVerticalAlignment("middle");
  sh.setRowHeight(1, 30);
  sh.getRange(2, 1, 1, nCols).merge().setValue("IC-restricted — voter-to-choice data. Never share this file.")
    .setBackground("#0EB3AE").setFontColor("#FFFFFF").setFontSize(9);
  // headers: meta + per category (selection, why)
  const heads = ["Voter Name", "Voter Email", "Submitted At"];
  CATS.forEach(c => { heads.push(c.name); heads.push("Why — " + c.name); });
  const hr = sh.getRange(3, 1, 1, nCols).setValues([heads]).setFontWeight("bold").setWrap(true).setVerticalAlignment("middle");
  sh.setRowHeight(3, 34);
  sh.getRange(3, 1, 1, 3).setBackground("#216AB1").setFontColor("#FFFFFF");
  CATS.forEach((c, i) => {
    sh.getRange(3, 4 + i * 2, 1, 1).setBackground(c.color).setFontColor("#FFFFFF");
    sh.getRange(3, 5 + i * 2, 1, 1).setBackground(c.tint).setFontColor("#17293B");
  });
  sh.setFrozenRows(3);
  sh.setColumnWidth(1, 150).setColumnWidth(2, 190).setColumnWidth(3, 130);
  CATS.forEach((c, i) => { sh.setColumnWidth(4 + i * 2, 160); sh.setColumnWidth(5 + i * 2, 240); });
  return sh;
}

/* ---------------- reads (page + console) ---------------- */
function getUrl() { return ScriptApp.getService().getUrl(); }

function doGet(e) {
  // No ?action → serve the voting page itself. With ?action → the JSON API below.
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile("Index")
      .setTitle("KABi Recognition")
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const a = (e.parameter.action || "state").toLowerCase();
  if (a === "state")  { const st = getState(); return out({ ok: true, cycle: st.cycle, open: st.open }); }
  if (a === "auth")   { return out(e.parameter.code === SECRET ? { ok: true } : { ok: false, error: "unauthorized" }); }
  if (a === "roster") { return out({ ok: true, roster: getRoster() }); }
  if (a === "votes")  {
    if (e.parameter.code !== SECRET) return out({ ok: false, error: "unauthorized" });
    const cycle = e.parameter.cycle || getState().cycle;
    const sh = SpreadsheetApp.getActive().getSheetByName(monthTag(cycle));
    const votes = [];
    if (sh && sh.getLastRow() > 3) {
      const rows = sh.getRange(4, 1, sh.getLastRow() - 3, 3 + CATS.length * 2).getValues();
      const roster = getRoster();
      rows.forEach(r => {
        const sel = {};
        CATS.forEach((c, i) => {
          const nm = String(r[3 + i * 2] || "");
          const p = roster.find(x => x.name === nm);
          sel[c.code] = { id: p ? p.id : nm, name: nm, why: String(r[4 + i * 2] || "") };
        });
        votes.push({ cycle: cycle, voterName: String(r[0]), voterEmail: String(r[1]),
                     submittedAt: String(r[2]).replace(" ", "T"), selections: sel });
      });
    }
    return out({ ok: true, votes: votes });
  }
  return out({ ok: false, error: "unknown action" });
}

/* ---------------- writes (votes + IC actions) ---------------- */
function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: "bad json" }); }
  const action = (b.action || "vote").toLowerCase();

  if (action !== "vote") {
    if (b.code !== SECRET) return out({ ok: false, error: "unauthorized" });
    if (action === "setstate") {
      if (b.cycle) props().setProperty("cycle", String(b.cycle));
      if (typeof b.open !== "undefined") props().setProperty("open", b.open ? "true" : "false");
      return out({ ok: true });
    }
    if (action === "addemp") {
      const emp = b.emp || {};
      const email = String(emp.email || "").toLowerCase().trim();
      if (!email.endsWith(DOMAIN)) return out({ ok: false, error: "email must be " + DOMAIN });
      if (getRoster().some(p => p.email === email)) return out({ ok: false, error: "duplicate email" });
      rosterSheet().appendRow([emp.id || "", emp.name || "", email, emp.jt || ""]);
      return out({ ok: true });
    }
    if (action === "bulkadd") {
      const emps = Array.isArray(b.emps) ? b.emps : [];
      const existing = getRoster(); const seen = {}; existing.forEach(p => seen[p.email] = 1);
      let maxN = 0; existing.forEach(p => { const n = parseInt(String(p.id).slice(3), 10); if (n > maxN) maxN = n; });
      const rows = []; const skipped = [];
      emps.forEach(emp => {
        const email = String(emp.email || "").toLowerCase().trim();
        if (!email.endsWith(DOMAIN)) { skipped.push(email + " — not " + DOMAIN); return; }
        if (seen[email]) { skipped.push(email + " — duplicate"); return; }
        seen[email] = 1; maxN++;
        rows.push(["EMP" + ("000" + maxN).slice(-3), String(emp.name || ""), email, String(emp.jt || "")]);
      });
      if (rows.length) {
        const sh = rosterSheet();
        sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
      }
      return out({ ok: true, added: rows.length, skipped: skipped });
    }
    if (action === "delemp") {
      const sh = rosterSheet();
      const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(b.id)) { sh.deleteRow(i + 2); return out({ ok: true }); }
      }
      return out({ ok: false, error: "not found" });
    }
    return out({ ok: false, error: "unknown action" });
  }

  /* ----- a vote ----- */
  const st = getState();
  if (!st.open) return out({ ok: false, error: "closed" });
  if (b.cycle && b.cycle !== st.cycle) return out({ ok: false, error: "closed" });

  const email = String(b.voterEmail || "").toLowerCase().trim();
  if (!email.endsWith(DOMAIN)) return out({ ok: false, error: "ineligible" });
  const roster = getRoster();
  const voter = roster.find(p => p.email === email);
  if (roster.length && !voter) return out({ ok: false, error: "ineligible" });

  const sel = b.selections || {};
  for (const c of CATS) {
    const x = sel[c.code];
    if (!x || !x.name) return out({ ok: false, error: "missing selection" });
    if (voter && (x.id === voter.id || String(x.name).toLowerCase() === voter.name.toLowerCase()))
      return out({ ok: false, error: "self-vote" });
  }
  // four values → four different colleagues
  const vals = ["BUILD", "INNOVATE", "KABITAL", "GROW"].map(k => sel[k].id + "|" + sel[k].name);
  if (new Set(vals).size !== 4) return out({ ok: false, error: "values must be different colleagues" });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = monthSheet(st.cycle);
    if (sh.getLastRow() > 3) {
      const emails = sh.getRange(4, 2, sh.getLastRow() - 3, 1).getValues().map(r => String(r[0]).toLowerCase());
      if (emails.indexOf(email) !== -1) return out({ ok: false, error: "duplicate" });
    }
    const row = [String(b.voterName || ""), email, Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm")];
    CATS.forEach(c => { row.push(sel[c.code].name); row.push(String(sel[c.code].why || "")); });
    sh.appendRow(row);
    const r = sh.getLastRow();
    if ((r - 4) % 2 === 1) sh.getRange(r, 1, 1, row.length).setBackground("#EAF6FB");
    sh.getRange(r, 1, 1, row.length).setWrap(true).setVerticalAlignment("top");
  } finally { lock.releaseLock(); }
  return out({ ok: true });
}
