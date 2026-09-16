# KABi EOM — النشر الأسهل: كل شيء داخل Google Apps Script (بدون GitHub)

رابط واحد = صفحة التصويت + المستقبِل + جدول الأصوات (ورقة لكل شهر).

## النشر (٥ دقائق)
1. افتح `sheets.new` → **Extensions → Apps Script**.
2. الصق `Code.gs` مكان الملف الافتراضي، وعدّل `SECRET` (رمز كونسول IC).
3. من زر **+** → **HTML** → سمِّه **Index** بالضبط → الصق محتوى `Index.html`.
4. **Deploy → New deployment → Web app**: `Execute as: Me` · `Who has access: Anyone`.
5. **رابط الـ Web app هو رابط التصويت** — أرسله للموظفين. لا يوجد أي إعداد آخر: الصفحة تكتشف رابطها بنفسها.

## بعدها
- عبّئ ورقة **Roster** (تُنشأ تلقائيًا) أو ارفع القالب من كونسول IC.
- كل التحديثات مستقبلًا: الصق الجديد → **Deploy → Manage deployments → Edit → New version** (نفس الرابط يبقى).
- الإكسل الشهري: `File → Download → Microsoft Excel` — كل شهر ورقته، أو زر «Excel — this month» من الكونسول.

## ملاحظتان بصراحة
- أعلى الصفحة يظهر شريط Google رفيع ("This application was created by another user…") — تجميلي فقط ولا يمكن إزالته.
- الرابط طويل الشكل؛ لا يهم عمليًا لأن الموظف يصل عبر زر **Vote Now** في إيميل IC.
