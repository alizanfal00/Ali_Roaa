/**
 * استقبال تأكيدات الحضور ورسائل التهنئة من صفحة الدعوة
 * - بيسجّلهم في Google Sheet (تبويبين: RSVP و Wishes)
 * - وبيبعتلك إيميل مع كل رد
 *
 * الخطوات في ملف DEPLOY.txt
 */
const OWNER_EMAIL = "PUT_YOUR_EMAIL@gmail.com";   // ← إيميلك

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const name = clean(d.name, 80);

    if (d.type === "rsvp") {
      const attending = d.attending === "yes";
      const guests = Math.max(1, Math.min(10, parseInt(d.guests, 10) || 1));
      getSheet(ss, "RSVP", ["Time", "Name", "Attending", "Guests"])
        .appendRow([now, name, attending ? "Yes ✅" : "No 🙏", guests]);
      MailApp.sendEmail(
        OWNER_EMAIL,
        (attending ? "✅ تأكيد حضور: " : "🙏 اعتذار: ") + name,
        name + (attending ? " هيحضر" : " اعتذر عن الحضور") + (attending ? "\nعدد الحضور: " + guests : "")
      );
    } else if (d.type === "wish") {
      const message = clean(d.message, 1000);
      getSheet(ss, "Wishes", ["Time", "Name", "Message"]).appendRow([now, name, message]);
      MailApp.sendEmail(OWNER_EMAIL, "💌 تهنئة جديدة من " + name, name + ":\n\n" + message);
    }
    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("error: " + err);
  }
}

function getSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); }
  return sh;
}

// تنظيف النص: تحديد الطول ومنع الصيغ (=, +, -, @) من إنها تتنفذ في الشيت
function clean(v, max) {
  let s = String(v == null ? "" : v).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
