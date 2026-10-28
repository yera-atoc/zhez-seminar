/**
 * Семинарға «Келдім» белгілерін Google кестеге жазатын скрипт.
 *
 * Кестенің БІРІНШІ парағы — мұғалімдер тізімі:
 *   A: №   B: Аты-жөні   C: Келді   D: Уақыты   E: ID
 * Мұғалім «Келдім» басқанда, оның жолында C және D бағандары толтырылады.
 * Тізімде жоқ адамдар «Тізімде жоқтар» парағына жазылады.
 */
const EXTRA_SHEET = "Тізімде жоқтар";
const TZ = "Asia/Almaty";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const name = String(d.name || "").slice(0, 120).trim();
    if (!name) return json_({ ok: false, error: "name" });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const time = Utilities.formatDate(new Date(), TZ, "dd.MM.yyyy HH:mm");
    const id = String(d.id || "").trim();

    if (id) {
      const roster = ss.getSheets()[0];
      const last = roster.getLastRow();
      if (last >= 2) {
        const ids = roster.getRange(2, 5, last - 1, 1).getValues().flat().map(String);
        const i = ids.indexOf(id);
        if (i !== -1) {
          const row = i + 2;
          if (!roster.getRange(row, 3).getValue()) {
            roster.getRange(row, 3, 1, 2).setValues([["✅ келді", time]]);
            roster.getRange(row, 1, 1, 5).setBackground("#e3f1e9");
          }
          return json_({ ok: true });
        }
      }
    }

    let extra = ss.getSheetByName(EXTRA_SHEET);
    if (!extra) {
      extra = ss.insertSheet(EXTRA_SHEET);
      extra.appendRow(["Уақыты", "Аты-жөні", "Мектеп"]);
      extra.setFrozenRows(1);
      extra.getRange(1, 1, 1, 3).setFontWeight("bold");
    }
    extra.appendRow([time, name, String(d.school || "").slice(0, 160)]);
    return json_({ ok: true, extra: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const roster = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const last = roster.getLastRow();
  const came = last >= 2 ? roster.getRange(2, 3, last - 1, 1).getValues().flat().filter(String).length : 0;
  return json_({ ok: true, came: came, total: Math.max(last - 1, 0) });
}

/** Тексеру үшін: редакторда осы функцияны іске қосыңыз — 1-мұғалім «келді» болып белгіленеді. */
function testCheckin() {
  const roster = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const id = String(roster.getRange(2, 5).getValue());
  doPost({ postData: { contents: JSON.stringify({ id: id, name: "тест" }) } });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
