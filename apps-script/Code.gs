/**
 * Семинарға «Келдім» белгілерін Google кестеге жазатын скрипт.
 * Орнату: README.md, 3-қадам.
 */
const SHEET_NAME = "Келгендер";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Уақыты", "№", "Аты-жөні", "Мектеп", "Тізімде бар ма", "ID"]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || "{}");
    const name = String(d.name || "").slice(0, 120);
    if (!name) return json_({ ok: false, error: "name" });
    const sh = getSheet_();
    const id = String(d.id || "");
    // Бір адам екі рет басса, қайта жазбаймыз
    if (id) {
      const ids = sh.getRange(2, 6, Math.max(sh.getLastRow() - 1, 1), 1).getValues().flat();
      if (ids.indexOf(id) !== -1) return json_({ ok: true, duplicate: true });
    }
    sh.appendRow([
      new Date(),
      d.n || "",
      name,
      String(d.school || "").slice(0, 160),
      d.listed ? "иә" : "жоқ",
      id
    ]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, count: Math.max(getSheet_().getLastRow() - 1, 0) });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
