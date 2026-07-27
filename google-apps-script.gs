/**
 * AI-Mind Care Breathing Flower V1.2.1
 * Minimal Google Sheets receiver
 *
 * Stored fields:
 * nickname, completion date/time, practice duration, flower reward
 */

const SHEET_NAME = "Breathing Flower Records";

function doGet() {
  return ContentService
    .createTextOutput("AI-Mind Care Breathing Flower V1.2.1 endpoint is active.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    const headers = [
      "Server Timestamp",
      "Record ID",
      "Nickname",
      "Completion Date and Time",
      "Practice Minutes",
      "Flower",
      "Source"
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }

    const p = e && e.parameter ? e.parameter : {};
    const recordId = clean_(p.recordId, 80);

    if (recordId && sheet.getLastRow() > 1) {
      const ids = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
      if (ids.includes(recordId)) return json_({ok:true, duplicate:true});
    }

    sheet.appendRow([
      new Date(),
      recordId,
      clean_(p.nickname, 30),
      clean_(p.completedAt, 60),
      number_(p.durationMinutes),
      clean_(p.flower, 40),
      clean_(p.source, 100)
    ]);

    return json_({ok:true});
  } catch (error) {
    return json_({ok:false, error:String(error)});
  } finally {
    lock.releaseLock();
  }
}

function clean_(value, maxLength) {
  return String(value == null ? "" : value).replace(/[<>{}]/g, "").slice(0, maxLength);
}

function number_(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : "";
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
