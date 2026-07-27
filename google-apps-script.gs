/**
 * AI-Mind Care Breathing Flower V1.2
 * Google Sheets receiver
 *
 * Recommended use:
 * 1. Create a Google Sheet.
 * 2. Open Extensions > Apps Script.
 * 3. Paste this file into Code.gs.
 * 4. Deploy as a Web App.
 * 5. Put the deployment /exec URL in config.js.
 */

const SHEET_NAME = "Breathing Flower Records";

function doGet() {
  return ContentService
    .createTextOutput("AI-Mind Care Breathing Flower record endpoint is active.")
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
      "Practice Date",
      "Start Time",
      "Completion Time",
      "Duration Minutes",
      "Breathing Cycles",
      "Emotion Before",
      "Emotion After",
      "Relaxation Before",
      "Relaxation After",
      "Source"
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }

    const p = (e && e.parameter) ? e.parameter : {};
    const recordId = clean_(p.recordId, 80);

    // Avoid accidental duplicate submissions with the same record ID.
    if (recordId && sheet.getLastRow() > 1) {
      const ids = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
      if (ids.includes(recordId)) {
        return json_({ok: true, duplicate: true});
      }
    }

    sheet.appendRow([
      new Date(),
      recordId,
      clean_(p.nickname, 30),
      clean_(p.practiceDate, 20),
      clean_(p.startTime, 50),
      clean_(p.completionTime, 50),
      number_(p.durationMinutes),
      number_(p.breathingCycles),
      number_(p.emotionBefore),
      number_(p.emotionAfter),
      number_(p.relaxationBefore),
      number_(p.relaxationAfter),
      clean_(p.source, 100)
    ]);

    return json_({ok: true});
  } catch (error) {
    return json_({ok: false, error: String(error)});
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
