/**
 * Backend for the AI Impact Tracker's shared branch data (logs + award
 * nominations). Paste this into Extensions > Apps Script on the Google
 * Sheet described in SETUP.md, then deploy it as a Web App.
 */

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logs = readSheet(ss.getSheetByName("Logs"));
  const submissions = readSheet(ss.getSheetByName("Submissions"));
  return ContentService.createTextOutput(JSON.stringify({ logs, submissions }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = data.type === "log" ? "Logs" : "Submissions";
  const sheet = ss.getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    if (h === "Timestamp") return new Date();
    return data.fields[h] !== undefined ? data.fields[h] : "";
  });
  sheet.appendRow(row);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function readSheet(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}
