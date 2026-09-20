/**
 * Backend for the AI Impact Tracker's shared branch data (logs + award
 * nominations). Paste this into Extensions > Apps Script on the Google
 * Sheet, then deploy it as a Web App.
 *
 * Award evidence files are saved into a Drive folder (created on first
 * use) and the Sheet's "Files" cell stores a JSON list of {name, url}.
 */

const EVIDENCE_FOLDER_NAME = "AI Impact Tracker Evidence";

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
  const fields = data.fields || {};

  if (data.type === "submission" && data.attachments && data.attachments.length) {
    fields.Files = JSON.stringify(saveAttachments(data.attachments));
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    if (h === "Timestamp") return new Date();
    return fields[h] !== undefined ? fields[h] : "";
  });
  sheet.appendRow(row);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveAttachments(attachments) {
  const folders = DriveApp.getFoldersByName(EVIDENCE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(EVIDENCE_FOLDER_NAME);
  return attachments.slice(0, 5).map(a => {
    const blob = Utilities.newBlob(Utilities.base64Decode(a.data), a.mime, a.name);
    const file = folder.createFile(blob);
    // Anyone with the link can open it, no Google account needed. The link
    // is long and unguessable, but treat it like a secret. To restrict to
    // your 180dc.org Workspace instead, use DriveApp.Access.DOMAIN_WITH_LINK.
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { name: a.name, url: file.getUrl() };
  });
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
