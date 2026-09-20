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
  const archive = readSheet(ss.getSheetByName("Archive"));
  return ContentService.createTextOutput(JSON.stringify({ logs, submissions, archive }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.type === "archive" || data.type === "delete" || data.type === "restore" || data.type === "purge") {
    return json(removeSubmission(data.type, data.fields || {}));
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = data.type === "log" ? "Logs" : "Submissions";
  const sheet = ss.getSheetByName(sheetName);
  const fields = data.fields || {};

  if (data.type === "submission" && data.attachments && data.attachments.length) {
    try {
      fields.Files = JSON.stringify(saveAttachments(data.attachments));
    } catch (err) {
      // Never lose the nomination because a file could not be saved.
      fields.Files = "File not saved: " + err.message;
    }
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

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * archive: Submissions -> Archive tab (files stay in Drive).
 * delete:  removes from Submissions and moves its Drive files to the bin.
 * restore: Archive tab -> back to Submissions.
 * purge:   removes from the Archive tab and moves its Drive files to the bin.
 * The row is matched on Name + Prompt + Achievement.
 */
function removeSubmission(action, match) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fromArchive = action === "restore" || action === "purge";
  const sheet = ss.getSheetByName(fromArchive ? "Archive" : "Submissions");
  if (!sheet) return { ok: true, found: false };
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const col = h => headers.indexOf(h);
  const same = (row, h) => String(row[col(h)]) === String(match[h] === undefined ? "" : match[h]);

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!(same(row, "Name") && same(row, "Prompt") && same(row, "Achievement"))) continue;

    if (action === "archive") {
      let archive = ss.getSheetByName("Archive");
      if (!archive) {
        archive = ss.insertSheet("Archive");
        archive.appendRow(headers.concat(["Archived at"]));
      }
      archive.appendRow(row.concat([new Date()]));
    } else if (action === "restore") {
      const main = ss.getSheetByName("Submissions");
      const mainHeaders = main.getRange(1, 1, 1, main.getLastColumn()).getValues()[0];
      main.appendRow(mainHeaders.map(h => (col(h) >= 0 ? row[col(h)] : "")));
    } else {
      trashFiles(row[col("Files")]);
    }
    sheet.deleteRow(i + 1);
    return { ok: true, found: true };
  }
  return { ok: true, found: false };
}

function trashFiles(cell) {
  try {
    JSON.parse(cell).forEach(f => {
      const m = String(f.url || "").match(/\/d\/([\w-]+)/);
      if (m) DriveApp.getFileById(m[1]).setTrashed(true);
    });
  } catch (e) {}
}

function saveAttachments(attachments) {
  const folders = DriveApp.getFoldersByName(EVIDENCE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(EVIDENCE_FOLDER_NAME);
  return attachments.slice(0, 5).map(a => {
    const blob = Utilities.newBlob(Utilities.base64Decode(a.data), a.mime, a.name);
    const file = folder.createFile(blob);
    // Prefer "anyone with the link". If your Workspace blocks that, fall back
    // to "anyone in your organisation with the link".
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e1) {
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {}
    }
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

// Run this once from the editor to grant full Drive access, then redeploy.
function authorizeDrive() {
  const folders = DriveApp.getFoldersByName(EVIDENCE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(EVIDENCE_FOLDER_NAME);
  Logger.log("Drive access is authorized. Folder: " + folder.getUrl());
}
