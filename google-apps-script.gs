/**
 * Backend for the AI Impact Tracker's shared branch data (logs + award
 * nominations). Paste this into Extensions > Apps Script on the Google
 * Sheet, then deploy it as a Web App.
 *
 * Award evidence files are saved into a Drive folder (created on first
 * use) and the Sheet's "Files" cell stores a JSON list of {name, url}.
 */

const EVIDENCE_FOLDER_NAME = "AI Impact Tracker Evidence";
const PROJECT_NAMES = ["Bath Mind", "Julian House", "Bath City Farm", "Genesis Trust", "Dorothy House Hospice", "Southside Family Project"];

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logs = readSheet(ss.getSheetByName("Logs"));
  const submissions = readSheet(ss.getSheetByName("Submissions"));
  const archive = readSheet(ss.getSheetByName("Archive"));
  const tools = readSheet(ss.getSheetByName("Tools"));
  const winners = readSheet(ss.getSheetByName("Winners"));
  const governance = readGovernance(ss);
  return ContentService.createTextOutput(JSON.stringify({ logs, submissions, archive, tools, winners, governance }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.type === "saveTools") {
    return json(saveTools(data.fields || {}));
  }
  if (data.type === "deleteLog") {
    return json(deleteLog(data.fields || {}));
  }
  if (data.type === "vote") {
    return json(changeVotes(data.fields || {}));
  }
  if (data.type === "award") {
    return json(awardSubmission(data.fields || {}));
  }
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

/**
 * award: moves a nomination from Submissions to the Winners tab (created
 * automatically) and stamps the date it was awarded. Files stay in Drive.
 * "Why it won" is an optional note typed by the exec who pressed Award.
 */
function awardSubmission(m) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Submissions");
  if (!sheet) return { ok: true, found: false };
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const col = h => headers.indexOf(h);
  const same = (row, h) => String(row[col(h)]) === String(m[h] === undefined ? "" : m[h]);
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!(same(row, "Name") && same(row, "Prompt") && same(row, "Achievement"))) continue;
    let winners = ss.getSheetByName("Winners");
    if (!winners) {
      winners = ss.insertSheet("Winners");
      winners.appendRow(headers.concat(["Awarded at", "Why it won"]));
    }
    const wHeaders = winners.getRange(1, 1, 1, winners.getLastColumn()).getValues()[0];
    winners.appendRow(wHeaders.map(h => {
      if (h === "Awarded at") return new Date();
      if (h === "Why it won") return m.Why || "";
      return col(h) >= 0 ? row[col(h)] : "";
    }));
    sheet.deleteRow(i + 1);
    return { ok: true, found: true };
  }
  return { ok: true, found: false };
}

/**
 * Adds or removes one upvote (Delta = 1 or -1) on the matching Submissions
 * row. Needs a "Votes" column in the Submissions header row; an empty Votes
 * cell counts as 1 (the nominee's own vote).
 */
function changeVotes(m) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Submissions");
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const col = h => headers.indexOf(h);
    const vc = col("Votes");
    if (vc < 0) return { ok: true, found: false, note: "No Votes column" };
    const same = (row, h) => String(row[col(h)]) === String(m[h] === undefined ? "" : m[h]);
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!(same(row, "Name") && same(row, "Prompt") && same(row, "Achievement"))) continue;
      const current = row[vc] === "" ? 1 : Number(row[vc]) || 0;
      sheet.getRange(i + 1, vc + 1).setValue(Math.max(0, current + (Number(m.Delta) || 0)));
      return { ok: true, found: true };
    }
    return { ok: true, found: false };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes one row from the Logs tab. Matched on Name, Use case, Project,
 * Tool, Task and Minutes (the first identical row is removed).
 */
function deleteLog(m) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  if (!sheet) return { ok: true, found: false };
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const cell = (row, h) => { const c = headers.indexOf(h); return c < 0 ? "" : String(row[c]); };
  const keys = ["Name", "Use case", "Project", "Tool", "Task", "Minutes"];
  for (let i = 1; i < values.length; i++) {
    if (keys.every(k => cell(values[i], k) === String(m[k] === undefined ? "" : m[k]))) {
      sheet.deleteRow(i + 1);
      return { ok: true, found: true };
    }
  }
  return { ok: true, found: false };
}

/**
 * The Governance tab is created automatically. Execs can edit it:
 *   A: Project, B: Last check-in (a date or any text)
 *   D2: Cycle start (a date; leave blank to count all logs)
 *   E2: Cycle length in weeks (default 10)
 */
function readGovernance(ss) {
  let sheet = ss.getSheetByName("Governance");
  if (!sheet) {
    sheet = ss.insertSheet("Governance");
    sheet.getRange(1, 1, 1, 2).setValues([["Project", "Last check-in"]]);
    sheet.getRange(2, 1, PROJECT_NAMES.length, 1).setValues(PROJECT_NAMES.map(p => [p]));
    sheet.getRange(1, 4, 1, 2).setValues([["Cycle start", "Cycle length (weeks)"]]);
    sheet.getRange(2, 5).setValue(10);
  }
  const rows = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  const checkins = {};
  rows.forEach(r => { if (r[0]) checkins[r[0]] = r[1]; });
  return {
    checkins: checkins,
    cycleStart: sheet.getRange(2, 4).getValue(),
    cycleWeeks: Number(sheet.getRange(2, 5).getValue()) || 10
  };
}

/** Saves which AI tools a project says it uses (one row per project on the Tools tab). */
function saveTools(m) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Tools");
    if (!sheet) {
      sheet = ss.insertSheet("Tools");
      sheet.appendRow(["Project", "Tools", "Updated"]);
    }
    const values = sheet.getDataRange().getValues();
    const row = [m.Project, JSON.stringify(m.Tools || []), new Date()];
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(m.Project)) {
        sheet.getRange(i + 1, 1, 1, 3).setValues([row]);
        return { ok: true };
      }
    }
    sheet.appendRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
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
