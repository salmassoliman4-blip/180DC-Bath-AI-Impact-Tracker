/**
 * Backend for the AI Impact Tracker's shared branch data (logs + award
 * nominations). Paste this into Extensions > Apps Script on the Google
 * Sheet, then deploy it as a Web App.
 *
 * Award evidence files are saved into a Drive folder (created on first
 * use) and the Sheet's "Files" cell stores a JSON list of {name, url}.
 */

const EVIDENCE_FOLDER_NAME = "AI Impact Tracker Evidence";
const PROJECT_NAMES = ["CCCYD", "The Alfred Gillett Trust", "KidsOut", "Motivation Africa", "TAFU", "I Can and I Am"];

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const view = e && e.parameter ? e.parameter.view : "";
  if (view === "accounts") return json({ accounts: readAccounts(ss) });
  if (view === "reset") return json(checkReset(ss, e.parameter.token));
  if (view === "votes") return json({ keys: votesBy(ss, e.parameter.voter) });
  const logs = readSheet(ss.getSheetByName("Logs"));
  const submissions = readSheet(ss.getSheetByName("Submissions"));
  const archive = readSheet(ss.getSheetByName("Archive"));
  const tools = readSheet(ss.getSheetByName("Tools"));
  const winners = readSheet(ss.getSheetByName("Winners"));
  const governance = readGovernance(ss);
  const accounts = readAccounts(ss);
  return ContentService.createTextOutput(JSON.stringify({ logs, submissions, archive, tools, winners, governance, accounts }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.type === "setPassword") return json(setPassword(data.fields || {}));
  if (data.type === "recordVote") return json(recordVote(data.fields || {}));
  if (data.type === "failedAttempt") return json(changeAttempts(data.fields || {}, 1));
  if (data.type === "clearAttempts") return json(changeAttempts(data.fields || {}, 0));
  if (data.type === "createReset") return json(createReset(data.fields || {}));
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

/**
 * Accounts tab (created automatically): one row per person who has set their
 * own password. The password itself is never stored, only a salted SHA-256
 * hash computed in the browser. Columns:
 *   A Username | B Salt | C Hash | D Updated | E Failed attempts | F Reset token | G Reset expires
 * Delete a person's row to put them back on their default password.
 * How it stays safe: the browser sends a "proof" (SHA-256 of salt + password)
 * and this script stores SHA-256 of that proof. Reading the sheet therefore
 * does not give anyone the proof they would need to change a password, and
 * only the people in RESET_ADMINS can create reset links.
 * Note: this web app has no sign-in of its own, so anyone with its URL can
 * still read the hashes and try to guess short passwords offline. Keep the
 * URL inside the branch.
 */
const RESET_ADMINS = ["salma.ibrahim", "jashan.ball"];

function sha256Hex(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b < 0 ? b + 256 : b)).toString(16).padStart(2, "0")).join("");
}
const ACCOUNT_HEADERS = ["Username", "Salt", "Hash", "Updated", "Failed attempts", "Reset token", "Reset expires"];

function accountsSheet(ss) {
  let sheet = ss.getSheetByName("Accounts");
  if (!sheet) {
    sheet = ss.insertSheet("Accounts");
    sheet.appendRow(ACCOUNT_HEADERS);
    // Keep long hex strings and tokens as plain text so Sheets never turns
    // one into a number (for example 12e45... becoming scientific notation).
    sheet.getRange("A:C").setNumberFormat("@");
    sheet.getRange("F:F").setNumberFormat("@");
  }
  return sheet;
}

function readAccounts(ss) {
  const values = accountsSheet(ss).getDataRange().getValues();
  const out = {};
  for (let i = 1; i < values.length; i++) {
    const u = String(values[i][0] || "");
    if (!u) continue;
    out[u] = {
      salt: String(values[i][1] || ""),
      hash: String(values[i][2] || ""),
      attempts: Number(values[i][4]) || 0
    };
  }
  return out;
}

/** Reads one account row, lets mutate(row, existed) change it, writes it back (or appends it). */
function updateAccount(username, mutate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = accountsSheet(SpreadsheetApp.getActiveSpreadsheet());
    const values = sheet.getDataRange().getValues();
    let rowNum = 0, cur = [username, "", "", "", 0, "", ""];
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === username) { rowNum = i + 1; cur = values[i].slice(0, 7); break; }
    }
    const next = mutate(cur, rowNum > 0);
    if (!next) return { ok: false };
    if (rowNum) {
      sheet.getRange(rowNum, 1, 1, 7).setValues([next]);
    } else {
      sheet.appendRow(next);
      sheet.getRange(sheet.getLastRow(), 1, 1, 3).setNumberFormat("@");
      sheet.getRange(sheet.getLastRow(), 6).setNumberFormat("@");
      sheet.getRange(sheet.getLastRow(), 1, 1, 7).setValues([next]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function cleanUsername(v) {
  const u = String(v || "").trim().toLowerCase();
  return u && u.length <= 80 ? u : "";
}

/**
 * Sets (or changes) a password. Hash is SHA-256 of the proof the browser
 * computed. Allowed when:
 *   - the person has no custom password yet, or
 *   - SHA-256(PrevProof) matches the stored hash (they knew the old password), or
 *   - Token matches a reset token that has not expired.
 */
function setPassword(m) {
  const username = cleanUsername(m.Username);
  const salt = String(m.Salt || ""), hash = String(m.Hash || "");
  if (!username || !/^[0-9a-f]{16,64}$/.test(salt) || !/^[0-9a-f]{64}$/.test(hash)) {
    return { ok: false, reason: "invalid" };
  }
  let reason = "";
  const res = updateAccount(username, cur => {
    const storedHash = String(cur[2] || ""), token = String(cur[5] || ""), expires = Number(cur[6]) || 0;
    const tokenOk = m.Token && token && String(m.Token) === token && Date.now() < expires;
    const prevOk = !storedHash || (m.PrevProof && sha256Hex(String(m.PrevProof)) === storedHash);
    if (!tokenOk && !prevOk) { reason = "changed"; return null; }
    cur[1] = salt; cur[2] = hash; cur[3] = new Date(); cur[4] = 0;
    if (tokenOk) { cur[5] = ""; cur[6] = ""; }
    return cur;
  });
  return res.ok ? res : { ok: false, reason: reason || "failed" };
}

/** delta 1 = one more failed attempt, 0 = clear the counter (unlock). */
function changeAttempts(m, delta) {
  const username = cleanUsername(m.Username);
  if (!username) return { ok: false };
  return updateAccount(username, (cur, existed) => {
    if (!delta && !existed) return null;   // nothing to clear, so don't create an empty row
    cur[4] = delta ? (Number(cur[4]) || 0) + 1 : 0;
    return cur;
  });
}

/**
 * Stores a one-time reset token (valid until Expires, a millisecond timestamp).
 * Only a person in RESET_ADMINS who proves their own password can do this.
 */
function createReset(m) {
  const username = cleanUsername(m.Username);
  const token = String(m.Token || "");
  if (!username || !/^[A-Za-z0-9]{8,64}$/.test(token)) return { ok: false, reason: "invalid" };
  const admin = cleanUsername(m.AdminUser);
  if (RESET_ADMINS.indexOf(admin) < 0) return { ok: false, reason: "not allowed" };
  const accounts = readAccounts(SpreadsheetApp.getActiveSpreadsheet());
  const stored = accounts[admin] && accounts[admin].hash;
  if (!stored || !m.AdminProof || sha256Hex(String(m.AdminProof)) !== stored) return { ok: false, reason: "not allowed" };
  const expires = Number(m.Expires) || (Date.now() + 24 * 60 * 60 * 1000);
  return updateAccount(username, cur => { cur[5] = token; cur[6] = expires; return cur; });
}

/** Used by a reset link: returns the username if the token is valid and unexpired. */
function checkReset(ss, token) {
  token = String(token || "");
  if (!token) return {};
  const values = accountsSheet(ss).getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][5]) === token && Date.now() < (Number(values[i][6]) || 0)) {
      return { username: String(values[i][0]) };
    }
  }
  return {};
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
 * VoteLog tab (created automatically): one row per person per nomination they
 * upvoted, so "I've already voted" follows the person across browsers.
 *   A Voter | B Key | C When
 */
function voteLogSheet(ss) {
  let sheet = ss.getSheetByName("VoteLog");
  if (!sheet) {
    sheet = ss.insertSheet("VoteLog");
    sheet.appendRow(["Voter", "Key", "When"]);
  }
  return sheet;
}

/** The nomination keys this person has upvoted. */
function votesBy(ss, voter) {
  const v = cleanUsername(voter);
  if (!v) return [];
  const values = voteLogSheet(ss).getDataRange().getValues();
  const keys = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === v) keys.push(String(values[i][1]));
  }
  return keys;
}

/** Row number (1-based) of this person's vote for this nomination, or 0. */
function findVoteRow(sheet, voter, key) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === voter && String(values[i][1]) === key) return i + 1;
  }
  return 0;
}

/**
 * Notes that a person voted for a nomination WITHOUT changing its count. Used
 * for a person's own vote on their nomination, and to move votes counted
 * before votes were tracked per person.
 */
function recordVote(m) {
  const voter = cleanUsername(m.Voter), key = String(m.Key || "");
  if (!voter || !key) return { ok: false };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = voteLogSheet(SpreadsheetApp.getActiveSpreadsheet());
    if (!findVoteRow(sheet, voter, key)) sheet.appendRow([voter, key, new Date()]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Adds or removes one upvote (Delta = 1 or -1) on the matching Submissions
 * row. Needs a "Votes" column in the Submissions header row; an empty Votes
 * cell counts as 1 (the nominee's own vote).
 * When Voter and Key are sent, each person can only add one vote per
 * nomination and can only remove a vote they actually cast, so repeating the
 * request (or voting from a second browser) never double counts.
 */
function changeVotes(m) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Submissions");
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const col = h => headers.indexOf(h);
    const vc = col("Votes");
    if (vc < 0) return { ok: true, found: false, note: "No Votes column" };
    const same = (row, h) => String(row[col(h)]) === String(m[h] === undefined ? "" : m[h]);
    const delta = Number(m.Delta) || 0;
    const voter = cleanUsername(m.Voter), key = String(m.Key || "");
    let log = null, logRow = 0;
    if (voter && key) {
      log = voteLogSheet(ss);
      logRow = findVoteRow(log, voter, key);
      if (delta > 0 && logRow) return { ok: true, found: true, unchanged: true };
      if (delta < 0 && !logRow) return { ok: true, found: true, unchanged: true };
    }
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!(same(row, "Name") && same(row, "Prompt") && same(row, "Achievement"))) continue;
      const current = row[vc] === "" ? 1 : Number(row[vc]) || 0;
      sheet.getRange(i + 1, vc + 1).setValue(Math.max(0, current + delta));
      if (log) {
        if (delta > 0) log.appendRow([voter, key, new Date()]);
        else if (logRow) log.deleteRow(logRow);
      }
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
