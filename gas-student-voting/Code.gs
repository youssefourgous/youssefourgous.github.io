/*
  Minimal Student Voting Web App (Google Apps Script)
  Sheets:
    Students: national_id | first_name | last_name
    Votes: timestamp | national_id | candidate
*/

const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const STUDENTS_SHEET = 'Students';
const VOTES_SHEET = 'Votes';

function doGet() {
  return HtmlService.createTemplateFromFile('form')
    .evaluate()
    .setTitle('Student Voting')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet_() {
  if (!SPREADSHEET_ID) {
    throw new Error('SPREADSHEET_ID is not set in Script Properties');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getStudentData(studentId) {
  if (!studentId) return null;
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(STUDENTS_SHEET);
  if (!sh) return null;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const rng = sh.getRange(2, 1, lastRow - 1, 3);
  const values = rng.getValues();
  for (let i = 0; i < values.length; i++) {
    const [national_id, first_name, last_name] = values[i].map(v => String(v).trim());
    if (national_id === String(studentId).trim()) {
      return { national_id, first_name, last_name };
    }
  }
  return null;
}

function hasVoted(studentId) {
  if (!studentId) return false;
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(VOTES_SHEET);
  if (!sh) return false;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const rng = sh.getRange(2, 1, lastRow - 1, 3);
  const values = rng.getValues();
  for (let i = 0; i < values.length; i++) {
    const [, id] = values[i];
    if (String(id).trim() === String(studentId).trim()) return true;
  }
  return false;
}

function submitVote(studentId, candidate) {
  if (!studentId || !candidate) {
    return { success: false, error: 'Missing fields' };
  }
  const student = getStudentData(studentId);
  if (!student) {
    return { success: false, error: 'Student not found' };
  }
  if (hasVoted(studentId)) {
    return { success: false, alreadyVoted: true };
  }
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(VOTES_SHEET);
  if (!sh) throw new Error('Votes sheet missing');
  sh.appendRow([new Date(), String(studentId).trim(), String(candidate).trim()]);
  return { success: true };
}

function getResults() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(VOTES_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const rng = sh.getRange(2, 1, lastRow - 1, 3);
  const values = rng.getValues();
  const counts = {};
  for (let i = 0; i < values.length; i++) {
    const candidate = String(values[i][2]).trim();
    if (!candidate) continue;
    counts[candidate] = (counts[candidate] || 0) + 1;
  }
  return Object.keys(counts).sort().map(name => ({ candidate: name, votes: counts[name] }));
}
