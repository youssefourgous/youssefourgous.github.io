/*
  Class Representative Voting Web App
  Google Apps Script server-side code
*/

const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const STUDENTS_SHEET = 'Students';
const VOTES_SHEET = 'Votes';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Voting Admin')
    .addItem('Initialize Sheets', 'initSheets')
    .addItem('Open Admin Web App', 'openAdminUrl')
    .addToUi();
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'index';
  if (page === 'admin') {
    return HtmlService.createTemplateFromFile('Admin')
      .evaluate()
      .setTitle('Class Rep Voting - Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Class Rep Voting')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function initSheets() {
  const ss = getSpreadsheet();
  let students = ss.getSheetByName(STUDENTS_SHEET);
  if (!students) {
    students = ss.insertSheet(STUDENTS_SHEET);
    students.getRange(1, 1, 1, 4).setValues([[
      'national_id', 'first_name', 'last_name', 'class'
    ]]);
  }

  let votes = ss.getSheetByName(VOTES_SHEET);
  if (!votes) {
    votes = ss.insertSheet(VOTES_SHEET);
    votes.getRange(1, 1, 1, 3).setValues([[
      'timestamp', 'student_id', 'candidate'
    ]]);
  }
}

function getSpreadsheet() {
  if (!SPREADSHEET_ID) {
    throw new Error('SPREADSHEET_ID is not set. Set it in Script Properties.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getStudentData(studentId) {
  if (!studentId) return null;
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName(STUDENTS_SHEET);
  if (!sh) return null;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null; // no data

  const range = sh.getRange(2, 1, lastRow - 1, 4);
  const values = range.getValues();
  const idxNational = 0, idxFirst = 1, idxLast = 2, idxClass = 3;
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (String(row[idxNational]).trim() === String(studentId).trim()) {
      return {
        national_id: String(row[idxNational]).trim(),
        first_name: String(row[idxFirst]).trim(),
        last_name: String(row[idxLast]).trim(),
        class: String(row[idxClass]).trim(),
      };
    }
  }
  return null;
}

function hasVoted(studentId) {
  if (!studentId) return false;
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName(VOTES_SHEET);
  if (!sh) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const range = sh.getRange(2, 1, lastRow - 1, 3);
  const values = range.getValues();
  const idxStudent = 1;
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (String(row[idxStudent]).trim() === String(studentId).trim()) {
      return true;
    }
  }
  return false;
}

function submitVote(studentId, candidate) {
  if (!studentId || !candidate) {
    return { success: false, error: 'Missing student or candidate' };
  }

  const student = getStudentData(studentId);
  if (!student || student.class !== 'GM2') {
    return { success: false, error: 'Student not in class GM2' };
  }

  if (hasVoted(studentId)) {
    return { success: false, alreadyVoted: true };
  }

  const ss = getSpreadsheet();
  const sh = ss.getSheetByName(VOTES_SHEET);
  if (!sh) throw new Error('Votes sheet missing');

  sh.appendRow([new Date(), String(studentId).trim(), String(candidate).trim()]);
  return { success: true };
}

function getResults() {
  const ss = getSpreadsheet();
  const sh = ss.getSheetByName(VOTES_SHEET);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const range = sh.getRange(2, 1, lastRow - 1, 3);
  const values = range.getValues();
  const counts = {};
  const idxCandidate = 2;
  for (let i = 0; i < values.length; i++) {
    const candidate = String(values[i][idxCandidate]).trim();
    if (!candidate) continue;
    counts[candidate] = (counts[candidate] || 0) + 1;
  }
  return Object.keys(counts).sort().map(name => ({ candidate: name, votes: counts[name] }));
}

function openAdminUrl() {
  const url = ScriptApp.getService().getUrl() + '?page=admin';
  const ui = SpreadsheetApp.getUi();
  ui.alert('Admin Web App URL', url, ui.ButtonSet.OK);
}
