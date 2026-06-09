/**
 * RTI / MTSS Kanban Manager — Google Apps Script backend
 * -----------------------------------------------------------------------------
 * Supports ELA + Math on one row per student.
 * Columns are resolved by HEADER NAME, not position, so the master sheet can be
 * reordered or extended without breaking the app.
 *
 * One-time setup:  run setupSheets() from the editor, then deploy as a Web App
 * (Deploy > New deployment > Web app).
 *   - "Execute as":        User accessing the web app   (so note authors are real)
 *   - "Who has access":    Anyone within <your school domain>
 */

const SPREADSHEET_ID = '1zM4-AcA2DacMMFq91xfYqS9lmf1oP03u4ETNlkprvKg';
const STUDENT_SHEET  = 'Students';
const NOTES_SHEET    = 'Notes';
const SUBJECTS       = ['ELA', 'Math'];
const DEFAULT_GROUPS = ['Unassigned', 'Tier 1', 'Tier 2', 'Tier 3'];

// Expected header names in the Students sheet. Group columns are derived as
// `${subject} Group`, `${subject} Score Level`, `${subject} Scale Score`.
const COL = {
  id:       'Student ID',
  name:     'Name',
  grade:    'Grade',
  team:     'Team',
  sped:     'SPED',
  el:       'EL',
  five04:   '504',
  absences: 'Absences',
  tardies:  'Tardies',
  updated:  'Last Updated'
};

const STUDENT_HEADERS = [
  COL.id, COL.name, COL.grade, COL.team, COL.sped, COL.el, COL.five04,
  COL.absences, COL.tardies,
  'ELA Group', 'ELA Score Level', 'ELA Scale Score',
  'Math Group', 'Math Score Level', 'Math Scale Score',
  COL.updated
];

const NOTES_HEADERS = ['Timestamp', 'Student ID', 'Student Name', 'Grade', 'Subject', 'Author', 'Note'];


/* ───────────────────────────── Web app entry ───────────────────────────── */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('RTI / MTSS Student Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/* ─────────────────────────────── Helpers ───────────────────────────────── */

function headerMap_(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    const key = String(h == null ? '' : h).trim();
    if (key) map[key] = i;
  });
  return map;
}

function str_(v) { return String(v == null ? '' : v).trim(); }
function yn_(v)  { return str_(v).toLowerCase() === 'yes' ? 'Yes' : 'No'; }

function studentSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(STUDENT_SHEET);
  if (!sheet) throw new Error('Sheet "' + STUDENT_SHEET + '" not found. Run setupSheets() first.');
  return sheet;
}

function loadJson_(key)      { return JSON.parse(PropertiesService.getScriptProperties().getProperty(key) || '{}'); }
function saveJson_(key, obj) { PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(obj)); }
function loadGroups_()       { return loadJson_('rti_groups'); }
function saveGroups_(o)      { saveJson_('rti_groups', o); }
function loadColors_()       { return loadJson_('rti_colors'); }
function saveColors_(o)      { saveJson_('rti_colors', o); }


/* ─────────────────────────────── Read data ─────────────────────────────── */

/**
 * Returns { students, groups, colors }.
 *   students[i] = { id, name, grade, team, sped, el, five04, absences, tardies,
 *                   ELA: {group, level, scale}, Math: {group, level, scale} }
 *   groups[subject][grade] = [column names]
 *   colors[subject][grade] = { columnName: colorKey }
 */
function getData() {
  const sheet  = studentSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return { students: [], groups: loadGroups_(), colors: loadColors_() };
  }

  const col = headerMap_(values.shift());
  const get = (row, name) => (col[name] !== undefined ? row[col[name]] : '');

  const subjectBlock = (row, subject) => ({
    group: str_(get(row, subject + ' Group')) || 'Unassigned',
    level: str_(get(row, subject + ' Score Level')),
    scale: get(row, subject + ' Scale Score')
  });

  const students = values
    .filter(r => str_(get(r, COL.id)) || str_(get(r, COL.name)))
    .map(r => ({
      id:       str_(get(r, COL.id))    || 'N/A',
      name:     str_(get(r, COL.name))  || 'Unknown',
      grade:    str_(get(r, COL.grade)) || 'Unassigned',
      team:     str_(get(r, COL.team)),
      sped:     yn_(get(r, COL.sped)),
      el:       yn_(get(r, COL.el)),
      five04:   yn_(get(r, COL.five04)),
      absences: Number(get(r, COL.absences)) || 0,
      tardies:  Number(get(r, COL.tardies))  || 0,
      ELA:      subjectBlock(r, 'ELA'),
      Math:     subjectBlock(r, 'Math')
    }));

  // Reconcile stored board columns with what actually appears in the data.
  const groups = loadGroups_();
  const colors = loadColors_();
  let dirty = false;
  const grades = [...new Set(students.map(s => s.grade))];

  SUBJECTS.forEach(subject => {
    if (!groups[subject]) { groups[subject] = {}; dirty = true; }
    if (!colors[subject]) { colors[subject] = {}; dirty = true; }

    grades.forEach(grade => {
      if (!groups[subject][grade]) { groups[subject][grade] = DEFAULT_GROUPS.slice(); dirty = true; }
      if (!colors[subject][grade]) { colors[subject][grade] = {}; dirty = true; }

      students.filter(s => s.grade === grade).forEach(s => {
        const g = s[subject].group;
        if (g && g !== 'Unassigned' && !groups[subject][grade].includes(g)) {
          groups[subject][grade].push(g);
          dirty = true;
        }
      });
    });
  });

  if (dirty) { saveGroups_(groups); saveColors_(colors); }

  return { students, groups, colors };
}


/* ─────────────────────────── Move a student ────────────────────────────── */

function updateStudentGroup(studentId, subject, newGroup) {
  if (SUBJECTS.indexOf(subject) === -1) throw new Error('Unknown subject: ' + subject);

  const sheet  = studentSheet_();
  const values = sheet.getDataRange().getValues();
  const col    = headerMap_(values[0]);

  const idCol      = col[COL.id];
  const groupCol   = col[subject + ' Group'];
  const updatedCol = col[COL.updated];
  if (idCol === undefined || groupCol === undefined) {
    throw new Error('Missing required columns (Student ID / ' + subject + ' Group).');
  }

  const target = str_(studentId);
  for (let i = 1; i < values.length; i++) {
    if (str_(values[i][idCol]) === target) {
      sheet.getRange(i + 1, groupCol + 1).setValue(newGroup === 'Unassigned' ? '' : newGroup);
      if (updatedCol !== undefined) sheet.getRange(i + 1, updatedCol + 1).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false, error: 'Student ID not found: ' + target };
}


/* ─────────────────────── Manage board columns ──────────────────────────── */

function addGroup(subject, grade, groupName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const groups = loadGroups_();
    if (!groups[subject]) groups[subject] = {};
    if (!groups[subject][grade]) groups[subject][grade] = DEFAULT_GROUPS.slice();
    if (!groups[subject][grade].includes(groupName)) {
      groups[subject][grade].push(groupName);
      saveGroups_(groups);
    }
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function renameGroup(subject, grade, oldName, newName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const groups = loadGroups_();
    const colors = loadColors_();
    if (!groups[subject] || !groups[subject][grade]) return { success: false, error: 'Group set not found' };

    const arr = groups[subject][grade];
    const idx = arr.indexOf(oldName);
    if (idx !== -1) arr[idx] = newName;
    else if (!arr.includes(newName)) arr.push(newName);
    saveGroups_(groups);

    if (colors[subject] && colors[subject][grade] && colors[subject][grade][oldName]) {
      colors[subject][grade][newName] = colors[subject][grade][oldName];
      delete colors[subject][grade][oldName];
      saveColors_(colors);
    }

    // Re-label matching cells on the sheet.
    const sheet    = studentSheet_();
    const values   = sheet.getDataRange().getValues();
    const col      = headerMap_(values[0]);
    const gradeCol = col[COL.grade];
    const groupCol = col[subject + ' Group'];
    for (let i = 1; i < values.length; i++) {
      const rowGrade = str_(values[i][gradeCol]) || 'Unassigned';
      const rowGroup = str_(values[i][groupCol]);
      if (rowGrade === str_(grade) && rowGroup === str_(oldName)) {
        sheet.getRange(i + 1, groupCol + 1).setValue(newName);
      }
    }
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function setGroupColor(subject, grade, groupName, color) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const colors = loadColors_();
    if (!colors[subject]) colors[subject] = {};
    if (!colors[subject][grade]) colors[subject][grade] = {};
    colors[subject][grade][groupName] = color;
    saveColors_(colors);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}


/* ────────────────────────── Intervention notes ─────────────────────────── */

function getNotes(studentId) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(NOTES_SHEET);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const col    = headerMap_(values.shift());
  const target = str_(studentId);
  const tz     = Session.getScriptTimeZone();

  return values
    .filter(r => str_(r[col['Student ID']]) === target)
    .map(r => ({
      date:    r[col['Timestamp']] ? Utilities.formatDate(new Date(r[col['Timestamp']]), tz, 'MMM d, yyyy · h:mm a') : '',
      subject: str_(r[col['Subject']]),
      author:  str_(r[col['Author']]),
      note:    String(r[col['Note']] == null ? '' : r[col['Note']])
    }))
    .reverse(); // newest first
}

function addNote(studentId, studentName, grade, subject, note) {
  if (!str_(note)) return { success: false, error: 'Note is empty' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(NOTES_SHEET);
  if (!sheet) sheet = createNotesSheet_(ss);

  let author = '';
  try { author = Session.getActiveUser().getEmail() || ''; } catch (e) { author = ''; }

  sheet.appendRow([new Date(), str_(studentId), str_(studentName), str_(grade), str_(subject), author, str_(note)]);
  return { success: true };
}


/* ───────────────────── One-time spreadsheet builder ────────────────────── */

function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // --- Students ---
  let s = ss.getSheetByName(STUDENT_SHEET) || ss.insertSheet(STUDENT_SHEET);
  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 1, STUDENT_HEADERS.length).setValues([STUDENT_HEADERS]).setFontWeight('bold');
    s.setFrozenRows(1);

    const sample = [
      ['1001', 'Kai A.',     '3', 'Room 12', 'No',  'Yes', 'No', 4, 2, 'Tier 2', 'Below',      412, 'Tier 1', 'At',    455, ''],
      ['1002', 'Leilani B.', '3', 'Room 12', 'Yes', 'No',  'No', 9, 1, 'Tier 3', 'Well Below', 388, 'Tier 2', 'Below', 430, ''],
      ['1003', 'Mateo C.',   '4', 'Room 21', 'No',  'No',  'No', 1, 0, 'Tier 1', 'At',         498, 'Tier 2', 'Below', 440, '']
    ];
    s.getRange(2, 1, sample.length, sample[0].length).setValues(sample);

    // Yes/No dropdowns on SPED / EL / 504.
    const yn = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build();
    const map = headerMap_(STUDENT_HEADERS);
    [COL.sped, COL.el, COL.five04].forEach(name => {
      s.getRange(2, map[name] + 1, 1000).setDataValidation(yn);
    });

    s.autoResizeColumns(1, STUDENT_HEADERS.length);
  }

  // --- Notes ---
  if (!ss.getSheetByName(NOTES_SHEET)) createNotesSheet_(ss);

  return 'Setup complete — Students and Notes sheets are ready.';
}

function createNotesSheet_(ss) {
  const n = ss.insertSheet(NOTES_SHEET);
  n.getRange(1, 1, 1, NOTES_HEADERS.length).setValues([NOTES_HEADERS]).setFontWeight('bold');
  n.setFrozenRows(1);
  n.setColumnWidth(NOTES_HEADERS.indexOf('Note') + 1, 480);
  return n;
}
