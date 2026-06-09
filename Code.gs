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

// Expected header names in the Students sheet. Group and scale columns are
// derived as `${subject} Group` and `${subject} Scale Score`. BOY level columns
// use the source sheet's exact headers: "ELA BOY Level" and "MATH BOY Level".
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
  'ELA Group', 'ELA BOY Level', 'ELA Scale Score',
  'Math Group', 'MATH BOY Level', 'Math Scale Score',
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

function boyLevelColumn_(subject) {
  return subject === 'Math' ? 'MATH BOY Level' : subject + ' BOY Level';
}

function pushUnique_(arr, value) {
  const name = str_(value) || 'Unassigned';
  if (arr.indexOf(name) === -1) arr.push(name);
}

function orderGroups_(sheetGroups, savedOrder) {
  const current = sheetGroups || [];
  const saved = savedOrder || [];
  return saved
    .filter(name => current.indexOf(name) !== -1)
    .concat(current.filter(name => saved.indexOf(name) === -1));
}

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

function sheetGroupsFromStudents_(students, savedGroups) {
  const groups = {};

  SUBJECTS.forEach(subject => {
    groups[subject] = {};
    students.forEach(student => {
      const grade = str_(student.grade) || 'Unassigned';
      if (!groups[subject][grade]) groups[subject][grade] = [];
      pushUnique_(groups[subject][grade], student[subject].group);
    });

    Object.keys(groups[subject]).forEach(grade => {
      const savedOrder = savedGroups[subject] && savedGroups[subject][grade];
      groups[subject][grade] = orderGroups_(groups[subject][grade], savedOrder);
    });
  });

  return groups;
}

function sheetGroupsFor_(subject, grade) {
  const sheet = studentSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const col = headerMap_(values[0]);
  const gradeCol = col[COL.grade];
  const groupCol = col[subject + ' Group'];
  if (gradeCol === undefined || groupCol === undefined) {
    throw new Error('Missing required columns (Grade / ' + subject + ' Group).');
  }

  const cleanGrade = str_(grade) || 'Unassigned';
  const groups = [];
  for (let i = 1; i < values.length; i++) {
    const rowGrade = str_(values[i][gradeCol]) || 'Unassigned';
    if (rowGrade === cleanGrade) pushUnique_(groups, values[i][groupCol]);
  }
  return groups;
}


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
    return { students: [], groups: {}, colors: loadColors_() };
  }

  const col = headerMap_(values.shift());
  const get = (row, name) => (col[name] !== undefined ? row[col[name]] : '');

  const subjectBlock = (row, subject) => ({
    group: str_(get(row, subject + ' Group')) || 'Unassigned',
    level: str_(get(row, boyLevelColumn_(subject))),
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

  // Visible board columns come from the sheet's ELA Group and Math Group
  // values. Saved groups only preserve ordering for groups still in the sheet.
  const savedGroups = loadGroups_();
  const groups = sheetGroupsFromStudents_(students, savedGroups);
  const colors = loadColors_();

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
    if (SUBJECTS.indexOf(subject) === -1) throw new Error('Unknown subject: ' + subject);
    const cleanGrade = str_(grade) || 'Unassigned';
    const cleanGroup = str_(groupName);
    if (!cleanGroup) throw new Error('Group name is required.');

    const groups = loadGroups_();
    if (!groups[subject]) groups[subject] = {};
    const sheetGroups = sheetGroupsFor_(subject, cleanGrade);
    const currentGroups = orderGroups_(sheetGroups, groups[subject][cleanGrade]);
    if (currentGroups.indexOf(cleanGroup) === -1) {
      currentGroups.push(cleanGroup);
    }
    groups[subject][cleanGrade] = currentGroups;
    saveGroups_(groups);
    return { success: true, groups: currentGroups };
  } finally {
    lock.releaseLock();
  }
}

function renameGroup(subject, grade, oldName, newName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (SUBJECTS.indexOf(subject) === -1) throw new Error('Unknown subject: ' + subject);
    const cleanGrade = str_(grade) || 'Unassigned';
    const cleanOldName = str_(oldName);
    const cleanNewName = str_(newName);
    if (!cleanOldName || !cleanNewName) throw new Error('Group names are required.');

    const groups = loadGroups_();
    const colors = loadColors_();
    if (!groups[subject]) groups[subject] = {};

    const arr = orderGroups_(sheetGroupsFor_(subject, cleanGrade), groups[subject][cleanGrade]);
    const idx = arr.indexOf(cleanOldName);
    if (idx !== -1) arr[idx] = cleanNewName;
    else if (arr.indexOf(cleanNewName) === -1) arr.push(cleanNewName);
    groups[subject][cleanGrade] = arr;
    saveGroups_(groups);

    if (colors[subject] && colors[subject][cleanGrade] && colors[subject][cleanGrade][cleanOldName]) {
      colors[subject][cleanGrade][cleanNewName] = colors[subject][cleanGrade][cleanOldName];
      delete colors[subject][cleanGrade][cleanOldName];
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
      if (rowGrade === cleanGrade && rowGroup === cleanOldName) {
        sheet.getRange(i + 1, groupCol + 1).setValue(cleanNewName);
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

function updateGroupOrder(subject, grade, orderedGroups) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (SUBJECTS.indexOf(subject) === -1) throw new Error('Unknown subject: ' + subject);
    const cleanGrade = str_(grade) || 'Unassigned';
    const requested = Array.isArray(orderedGroups) ? orderedGroups : [];
    const seen = {};
    const cleanOrder = requested
      .map(str_)
      .filter(name => {
        if (!name || seen[name]) return false;
        seen[name] = true;
        return true;
      });
    if (cleanOrder.length === 0) throw new Error('Column order is empty.');

    const groups = loadGroups_();
    if (!groups[subject]) groups[subject] = {};
    groups[subject][cleanGrade] = cleanOrder;
    saveGroups_(groups);
    return { success: true, groups: cleanOrder };
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
      ['1001', 'Kai A.',     '3', 'Room 12', 'No',  'Yes', 'No', 4, 2, 'Tier 2', 'One grade level below',       412, 'Tier 1', 'On grade',        455, ''],
      ['1002', 'Leilani B.', '3', 'Room 12', 'Yes', 'No',  'No', 9, 1, 'Tier 3', 'Two or more grade levels below', 388, 'Tier 2', 'Below grade',    430, ''],
      ['1003', 'Mateo C.',   '4', 'Room 21', 'No',  'No',  'No', 1, 0, 'Tier 1', 'On or Above grade level',     498, 'Tier 2', 'Above grade',    440, '']
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
