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
const GROUP_METADATA_SHEET = 'Group Metadata';
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
  'ELA Group', 'ELA BOY Level', 'Current ELA Level', 'Current ORF',
  'Math Group', 'MATH BOY Level', 'Current Math Level', 'Math Scale Score',
  COL.updated
];

const NOTES_HEADERS = ['Timestamp', 'Student ID', 'Student Name', 'Grade', 'Subject', 'Author', 'Note'];
const GROUP_METADATA_HEADERS = ['Subject', 'Grade', 'Group', 'Color', 'Location', 'Skill', 'Curriculum', 'Sort Order'];


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

function canonicalSubject_(subject) {
  const clean = str_(subject).toLowerCase();
  return SUBJECTS.filter(s => s.toLowerCase() === clean)[0] || '';
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

function ensureSheetHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = headerMap_(current);
  let nextCol = lastCol + 1;
  headers.forEach(header => {
    if (col[header] === undefined) {
      sheet.getRange(1, nextCol).setValue(header).setFontWeight('bold');
      nextCol++;
    }
  });
  sheet.setFrozenRows(1);
}

function groupMetadataSheet_(ss) {
  const sheet = ss.getSheetByName(GROUP_METADATA_SHEET) || ss.insertSheet(GROUP_METADATA_SHEET);
  ensureSheetHeaders_(sheet, GROUP_METADATA_HEADERS);
  return sheet;
}

function metadataOrder_(metadata, subject, grade) {
  const byGroup = metadata[subject] && metadata[subject][grade] ? metadata[subject][grade] : {};
  return Object.keys(byGroup).sort((a, b) => {
    const ao = byGroup[a].order;
    const bo = byGroup[b].order;
    if (ao !== bo) return ao - bo;
    return byGroup[a].row - byGroup[b].row;
  });
}

function readGroupMetadata_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = groupMetadataSheet_(ss);
  const values = sheet.getDataRange().getValues();
  const metadata = {};
  if (values.length <= 1) return metadata;

  const col = headerMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    const subject = canonicalSubject_(values[i][col['Subject']]);
    if (!subject) continue;

    const grade = str_(values[i][col['Grade']]) || 'Unassigned';
    const group = str_(values[i][col['Group']]) || 'Unassigned';
    const rawOrder = Number(values[i][col['Sort Order']]);

    if (!metadata[subject]) metadata[subject] = {};
    if (!metadata[subject][grade]) metadata[subject][grade] = {};
    metadata[subject][grade][group] = {
      color: str_(values[i][col['Color']]),
      location: str_(values[i][col['Location']]),
      skill: str_(values[i][col['Skill']]),
      curriculum: str_(values[i][col['Curriculum']]),
      order: isNaN(rawOrder) ? i : rawOrder,
      row: i
    };
  }
  return metadata;
}

function metadataGroupsFor_(metadata, subject, grade) {
  return Object.keys((metadata[subject] && metadata[subject][grade]) || {});
}

function currentGroupsFor_(subject, grade, metadata) {
  const groups = sheetGroupsFor_(subject, grade);
  metadataGroupsFor_(metadata, subject, grade).forEach(group => pushUnique_(groups, group));
  return orderGroups_(groups, metadataOrder_(metadata, subject, grade));
}

function colorsFromMetadata_(metadata) {
  const colors = {};
  SUBJECTS.forEach(subject => {
    colors[subject] = {};
    Object.keys(metadata[subject] || {}).forEach(grade => {
      colors[subject][grade] = {};
      Object.keys(metadata[subject][grade]).forEach(group => {
        const color = metadata[subject][grade][group].color;
        if (color) colors[subject][grade][group] = color;
      });
    });
  });
  return colors;
}

function columnMetaFromMetadata_(metadata) {
  const columnMeta = {};
  SUBJECTS.forEach(subject => {
    columnMeta[subject] = {};
    Object.keys(metadata[subject] || {}).forEach(grade => {
      columnMeta[subject][grade] = {};
      Object.keys(metadata[subject][grade]).forEach(group => {
        const meta = metadata[subject][grade][group];
        columnMeta[subject][grade][group] = {
          location: meta.location,
          skill: meta.skill,
          curriculum: meta.curriculum
        };
      });
    });
  });
  return columnMeta;
}

function upsertGroupMetadata_(subject, grade, groupName, patch) {
  const cleanSubject = canonicalSubject_(subject);
  if (!cleanSubject) throw new Error('Unknown subject: ' + subject);
  const cleanGrade = str_(grade) || 'Unassigned';
  const cleanGroup = str_(groupName) || 'Unassigned';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = groupMetadataSheet_(ss);
  const values = sheet.getDataRange().getValues();
  const col = headerMap_(values[0]);

  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (
      canonicalSubject_(values[i][col['Subject']]) === cleanSubject &&
      (str_(values[i][col['Grade']]) || 'Unassigned') === cleanGrade &&
      (str_(values[i][col['Group']]) || 'Unassigned') === cleanGroup
    ) {
      rowIndex = i + 1;
      break;
    }
  }

  const row = rowIndex === -1
    ? new Array(sheet.getLastColumn()).fill('')
    : sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

  row[col['Subject']] = cleanSubject;
  row[col['Grade']] = cleanGrade;
  row[col['Group']] = cleanGroup;
  if (patch.color !== undefined) row[col['Color']] = str_(patch.color);
  if (patch.location !== undefined) row[col['Location']] = str_(patch.location);
  if (patch.skill !== undefined) row[col['Skill']] = str_(patch.skill);
  if (patch.curriculum !== undefined) row[col['Curriculum']] = str_(patch.curriculum);
  if (patch.order !== undefined) row[col['Sort Order']] = patch.order;

  if (rowIndex === -1) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  }
}

function renameGroupMetadata_(subject, grade, oldName, newName) {
  const cleanSubject = canonicalSubject_(subject);
  if (!cleanSubject) throw new Error('Unknown subject: ' + subject);
  const cleanGrade = str_(grade) || 'Unassigned';
  const cleanOldName = str_(oldName) || 'Unassigned';
  const cleanNewName = str_(newName) || 'Unassigned';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = groupMetadataSheet_(ss);
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    upsertGroupMetadata_(cleanSubject, cleanGrade, cleanNewName, {});
    return;
  }

  const col = headerMap_(values[0]);
  for (let i = 1; i < values.length; i++) {
    if (
      canonicalSubject_(values[i][col['Subject']]) === cleanSubject &&
      (str_(values[i][col['Grade']]) || 'Unassigned') === cleanGrade &&
      (str_(values[i][col['Group']]) || 'Unassigned') === cleanOldName
    ) {
      sheet.getRange(i + 1, col['Group'] + 1).setValue(cleanNewName);
      return;
    }
  }
  upsertGroupMetadata_(cleanSubject, cleanGrade, cleanNewName, {});
}

function sheetGroupsFromStudents_(students, metadata) {
  const groups = {};

  SUBJECTS.forEach(subject => {
    groups[subject] = {};
    students.forEach(student => {
      const grade = str_(student.grade) || 'Unassigned';
      if (!groups[subject][grade]) groups[subject][grade] = [];
      pushUnique_(groups[subject][grade], student[subject].group);
    });

    Object.keys(metadata[subject] || {}).forEach(grade => {
      if (!groups[subject][grade]) groups[subject][grade] = [];
      metadataGroupsFor_(metadata, subject, grade).forEach(group => pushUnique_(groups[subject][grade], group));
    });

    Object.keys(groups[subject]).forEach(grade => {
      groups[subject][grade] = orderGroups_(groups[subject][grade], metadataOrder_(metadata, subject, grade));
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
 *                   ELA: {group, level, currentLevel, scale},
 *                   Math: {group, level, currentLevel, scale} }
 *   groups[subject][grade] = [column names]
 *   colors[subject][grade] = { columnName: colorKey }
 */
function getData() {
  const sheet  = studentSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return { students: [], groups: {}, colors: {}, columnMeta: {} };
  }

  const col = headerMap_(values.shift());
  const get = (row, name) => (col[name] !== undefined ? row[col[name]] : '');

  const subjectBlock = (row, subject) => ({
    group: str_(get(row, subject + ' Group')) || 'Unassigned',
    level: str_(get(row, boyLevelColumn_(subject))),
    currentLevel: str_(get(row, 'Current ' + subject + ' Level')),
    scale: get(row, subject === 'ELA' ? 'Current ORF' : subject + ' Scale Score')
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

  // Visible board columns come from the sheet's ELA/Math group values plus any
  // explicit rows in Group Metadata. Metadata rows hold order, color, and header
  // fields for empty groups as well as groups with assigned students.
  const groupMetadata = readGroupMetadata_();
  const groups = sheetGroupsFromStudents_(students, groupMetadata);
  const colors = colorsFromMetadata_(groupMetadata);
  const columnMeta = columnMetaFromMetadata_(groupMetadata);

  return { students, groups, colors, columnMeta };
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

    const metadata = readGroupMetadata_();
    const currentGroups = currentGroupsFor_(subject, cleanGrade, metadata);
    if (currentGroups.indexOf(cleanGroup) === -1) {
      currentGroups.push(cleanGroup);
    }
    upsertGroupMetadata_(subject, cleanGrade, cleanGroup, { order: currentGroups.indexOf(cleanGroup) + 1 });
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

    const metadata = readGroupMetadata_();
    const arr = currentGroupsFor_(subject, cleanGrade, metadata);
    const idx = arr.indexOf(cleanOldName);
    if (idx !== -1) arr[idx] = cleanNewName;
    else if (arr.indexOf(cleanNewName) === -1) arr.push(cleanNewName);

    renameGroupMetadata_(subject, cleanGrade, cleanOldName, cleanNewName);
    arr.forEach((name, index) => upsertGroupMetadata_(subject, cleanGrade, name, { order: index + 1 }));

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
    upsertGroupMetadata_(subject, grade, groupName, { color: str_(color) });
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function setGroupMetadata(subject, grade, groupName, metadata) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const allowed = {};
    metadata = metadata || {};
    ['location', 'skill', 'curriculum'].forEach(field => {
      if (metadata[field] !== undefined) allowed[field] = metadata[field];
    });
    upsertGroupMetadata_(subject, grade, groupName, allowed);
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

    cleanOrder.forEach((name, index) => {
      upsertGroupMetadata_(subject, cleanGrade, name, { order: index + 1 });
    });
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
      ['1001', 'Kai A.',     '3', 'Room 12', 'No',  'Yes', 'No', 4, 2, 'Tier 2', 'One grade level below',       'On or Above grade level', 412, 'Tier 1', 'On grade',        'On grade',   455, ''],
      ['1002', 'Leilani B.', '3', 'Room 12', 'Yes', 'No',  'No', 9, 1, 'Tier 3', 'Two or more grade levels below', 'One grade level below', 388, 'Tier 2', 'Below grade',    'On grade',   430, ''],
      ['1003', 'Mateo C.',   '4', 'Room 21', 'No',  'No',  'No', 1, 0, 'Tier 1', 'On or Above grade level',     'On or Above grade level', 498, 'Tier 2', 'Above grade',    'Above grade', 440, '']
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

  // --- Group Metadata ---
  groupMetadataSheet_(ss).autoResizeColumns(1, GROUP_METADATA_HEADERS.length);

  return 'Setup complete — Students, Notes, and Group Metadata sheets are ready.';
}

function createNotesSheet_(ss) {
  const n = ss.insertSheet(NOTES_SHEET);
  n.getRange(1, 1, 1, NOTES_HEADERS.length).setValues([NOTES_HEADERS]).setFontWeight('bold');
  n.setFrozenRows(1);
  n.setColumnWidth(NOTES_HEADERS.indexOf('Note') + 1, 480);
  return n;
}
