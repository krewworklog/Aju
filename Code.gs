/* ═══════════════════════════════════════════════════════════════
   ██  KREW  MARKETING  —  WorkLog  backend
   ██  PASTE THIS ONLY INTO THE SCRIPT BOUND TO:  "worklog database"
   ██  (Krew staff work logs, attendance, leads, increment program)
   ██  DO NOT paste into the Al Rasa project.
   ═══════════════════════════════════════════════════════════════ */
/**
 * WorkLog — Google Sheets Backend
 *
 * v9  — entry proof screenshots + manager reviews
 * v10 — proof auto-cleanup (7 days), announcements, leaderboard
 *
 * Deploy as Web App: Execute as "Me", Access "Anyone"
 */

const SHEETS = {
  staff:    ['id', 'name', 'username', 'password', 'createdAt'],
  managers: ['id', 'name', 'username', 'password', 'createdAt'],
  clients:  ['name', 'services'],
  entries:  ['id', 'staffId', 'staffName', 'date', 'gnotes', 'rows', 'late', 'reason', 'approved', 'submittedAt', 'usedCodeId', 'editedAt', 'entry', 'entryLabel', 'deadline'],
  codes:    ['id', 'code', 'staffId', 'staffName', 'note', 'issuedByRole', 'issuedById', 'issuedByName', 'generatedAt', 'generatedDate', 'generatedTime', 'used', 'usedBy', 'usedByName', 'usedAt'],
  reports:  ['id', 'weekStart', 'weekEnd', 'submittedAt', 'submittedBy', 'submittedById', 'breakdown', 'note', 'totalHours'],
  tasks:    ['id', 'title', 'description', 'assignedTo', 'assignedToName', 'createdByRole', 'createdById', 'createdByName', 'deadline', 'done', 'doneAt', 'staffNote', 'imageUrl', 'createdAt'],
  leads:    ['id', 'targetId', 'staffId', 'staffName', 'f1', 'f2', 'f3', 'success', 'createdAt', 'salesStatus', 'salesReason', 'salesFee', 'salesBy', 'salesAt', 'salesNextStep'],
  attendance: ['id', 'staffId', 'staffName', 'date', 'checkIn', 'checkOut', 'otHours', 'otNote', 'createdAt'],
  // Salary Increment Program
  incMembers:  ['id', 'staffId', 'staffName', 'joinedAt', 'active'],
  // What each client is owed each month, per service (from the onboarding form)
  deliverables:['id', 'client', 'service', 'qty', 'note'],
  // A staff member ticking off a delivered item for a given month
  incWork:     ['id', 'month', 'staffId', 'staffName', 'client', 'service', 'done', 'doneAt', 'note'],
  // Monthly PDF report, due by the 5th of the following month
  incReports:  ['id', 'month', 'staffId', 'staffName', 'fileUrl', 'fileName', 'submittedAt', 'onTime'],
  // v10 — HR / executive announcements
  announcements: ['id', 'title', 'body', 'postedById', 'postedByName', 'postedAt', 'showFrom', 'showUntil'],
  // v10 — executive management score for the leaderboard
  mgmtScores:  ['id', 'month', 'staffId', 'staffName', 'score', 'note', 'setBy', 'setAt'],
  config:   ['key', 'value']
};

const CODE_VERSION = 'v22-announcements-leaderboard';

const DEFAULT_CLIENTS = ['Baaqat Flowers','Flovia Flowers','8th Cafe','Florens Flowers','Flat Chocolate','Al Rasa','Fedora Perfumes','Elite Party','Hair Salon'];
const DEFAULT_EXEC_CODE = '99999';
const UPLOAD_FOLDER = 'WorkLog Uploads';

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }


// ── Wrong-project guard ────────────────────────────────────────
// This code must only ever run against the Krew WorkLog spreadsheet. If it is
// pasted into another project by mistake, refuse instead of taking over.
const APP_ID = 'krew';
function assertRightSpreadsheet(ss) {
  // Two cheap lookups instead of enumerating every tab, and cached — this
  // runs on every request, so it must not cost anything noticeable.
  const cache = CacheService.getScriptCache();
  if (cache.get('sheet_ok_krew') === '1') return;
  // Tab names alone are no longer conclusive: both spreadsheets picked up
  // stray tabs from the other app during the mix-up, so Al Rasa's database
  // now HAS an 'entries' tab. The file name is the reliable signal.
  const name = String(ss.getName() || '').toLowerCase();
  if (name.indexOf('al rasa') >= 0 || name.indexOf('alrasa') >= 0) {
    throw new Error('WRONG PROJECT: this is the KREW backend, but it is bound to "' + ss.getName() + '" (the Al Rasa database). Paste Desktop/AlRasa/Code.gs here instead.');
  }
  // The decisive test: Al Rasa's config tab holds openingBalance/openingDate,
  // which Krew never writes. This survives the stray-tab contamination that
  // made the old tab-name check useless.
  const cfg = ss.getSheetByName('config');
  if (cfg) {
    const keys = cfg.getDataRange().getValues().map(r => String(r[0]));
    if (keys.indexOf('openingBalance') >= 0 || keys.indexOf('openingDate') >= 0) {
      throw new Error('WRONG PROJECT: this is the KREW backend, but it is bound to "' + ss.getName() + '", which is the Al Rasa database. Paste Desktop/AlRasa/Code.gs here instead.');
    }
  }
  const hasEntries = !!ss.getSheetByName('entries');
  const hasInvoices = !!ss.getSheetByName('invoices');
  if (!hasEntries && hasInvoices) {
    throw new Error('WRONG PROJECT: this is the KREW backend, but the spreadsheet looks like Al Rasa Collection DB. Paste Desktop/AlRasa/Code.gs here instead.');
  }
  cache.put('sheet_ok_krew', '1', 21600);
}

function handleRequest(e) {
  try {
    const params = (e.parameter || {});
    const action = params.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    assertRightSpreadsheet(ss);
    ensureSheets(ss);
    try { ensureMonthlyReportTasks(ss); } catch (e) { /* never block a request on this */ }

    let result;
    switch (action) {
      case 'getAll':       result = getAll(ss); break;
      case 'addStaff':     result = addRow(ss, 'staff', JSON.parse(params.data)); break;
      case 'deleteStaff':  result = deleteRow(ss, 'staff', params.id); break;
      case 'resetStaff':   result = updateField(ss, 'staff', params.id, 'password', params.password); break;
      case 'addManager':   result = addRow(ss, 'managers', JSON.parse(params.data)); break;
      case 'deleteManager':result = deleteRow(ss, 'managers', params.id); break;
      case 'resetManager': result = updateField(ss, 'managers', params.id, 'password', params.password); break;
      case 'addClient':    result = addClient(ss, params.name, params.services); break;
      case 'setClientServices': result = setClientServices(ss, params.name, params.services); break;
      case 'deleteClient': result = deleteClient(ss, params.name); break;
      case 'addEntry':     result = addEntryWithProofs(ss, params); break;
      case 'updateEntry':  result = updateEntryWithProofs(ss, params); break;
      case 'deleteEntry':  result = deleteRow(ss, 'entries', params.id); break;
      case 'reviewEntry':  result = reviewEntry(ss, params); break;
      case 'addCode':      result = addRow(ss, 'codes', JSON.parse(params.data)); break;
      case 'useCode':      result = useCode(ss, params.codeId, params.staffId, params.staffName); break;
      case 'revokeCode':   result = deleteRow(ss, 'codes', params.id); break;
      case 'addReport':    result = addRow(ss, 'reports', JSON.parse(params.data)); break;
      case 'deleteReport': result = deleteRow(ss, 'reports', params.id); break;
      case 'addTask':      result = addRow(ss, 'tasks', JSON.parse(params.data)); break;
      case 'deleteTask':   result = deleteRow(ss, 'tasks', params.id); break;
      case 'completeTask': result = completeTask(ss, params); break;
      case 'addLead':      result = addRow(ss, 'leads', JSON.parse(params.data)); break;
      case 'setLeadSuccess': result = updateField(ss, 'leads', params.id, 'success', params.value === 'true'); break;
      case 'setLeadSales': result = setLeadSales(ss, params); break;
      case 'deleteLead':   result = deleteRow(ss, 'leads', params.id); break;
      case 'attnCheckIn':  result = attnUpsert(ss, params); break;
      case 'attnCheckOut': result = attnUpsert(ss, params); break;
      case 'attnOvertime': result = attnUpsert(ss, params); break;
      case 'deleteAttendance': result = deleteRow(ss, 'attendance', params.id); break;
      case 'setExecCode':  result = setConfig(ss, 'execCode', params.value); break;
      // ---- Salary Increment Program ----
      case 'incJoin':      result = incJoin(ss, params); break;
      case 'incLeave':     result = incSetActive(ss, params.staffId, false); break;
      case 'addDeliverable':    result = addRow(ss, 'deliverables', JSON.parse(params.data)); break;
      case 'deleteDeliverable': result = deleteRow(ss, 'deliverables', params.id); break;
      case 'setIncWork':   result = setIncWork(ss, params); break;
      case 'submitIncReport': result = submitIncReport(ss, params); break;
      case 'setConfigKey': result = setConfig(ss, params.key, params.value); break;
      // ---- v10: announcements + leaderboard ----
      case 'addAnnouncement':    result = addAnnouncement(ss, params); break;
      case 'deleteAnnouncement': result = deleteRow(ss, 'announcements', params.id); break;
      case 'setMgmtScore':       result = setMgmtScore(ss, params); break;
      case 'getLeaderboard':     result = getLeaderboard(ss, params.month); break;
      default: result = { error: 'Unknown action' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: String(err), stack: err.stack });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ensureSheets(ss) {
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
      sh.setFrozenRows(1);
      if (name === 'clients') {
        DEFAULT_CLIENTS.forEach(c => sh.appendRow([c]));
      }
      if (name === 'config') {
        sh.appendRow(['execCode', DEFAULT_EXEC_CODE]);
      }
    } else {
      // Migration: append any header columns added in later versions
      // without disturbing existing data.
      const lastCol = sh.getLastColumn();
      const existing = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      const missing = SHEETS[name].filter(h => existing.indexOf(h) === -1);
      if (missing.length) {
        sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
        sh.setFrozenRows(1);
      }
    }
  });
}

// The sheet's real header row — respects whatever column order the live
// spreadsheet has, so a script update can never scramble columns.
function sheetHeaders(sh, name) {
  const lastCol = sh.getLastColumn();
  const headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].filter(String) : [];
  return headers.length ? headers : SHEETS[name];
}

function sheetToObjects(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).filter(row => row.some(c => c !== '' && c !== null)).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// Sheets auto-converts date strings in cells to Date values; send them back
// as plain yyyy-MM-dd strings (in the spreadsheet's timezone) so the frontend
// can compare them.
function dateCellToString(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return v;
}
// Same idea for time-only cells (checkIn/checkOut) that Sheets stored as a
// time value — format in the sheet timezone so it round-trips exactly.
function timeCellToString(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'HH:mm');
  }
  return v;
}

function getAll(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const staff = sheetToObjects(ss.getSheetByName('staff')).map(s => ({ ...s, createdAt: dateCellToString(s.createdAt, tz) }));
  const managers = sheetToObjects(ss.getSheetByName('managers')).map(m => ({ ...m, createdAt: dateCellToString(m.createdAt, tz) }));
  const rawClients = sheetToObjects(ss.getSheetByName('clients'));
  const clients = rawClients.map(c => c.name);
  const clientServices = {};
  rawClients.forEach(c => {
    let svc = [];
    if (c.services) { try { svc = typeof c.services === 'string' ? JSON.parse(c.services) : c.services; } catch (e) { svc = []; } }
    clientServices[c.name] = Array.isArray(svc) ? svc : [];
  });
  const rawEntries = sheetToObjects(ss.getSheetByName('entries'));
  const entries = rawEntries.map(e => ({
    ...e,
    date: dateCellToString(e.date, tz),
    rows: e.rows ? JSON.parse(e.rows) : [],
    late: e.late === true || e.late === 'TRUE' || e.late === 'true',
    approved: e.approved === true || e.approved === 'TRUE' || e.approved === 'true'
  }));
  const rawCodes = sheetToObjects(ss.getSheetByName('codes'));
  const codes = rawCodes.map(c => ({
    ...c,
    used: c.used === true || c.used === 'TRUE' || c.used === 'true'
  }));
  const rawReports = sheetToObjects(ss.getSheetByName('reports'));
  const reports = rawReports.map(r => ({
    ...r,
    weekStart: dateCellToString(r.weekStart, tz),
    weekEnd: dateCellToString(r.weekEnd, tz),
    breakdown: r.breakdown ? (typeof r.breakdown === 'string' ? JSON.parse(r.breakdown) : r.breakdown) : [],
    totalHours: parseFloat(r.totalHours) || 0
  }));
  const rawTasks = sheetToObjects(ss.getSheetByName('tasks'));
  const tasks = rawTasks.map(t => ({
    ...t,
    done: t.done === true || t.done === 'TRUE' || t.done === 'true'
  }));
  const rawLeads = sheetToObjects(ss.getSheetByName('leads'));
  const leads = rawLeads.map(l => ({
    ...l,
    createdAt: dateCellToString(l.createdAt, tz),
    success: l.success === true || l.success === 'TRUE' || l.success === 'true'
  }));
  const rawAttendance = sheetToObjects(ss.getSheetByName('attendance'));
  const attendance = rawAttendance.map(a => ({
    ...a,
    date: dateCellToString(a.date, tz),
    checkIn: timeCellToString(a.checkIn, tz),
    checkOut: timeCellToString(a.checkOut, tz)
  }));
  const config = sheetToObjects(ss.getSheetByName('config'));
  const execCode = (config.find(c => c.key === 'execCode') || {}).value || DEFAULT_EXEC_CODE;
  const configMap = {}; config.forEach(c => { configMap[c.key] = c.value; });
  // ---- Salary Increment Program ----
  const incMembers = sheetToObjects(ss.getSheetByName('incMembers')).map(m => ({
    ...m, joinedAt: dateCellToString(m.joinedAt, tz),
    active: !(m.active === false || m.active === 'FALSE' || m.active === 'false')
  }));
  const deliverables = sheetToObjects(ss.getSheetByName('deliverables'));
  const incWork = sheetToObjects(ss.getSheetByName('incWork')).map(w => ({
    ...w, doneAt: dateCellToString(w.doneAt, tz),
    done: w.done === true || w.done === 'TRUE' || w.done === 'true'
  }));
  const incReports = sheetToObjects(ss.getSheetByName('incReports')).map(r => ({
    ...r, submittedAt: dateCellToString(r.submittedAt, tz),
    onTime: r.onTime === true || r.onTime === 'TRUE' || r.onTime === 'true'
  }));
  // ---- v10: announcements ----
  const announcements = getAnnouncements(ss);
  return { staff, managers, clients, clientServices, entries, codes, reports, tasks, leads, attendance,
           incMembers, deliverables, incWork, incReports, announcements,
           config: configMap, codeVersion: CODE_VERSION, execCode: String(execCode) };
}

// Attendance: one record per staff per day. Check-in / check-out / overtime all
// upsert into that day's row (finding by staffId + date, creating if absent).
function attnUpsert(ss, params) {
  const sh = ss.getSheetByName('attendance');
  const headers = sheetHeaders(sh, 'attendance');
  const col = h => headers.indexOf(h);
  const tz = ss.getSpreadsheetTimeZone();
  const data = sh.getDataRange().getValues();
  // Write text-forced cells so Sheets never turns "10:05" or "2026-07-27" into
  // a time/date serial (which caused mismatched rows and drifted times).
  const setText = (r, c, v) => { const cell = sh.getRange(r, c + 1); cell.setNumberFormat('@'); cell.setValue(v); };
  for (let i = 1; i < data.length; i++) {
    const rowDate = dateCellToString(data[i][col('date')], tz);
    if (String(data[i][col('staffId')]) === String(params.staffId) && String(rowDate) === String(params.date)) {
      if (params.checkIn) setText(i + 1, col('checkIn'), params.checkIn);
      if (params.checkOut) setText(i + 1, col('checkOut'), params.checkOut);
      if (params.otHours !== undefined && params.otHours !== '') sh.getRange(i + 1, col('otHours') + 1).setValue(params.otHours);
      if (params.otNote !== undefined && params.otNote !== '') sh.getRange(i + 1, col('otNote') + 1).setValue(params.otNote);
      return { ok: true, updated: true };
    }
  }
  const obj = {
    id: params.id || ('a' + Date.now()),
    staffId: params.staffId, staffName: params.staffName, date: params.date,
    checkIn: params.checkIn || '', checkOut: params.checkOut || '',
    otHours: params.otHours || '', otNote: params.otNote || '', createdAt: params.date
  };
  const r = sh.getLastRow() + 1;
  headers.forEach((h, idx) => {
    const v = obj[h] !== undefined ? obj[h] : '';
    if (h === 'date' || h === 'checkIn' || h === 'checkOut' || h === 'createdAt') setText(r, idx, v);
    else sh.getRange(r, idx + 1).setValue(v);
  });
  return { ok: true, created: true };
}

// Columns that must stay literal text — Sheets would otherwise turn values
// like "3-4" into a date, or strip leading zeros.
const TEXT_COLS = { qty: 1, month: 1, date: 1, chequeNo: 1, joinedAt: 1, doneAt: 1, submittedAt: 1 };

function addRow(ss, sheetName, obj) {
  const sh = ss.getSheetByName(sheetName);
  const headers = sheetHeaders(sh, sheetName);
  const r = sh.getLastRow() + 1;
  headers.forEach((h, i) => {
    let v = obj[h];
    if ((h === 'rows' || h === 'breakdown') && (Array.isArray(v) || typeof v === 'object')) v = JSON.stringify(v);
    if (v === undefined || v === null) v = '';
    const cell = sh.getRange(r, i + 1);
    if (TEXT_COLS[h]) cell.setNumberFormat('@');
    cell.setValue(v);
  });
  return { ok: true };
}

function deleteRow(ss, sheetName, id) {
  const sh = ss.getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

function updateField(ss, sheetName, id, field, value) {
  const sh = ss.getSheetByName(sheetName);
  const headers = sheetHeaders(sh, sheetName);
  const colIdx = headers.indexOf(field);
  if (colIdx < 0) return { ok: false, error: 'Bad field' };
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.getRange(i + 1, colIdx + 1).setValue(value);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

// ---- Entry proofs (v9) ----

function uploadFolder() {
  const it = DriveApp.getFoldersByName(UPLOAD_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(UPLOAD_FOLDER);
}

function saveImageToDrive(imageData, imageType, name) {
  const blob = Utilities.newBlob(Utilities.base64Decode(imageData), imageType || 'image/jpeg', name);
  const file = uploadFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

// proofs param: JSON [{i, imageData, imageType}] — uploads each image and
// stores its Drive URL on rows[i].proofImg before the entry is written.
function applyProofs(obj, proofsJson) {
  if (!proofsJson) return obj;
  const proofs = JSON.parse(proofsJson);
  const rows = obj.rows || [];
  proofs.forEach(p => {
    if (!p || !p.imageData || rows[p.i] === undefined) return;
    rows[p.i].proofImg = saveImageToDrive(p.imageData, p.imageType, 'proof_' + obj.id + '_row' + p.i + '.jpg');
  });
  obj.rows = rows;
  return obj;
}

function addEntryWithProofs(ss, params) {
  const obj = applyProofs(JSON.parse(params.data), params.proofs);
  // One entry per staff per date per slot. If the same slot is submitted
  // again — a double tap, or a retry after a lost response — update that
  // row instead of adding a second one that would be counted twice.
  const sh = ss.getSheetByName('entries');
  const headers = sheetHeaders(sh, 'entries');
  const tz = ss.getSpreadsheetTimeZone();
  const col = h => headers.indexOf(h);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const sameStaff = String(data[i][col('staffId')]) === String(obj.staffId);
    const sameDate  = String(dateCellToString(data[i][col('date')], tz)) === String(obj.date);
    const sameSlot  = String(data[i][col('entry')]) === String(obj.entry);
    const sameId    = String(data[i][0]) === String(obj.id);
    if (sameId || (sameStaff && sameDate && sameSlot && obj.entry)) {
      obj.id = data[i][0];
      return updateEntryRow(ss, obj, i + 1);
    }
  }
  return addRow(ss, 'entries', obj);
}

// Write an entry object over an existing row.
function updateEntryRow(ss, obj, rowNum) {
  const sh = ss.getSheetByName('entries');
  const headers = sheetHeaders(sh, 'entries');
  const row = headers.map(h => {
    let v = obj[h];
    if (h === 'rows' && Array.isArray(v)) v = JSON.stringify(v);
    if (v === undefined || v === null) v = '';
    return v;
  });
  sh.getRange(rowNum, 1, 1, headers.length).setValues([row]);
  return { ok: true, deduped: true };
}

function updateEntryWithProofs(ss, params) {
  const obj = applyProofs(JSON.parse(params.data), params.proofs);
  const sh = ss.getSheetByName('entries');
  const headers = sheetHeaders(sh, 'entries');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(obj.id)) {
      const row = headers.map(h => {
        let v = obj[h];
        if (h === 'rows' && Array.isArray(v)) v = JSON.stringify(v);
        if (v === undefined || v === null) v = '';
        return v;
      });
      sh.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

// Manager review verdict, stored inside the rows JSON (rows[0]._review) so no
// new sheet column is needed: {status, by, note, at}.
function reviewEntry(ss, params) {
  const sh = ss.getSheetByName('entries');
  const headers = sheetHeaders(sh, 'entries');
  const rowsIdx = headers.indexOf('rows');
  if (rowsIdx < 0) return { ok: false, error: 'No rows column' };
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.id)) {
      let rows = [];
      try { rows = data[i][rowsIdx] ? JSON.parse(data[i][rowsIdx]) : []; } catch (e) { rows = []; }
      if (!rows.length) rows = [{}];
      rows[0]._review = {
        status: params.status === 'rejected' ? 'rejected' : 'approved',
        by: params.by || '',
        note: params.note || '',
        at: params.at || ''
      };
      sh.getRange(i + 1, rowsIdx + 1).setValue(JSON.stringify(rows));
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

// ---- Tasks (v8) ----

function completeTask(ss, params) {
  const sh = ss.getSheetByName('tasks');
  const headers = sheetHeaders(sh, 'tasks');
  const data = sh.getDataRange().getValues();
  const col = h => headers.indexOf(h);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.id)) {
      const done = params.done === true || params.done === 'true';
      sh.getRange(i + 1, col('done') + 1).setValue(done);
      sh.getRange(i + 1, col('doneAt') + 1).setValue(done ? new Date().toISOString() : '');
      if (params.staffNote !== undefined) sh.getRange(i + 1, col('staffNote') + 1).setValue(params.staffNote);
      if (params.imageData) {
        const url = saveImageToDrive(params.imageData, params.imageType, 'task_' + params.id + '.jpg');
        sh.getRange(i + 1, col('imageUrl') + 1).setValue(url);
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

// Manager sets the sales outcome for a logged enquiry.
function setLeadSales(ss, params) {
  const sh = ss.getSheetByName('leads');
  const headers = sheetHeaders(sh, 'leads');
  const data = sh.getDataRange().getValues();
  const col = h => headers.indexOf(h);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(params.id)) {
      const set = (h, v) => { if (col(h) >= 0) sh.getRange(i + 1, col(h) + 1).setValue(v); };
      set('salesStatus', params.status || '');
      set('salesReason', params.reason || '');
      set('salesFee', params.fee || '');
      set('salesBy', params.by || '');
      set('salesAt', params.at || '');
      if (params.nextStep !== undefined) set('salesNextStep', params.nextStep || '');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

function addClient(ss, name, services) {
  const sh = ss.getSheetByName('clients');
  const existing = sheetToObjects(sh).map(c => c.name);
  if (existing.includes(name)) return { ok: false, error: 'exists' };
  const headers = sheetHeaders(sh, 'clients');
  const row = headers.map(h => h === 'name' ? name : (h === 'services' ? (services || '[]') : ''));
  sh.appendRow(row);
  return { ok: true };
}

function setClientServices(ss, name, services) {
  const sh = ss.getSheetByName('clients');
  const headers = sheetHeaders(sh, 'clients');
  const col = headers.indexOf('services');
  if (col < 0) return { ok: false, error: 'no services column' };
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(name)) {
      sh.getRange(i + 1, col + 1).setValue(services || '[]');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

function deleteClient(ss, name) {
  const sh = ss.getSheetByName('clients');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(name)) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

function useCode(ss, codeId, staffId, staffName) {
  const sh = ss.getSheetByName('codes');
  const headers = sheetHeaders(sh, 'codes');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(codeId)) {
      const usedIdx = headers.indexOf('used');
      const byIdx = headers.indexOf('usedBy');
      const byNameIdx = headers.indexOf('usedByName');
      const atIdx = headers.indexOf('usedAt');
      if (data[i][usedIdx] === true || data[i][usedIdx] === 'TRUE') {
        return { ok: false, error: 'already used' };
      }
      sh.getRange(i + 1, usedIdx + 1).setValue(true);
      sh.getRange(i + 1, byIdx + 1).setValue(staffId);
      sh.getRange(i + 1, byNameIdx + 1).setValue(staffName);
      sh.getRange(i + 1, atIdx + 1).setValue(new Date().toISOString());
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not found' };
}

function setConfig(ss, key, value) {
  const sh = ss.getSheetByName('config');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return { ok: true };
    }
  }
  sh.appendRow([key, value]);
  return { ok: true };
}

// ══════════════════════════════════════════════════════════════
//  ZOHO MAIL — send email from your Zoho address (server-side only)
//  Credentials live ONLY in Script Properties (Apps Script → Project
//  Settings → Script properties) — NEVER in the public frontend:
//    ZOHO_CLIENT_ID       from api-console.zoho.com (Self Client)
//    ZOHO_CLIENT_SECRET   from the same Self Client
//    ZOHO_REFRESH_TOKEN   generated once via the grant-token flow
//    ZOHO_FROM            your sending address, e.g. hr@yourdomain.com
//    ZOHO_DC  (optional)  data centre: com | eu | in | sa | com.au  (default com)
//  Needs the manifest scope https://www.googleapis.com/auth/script.external_request
// ══════════════════════════════════════════════════════════════
function zohoProp(k){ return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function zohoDC(){ return zohoProp('ZOHO_DC') || 'com'; }

function zohoAccessToken(){
  const res = UrlFetchApp.fetch('https://accounts.zoho.' + zohoDC() + '/oauth/v2/token', {
    method: 'post',
    payload: {
      refresh_token: zohoProp('ZOHO_REFRESH_TOKEN'),
      client_id: zohoProp('ZOHO_CLIENT_ID'),
      client_secret: zohoProp('ZOHO_CLIENT_SECRET'),
      grant_type: 'refresh_token'
    },
    muteHttpExceptions: true
  });
  const j = JSON.parse(res.getContentText());
  if (!j.access_token) throw new Error('Zoho token error: ' + res.getContentText());
  return j.access_token;
}

function zohoAccountId(token){
  const cached = zohoProp('ZOHO_ACCOUNT_ID');
  if (cached) return cached;
  const res = UrlFetchApp.fetch('https://mail.zoho.' + zohoDC() + '/api/accounts', {
    headers: { Authorization: 'Zoho-oauthtoken ' + token }, muteHttpExceptions: true
  });
  const j = JSON.parse(res.getContentText());
  const acc = (j.data && j.data[0]) ? j.data[0] : null;
  if (!acc) throw new Error('Zoho accounts error: ' + res.getContentText());
  const id = String(acc.accountId);
  PropertiesService.getScriptProperties().setProperty('ZOHO_ACCOUNT_ID', id);
  return id;
}

// Send an email from the configured Zoho address. Returns {ok:true} or throws.
function zohoSendMail(to, subject, htmlBody, cc){
  const token = zohoAccessToken();
  const accountId = zohoAccountId(token);
  const payload = {
    fromAddress: zohoProp('ZOHO_FROM'),
    toAddress: to,
    subject: subject || '',
    content: htmlBody || '',
    mailFormat: 'html'
  };
  if (cc) payload.ccAddress = cc;
  const res = UrlFetchApp.fetch('https://mail.zoho.' + zohoDC() + '/api/accounts/' + accountId + '/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Zoho-oauthtoken ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const txt = res.getContentText();
  let j; try { j = JSON.parse(txt); } catch (e) { throw new Error('Zoho send error: ' + txt); }
  if (j.status && j.status.code !== 200) throw new Error('Zoho send failed: ' + txt);
  return { ok: true };
}

// Run this manually from the editor to verify the connection. It emails your
// own ZOHO_FROM address — check that mailbox for "WorkLog test".
function zohoTest(){
  const r = zohoSendMail(zohoProp('ZOHO_FROM'), 'WorkLog test ✅', '<p>Zoho Mail is connected to your WorkLog app.</p>');
  Logger.log(JSON.stringify(r));
  return r;
}

// One-time: paste the grant code (from api-console.zoho.com Self Client) into a
// Script Property ZOHO_GRANT_CODE, then Run this once. It swaps the code for a
// long-lived refresh token and saves it — after that, delete ZOHO_GRANT_CODE.
function zohoExchangeGrant(){
  const code = zohoProp('ZOHO_GRANT_CODE');
  if (!code) throw new Error('Set ZOHO_GRANT_CODE in Script Properties first.');
  const res = UrlFetchApp.fetch('https://accounts.zoho.' + zohoDC() + '/oauth/v2/token', {
    method: 'post',
    payload: {
      grant_type: 'authorization_code',
      client_id: zohoProp('ZOHO_CLIENT_ID'),
      client_secret: zohoProp('ZOHO_CLIENT_SECRET'),
      code: code
    },
    muteHttpExceptions: true
  });
  const j = JSON.parse(res.getContentText());
  if (!j.refresh_token) throw new Error('No refresh token returned: ' + res.getContentText());
  const props = PropertiesService.getScriptProperties();
  props.setProperty('ZOHO_REFRESH_TOKEN', j.refresh_token);
  props.deleteProperty('ZOHO_GRANT_CODE');
  Logger.log('✅ Refresh token saved. Now run zohoTest().');
  return '✅ Refresh token saved. Now run zohoTest().';
}

// ══════════════════════════════════════════════════════════════
//  SALARY INCREMENT PROGRAM
//  Staff opt in, then qualify each month by hitting three bars:
//    1. >= 95% entry compliance
//    2. all their client deliverables ticked off
//    3. previous month's PDF report submitted by the 5th
// ══════════════════════════════════════════════════════════════

function incJoin(ss, params) {
  const sh = ss.getSheetByName('incMembers');
  const headers = sheetHeaders(sh, 'incMembers');
  const col = h => headers.indexOf(h);
  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col('staffId')]) === String(params.staffId)) {
      sh.getRange(i + 1, col('active') + 1).setValue(true);   // re-joining
      return { ok: true, rejoined: true };
    }
  }
  const obj = { id: 'inc' + Date.now(), staffId: params.staffId, staffName: params.staffName, joinedAt: today, active: true };
  const r = sh.getLastRow() + 1;
  headers.forEach((h, i) => {
    const cell = sh.getRange(r, i + 1);
    if (h === 'joinedAt') cell.setNumberFormat('@');
    cell.setValue(obj[h] !== undefined ? obj[h] : '');
  });
  return { ok: true, joined: true };
}

function incSetActive(ss, staffId, active) {
  const sh = ss.getSheetByName('incMembers');
  const headers = sheetHeaders(sh, 'incMembers');
  const col = h => headers.indexOf(h);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col('staffId')]) === String(staffId)) {
      sh.getRange(i + 1, col('active') + 1).setValue(active);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Not enrolled' };
}

// One row per (month, staff, client, service). Ticking is idempotent.
function setIncWork(ss, params) {
  const sh = ss.getSheetByName('incWork');
  const headers = sheetHeaders(sh, 'incWork');
  const col = h => headers.indexOf(h);
  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const done = params.done === 'true' || params.done === true;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col('month')]) === String(params.month) &&
        String(data[i][col('staffId')]) === String(params.staffId) &&
        String(data[i][col('client')]) === String(params.client) &&
        String(data[i][col('service')]) === String(params.service)) {
      sh.getRange(i + 1, col('done') + 1).setValue(done);
      sh.getRange(i + 1, col('doneAt') + 1).setValue(done ? today : '');
      if (params.note !== undefined) sh.getRange(i + 1, col('note') + 1).setValue(params.note || '');
      return { ok: true, updated: true };
    }
  }
  const obj = {
    id: 'w' + Date.now() + Math.floor(Math.random() * 1000),
    month: params.month, staffId: params.staffId, staffName: params.staffName,
    client: params.client, service: params.service,
    done: done, doneAt: done ? today : '', note: params.note || ''
  };
  const r = sh.getLastRow() + 1;
  headers.forEach((h, i) => {
    const cell = sh.getRange(r, i + 1);
    if (h === 'month' || h === 'doneAt') cell.setNumberFormat('@');
    cell.setValue(obj[h] !== undefined ? obj[h] : '');
  });
  return { ok: true, created: true };
}

// Monthly report PDF -> Drive. On time = submitted by the 5th of the month
// AFTER the reporting month.
function submitIncReport(ss, params) {
  const sh = ss.getSheetByName('incReports');
  const headers = sheetHeaders(sh, 'incReports');
  const col = h => headers.indexOf(h);
  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // deadline = 5th of the month following params.month ("2026-07" -> 2026-08-05)
  const y = parseInt(String(params.month).slice(0, 4), 10);
  const m = parseInt(String(params.month).slice(5, 7), 10);
  const dl = new Date(Date.UTC(y, m, 5));           // m is 0-based next month
  const deadline = dl.toISOString().slice(0, 10);
  const onTime = today <= deadline;

  const blob = Utilities.newBlob(Utilities.base64Decode(params.fileData), 'application/pdf',
                                 (params.fileName || ('report_' + params.month + '.pdf')));
  const file = uploadFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/file/d/' + file.getId() + '/view';

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col('month')]) === String(params.month) &&
        String(data[i][col('staffId')]) === String(params.staffId)) {
      sh.getRange(i + 1, col('fileUrl') + 1).setValue(url);
      sh.getRange(i + 1, col('fileName') + 1).setValue(params.fileName || '');
      sh.getRange(i + 1, col('submittedAt') + 1).setValue(today);
      sh.getRange(i + 1, col('onTime') + 1).setValue(onTime);
      return { ok: true, updated: true, onTime: onTime, url: url };
    }
  }
  const obj = {
    id: 'rep' + Date.now(), month: params.month, staffId: params.staffId, staffName: params.staffName,
    fileUrl: url, fileName: params.fileName || '', submittedAt: today, onTime: onTime
  };
  const r = sh.getLastRow() + 1;
  headers.forEach((h, i) => {
    const cell = sh.getRange(r, i + 1);
    if (h === 'month' || h === 'submittedAt') cell.setNumberFormat('@');
    cell.setValue(obj[h] !== undefined ? obj[h] : '');
  });
  return { ok: true, created: true, onTime: onTime, url: url };
}

// ══════════════════════════════════════════════════════════════
//  AUTO TASKS — on the 1st request of each month, every staff member gets a
//  task to upload the previous month's report (due the 5th, 7 PM Dubai).
//  Guarded by a config key so it can only ever run once per month.
// ══════════════════════════════════════════════════════════════
function ensureMonthlyReportTasks(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const month = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const cache = CacheService.getScriptCache();
  if (cache.get('rptTasks_' + month) === '1') return;

  const cfg = ss.getSheetByName('config');
  const data = cfg.getDataRange().getValues();
  let row = -1, done = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'reportTasksMonth') {
      row = i + 1;
      // Sheets may have stored "2026-08" as a real date — normalise before
      // comparing, otherwise the guard never matches and tasks duplicate.
      done = String(dateCellToString(data[i][1], tz)).slice(0, 7);
      break;
    }
  }
  if (done === month) { cache.put('rptTasks_' + month, '1', 21600); return; }

  // previous month, e.g. "July 2026"
  const d = new Date(); d.setDate(1); d.setDate(0);
  const prevName = Utilities.formatDate(d, tz, 'MMMM yyyy');
  const deadline = month + '-05T19:00';

  const tasksSh = ss.getSheetByName('tasks');
  const headers = sheetHeaders(tasksSh, 'tasks');
  const staff = sheetToObjects(ss.getSheetByName('staff'));
  staff.forEach(s => {
    const obj = {
      id: 't' + Date.now() + Math.floor(Math.random() * 100000),
      title: 'Monthly report — ' + prevName,
      description: 'Upload your ' + prevName + ' monthly report as a PDF, covering each client you handle. Submit it on your Increment page before the 5th.',
      assignedTo: s.id, assignedToName: s.name,
      createdByRole: 'exec', createdById: 'system', createdByName: 'Auto',
      deadline: deadline, done: false, doneAt: '', staffNote: '', imageUrl: '',
      createdAt: new Date().toISOString()
    };
    tasksSh.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
  });
  // Store as text so Sheets cannot reinterpret "2026-08" as a date.
  if (row < 0) { cfg.appendRow(['reportTasksMonth', '']); row = cfg.getLastRow(); }
  const cell = cfg.getRange(row, 2);
  cell.setNumberFormat('@');
  cell.setValue(month);
  cache.put('rptTasks_' + month, '1', 21600);
}


// ══════════════════════════════════════════════════════════════
//  AUTO-DELETE ENTRY PROOFS (v10)
//  Proof images are deleted from Drive 7 days after the entry's date.
//  The Sheet row is kept; rows[i].proofImg is cleared and a
//  proofDeletedAt note is added.
//  Run installProofCleanupTrigger() ONCE from the editor to schedule.
// ══════════════════════════════════════════════════════════════

const PROOF_RETENTION_DAYS = 7;

function installProofCleanupTrigger() {
  // Remove any existing trigger for this function first, so re-running
  // this doesn't create duplicates.
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'cleanupOldProofs') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('cleanupOldProofs')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
  return 'Daily proof-cleanup trigger installed.';
}

// Pulls the Drive file ID out of the URL format saveImageToDrive() produces:
// https://drive.google.com/uc?export=view&id=FILE_ID
function driveIdFromProofUrl(url) {
  const m = String(url || '').match(/[?&]id=([^&]+)/);
  return m ? m[1] : null;
}

function cleanupOldProofs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  assertRightSpreadsheet(ss);
  const sh = ss.getSheetByName('entries');
  const headers = sheetHeaders(sh, 'entries');
  const tz = ss.getSpreadsheetTimeZone();
  const rowsIdx = headers.indexOf('rows');
  const dateIdx = headers.indexOf('date');
  if (rowsIdx < 0) return;

  const data = sh.getDataRange().getValues();
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - PROOF_RETENTION_DAYS);

  let deletedCount = 0;
  const logSh = ensureProofLogSheet(ss);

  for (let i = 1; i < data.length; i++) {
    const entryDateRaw = dateCellToString(data[i][dateIdx], tz);
    const entryDate = new Date(entryDateRaw);
    if (isNaN(entryDate.getTime()) || entryDate >= cutoff) continue;

    let rows = [];
    try { rows = data[i][rowsIdx] ? JSON.parse(data[i][rowsIdx]) : []; } catch (e) { continue; }
    if (!rows.length) continue;

    let changed = false;
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    rows.forEach(r => {
      if (r && r.proofImg) {
        const fileId = driveIdFromProofUrl(r.proofImg);
        if (fileId) {
          try {
            DriveApp.getFileById(fileId).setTrashed(true);
            logSh.appendRow([new Date().toISOString(), data[i][0], fileId, r.proofImg]);
            deletedCount++;
          } catch (e) {
            // File already gone or inaccessible — still clear the reference.
          }
        }
        r.proofImg = '';
        r.proofDeletedAt = today;
        changed = true;
      }
    });

    if (changed) {
      sh.getRange(i + 1, rowsIdx + 1).setValue(JSON.stringify(rows));
    }
  }

  return { ok: true, deleted: deletedCount };
}

function ensureProofLogSheet(ss) {
  let sh = ss.getSheetByName('proofDeletionLog');
  if (!sh) {
    sh = ss.insertSheet('proofDeletionLog');
    sh.getRange(1, 1, 1, 4).setValues([['deletedAt', 'entryId', 'driveFileId', 'oldUrl']]);
    sh.setFrozenRows(1);
  }
  return sh;
}


// ══════════════════════════════════════════════════════════════
//  CONFIG HELPERS (v10)
//  Override any of these by adding a row to the 'config' tab:
//    offDay          0=Sunday, 1=Monday ... 6=Saturday   (default 0)
//    wAttendance     weight for attendance    (default 40)
//    wCompliance     weight for entry logs    (default 30)
//    wManagement     weight for exec score    (default 30)
//    entriesPerDay   expected entries per day (default 2)
// ══════════════════════════════════════════════════════════════

function cfgNum(ss, key, fallback) {
  const sh = ss.getSheetByName('config');
  if (!sh) return fallback;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      const n = parseFloat(data[i][1]);
      return isNaN(n) ? fallback : n;
    }
  }
  return fallback;
}


// ══════════════════════════════════════════════════════════════
//  ANNOUNCEMENTS (v10)
//  Posted by an executive. Goes live IMMEDIATELY and disappears
//  exactly 24 hours later.
// ══════════════════════════════════════════════════════════════

// Kept for reference — no longer used now that announcements go live at once.
function nextWorkingDay10(ss) {
  const offDay = cfgNum(ss, 'offDay', 0);
  const now = new Date();
  const d = new Date();
  d.setHours(10, 0, 0, 0);
  if (now.getTime() >= d.getTime()) d.setDate(d.getDate() + 1);
  let guard = 0;
  while (d.getDay() === offDay && guard++ < 10) {
    d.setDate(d.getDate() + 1);
  }
  return d;
}

function addAnnouncement(ss, params) {
  const sh = ss.getSheetByName('announcements');
  const headers = sheetHeaders(sh, 'announcements');
  const showFrom = new Date();                       // live the moment it is posted
  const showUntil = new Date(showFrom.getTime() + 24 * 60 * 60 * 1000);

  const obj = {
    id: 'ann' + Date.now(),
    title: params.title || '',
    body: params.body || '',
    postedById: params.postedById || '',
    postedByName: params.postedByName || '',
    postedAt: new Date().toISOString(),
    showFrom: showFrom.toISOString(),
    showUntil: showUntil.toISOString()
  };

  const r = sh.getLastRow() + 1;
  headers.forEach((h, i) => {
    sh.getRange(r, i + 1).setValue(obj[h] !== undefined ? obj[h] : '');
  });
  return { ok: true, showFrom: obj.showFrom, showUntil: obj.showUntil };
}

// Every announcement, flagged so the frontend can show staff only the
// live one while execs also see what is scheduled and what has expired.
function getAnnouncements(ss) {
  const sh = ss.getSheetByName('announcements');
  if (!sh) return [];
  const now = new Date().getTime();
  return sheetToObjects(sh).map(a => {
    const from = new Date(a.showFrom).getTime();
    const until = new Date(a.showUntil).getTime();
    return {
      ...a,
      active: !isNaN(from) && !isNaN(until) && now >= from && now < until,
      expired: !isNaN(until) && now >= until
    };
  });
}


// ══════════════════════════════════════════════════════════════
//  MANAGEMENT SCORE (v10) — executive rates each staff member 0-100
//  for a given month. One row per (month, staff); re-setting overwrites.
// ══════════════════════════════════════════════════════════════

function setMgmtScore(ss, params) {
  const sh = ss.getSheetByName('mgmtScores');
  const headers = sheetHeaders(sh, 'mgmtScores');
  const col = h => headers.indexOf(h);
  const score = Math.max(0, Math.min(100, parseFloat(params.score) || 0));
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col('month')]) === String(params.month) &&
        String(data[i][col('staffId')]) === String(params.staffId)) {
      sh.getRange(i + 1, col('score') + 1).setValue(score);
      sh.getRange(i + 1, col('note') + 1).setValue(params.note || '');
      sh.getRange(i + 1, col('setBy') + 1).setValue(params.setBy || '');
      sh.getRange(i + 1, col('setAt') + 1).setValue(new Date().toISOString());
      return { ok: true, updated: true };
    }
  }

  const obj = {
    id: 'ms' + Date.now(),
    month: params.month,
    staffId: params.staffId,
    staffName: params.staffName || '',
    score: score,
    note: params.note || '',
    setBy: params.setBy || '',
    setAt: new Date().toISOString()
  };
  const r = sh.getLastRow() + 1;
  headers.forEach((h, i) => {
    const cell = sh.getRange(r, i + 1);
    if (h === 'month') cell.setNumberFormat('@');
    cell.setValue(obj[h] !== undefined ? obj[h] : '');
  });
  return { ok: true, created: true };
}


// ══════════════════════════════════════════════════════════════
//  LEADERBOARD (v10)
//
//    ATTENDANCE   days checked in / working days elapsed
//    COMPLIANCE   entries submitted / expected entries
//                 (an entry flagged late counts as half)
//    MANAGEMENT   set by the executive, 0-100 (0 if unset)
//
//  Total = attendance*40% + compliance*30% + management*30%
//  month format: "2026-08". Omit to use the current month.
// ══════════════════════════════════════════════════════════════

function getLeaderboard(ss, month) {
  const tz = ss.getSpreadsheetTimeZone();
  const m = month || Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const offDay = cfgNum(ss, 'offDay', 0);
  const wA = cfgNum(ss, 'wAttendance', 40);
  const wC = cfgNum(ss, 'wCompliance', 30);
  const wM = cfgNum(ss, 'wManagement', 30);
  const perDay = cfgNum(ss, 'entriesPerDay', 2);
  const wTotal = wA + wC + wM || 1;

  // ---- working days elapsed so far this month ----
  const y = parseInt(m.slice(0, 4), 10);
  const mo = parseInt(m.slice(5, 7), 10);
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const isCurrentMonth = todayStr.slice(0, 7) === m;
  const lastDayOfMonth = new Date(y, mo, 0).getDate();
  const lastDay = isCurrentMonth ? parseInt(todayStr.slice(8, 10), 10) : lastDayOfMonth;

  let workingDays = 0;
  for (let d = 1; d <= lastDay; d++) {
    if (new Date(y, mo - 1, d).getDay() !== offDay) workingDays++;
  }
  if (workingDays < 1) workingDays = 1;

  // ---- raw data ----
  const staff = sheetToObjects(ss.getSheetByName('staff'));

  const attendance = sheetToObjects(ss.getSheetByName('attendance')).map(a => ({
    staffId: String(a.staffId),
    date: String(dateCellToString(a.date, tz)),
    checkIn: a.checkIn
  })).filter(a => a.date.slice(0, 7) === m && String(a.checkIn).trim() !== '');

  const entries = sheetToObjects(ss.getSheetByName('entries')).map(e => ({
    staffId: String(e.staffId),
    date: String(dateCellToString(e.date, tz)),
    late: e.late === true || e.late === 'TRUE' || e.late === 'true'
  })).filter(e => e.date.slice(0, 7) === m);

  const mgmt = {};
  const scoreSh = ss.getSheetByName('mgmtScores');
  if (scoreSh) {
    sheetToObjects(scoreSh).forEach(s => {
      if (String(s.month) === m) mgmt[String(s.staffId)] = parseFloat(s.score) || 0;
    });
  }

  // ---- score each staff member ----
  const rows = staff.map(s => {
    const sid = String(s.id);

    const days = {};
    attendance.forEach(a => { if (a.staffId === sid) days[a.date] = 1; });
    const daysPresent = Object.keys(days).length;
    const attScore = Math.min(100, (daysPresent / workingDays) * 100);

    let credit = 0, totalEntries = 0;
    entries.forEach(e => {
      if (e.staffId !== sid) return;
      totalEntries++;
      credit += e.late ? 0.5 : 1;
    });
    const expected = workingDays * perDay;
    const compScore = Math.min(100, (credit / expected) * 100);

    const mgmtScore = mgmt[sid] !== undefined ? mgmt[sid] : 0;
    const total = (attScore * wA + compScore * wC + mgmtScore * wM) / wTotal;

    return {
      staffId: sid,
      staffName: s.name,
      attendance: Math.round(attScore * 10) / 10,
      compliance: Math.round(compScore * 10) / 10,
      management: Math.round(mgmtScore * 10) / 10,
      total: Math.round(total * 10) / 10,
      daysPresent: daysPresent,
      entriesLogged: totalEntries
    };
  });

  rows.sort((a, b) => b.total - a.total);
  rows.forEach((r, i) => { r.rank = i + 1; });

  return {
    month: m,
    workingDays: workingDays,
    expectedEntries: workingDays * perDay,
    weights: { attendance: wA, compliance: wC, management: wM },
    rows: rows
  };
}
