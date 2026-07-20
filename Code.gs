/**
 * WorkLog — Google Sheets Backend (v9 — entry proof screenshots + manager reviews)
 * Deploy as Web App: Execute as "Me", Access "Anyone"
 *
 * v8 added Tasks (with photo upload to Drive); v9 adds:
 *  - addEntry/updateEntry accept a `proofs` param: [{i, imageData, imageType}]
 *    Each image is saved to Drive and its URL stored as rows[i].proofImg.
 *  - reviewEntry action: manager approval/rejection stored inside rows[0]._review.
 */

const SHEETS = {
  staff:    ['id', 'name', 'username', 'password', 'createdAt'],
  managers: ['id', 'name', 'username', 'password', 'createdAt'],
  clients:  ['name'],
  entries:  ['id', 'staffId', 'staffName', 'date', 'gnotes', 'rows', 'late', 'reason', 'approved', 'submittedAt', 'usedCodeId', 'editedAt', 'entry', 'entryLabel', 'deadline'],
  codes:    ['id', 'code', 'staffId', 'staffName', 'note', 'issuedByRole', 'issuedById', 'issuedByName', 'generatedAt', 'generatedDate', 'generatedTime', 'used', 'usedBy', 'usedByName', 'usedAt'],
  reports:  ['id', 'weekStart', 'weekEnd', 'submittedAt', 'submittedBy', 'submittedById', 'breakdown', 'note', 'totalHours'],
  tasks:    ['id', 'title', 'description', 'assignedTo', 'assignedToName', 'createdByRole', 'createdById', 'createdByName', 'deadline', 'done', 'doneAt', 'staffNote', 'imageUrl', 'createdAt'],
  leads:    ['id', 'targetId', 'staffId', 'staffName', 'f1', 'f2', 'f3', 'success', 'createdAt'],
  config:   ['key', 'value']
};

const DEFAULT_CLIENTS = ['Baaqat Flowers','Flovia Flowers','8th Cafe','Florens Flowers','Flat Chocolate','Al Rasa','Fedora Perfumes','Elite Party','Hair Salon'];
const DEFAULT_EXEC_CODE = '99999';
const UPLOAD_FOLDER = 'WorkLog Uploads';

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    const params = (e.parameter || {});
    const action = params.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheets(ss);

    let result;
    switch (action) {
      case 'getAll':       result = getAll(ss); break;
      case 'addStaff':     result = addRow(ss, 'staff', JSON.parse(params.data)); break;
      case 'deleteStaff':  result = deleteRow(ss, 'staff', params.id); break;
      case 'resetStaff':   result = updateField(ss, 'staff', params.id, 'password', params.password); break;
      case 'addManager':   result = addRow(ss, 'managers', JSON.parse(params.data)); break;
      case 'deleteManager':result = deleteRow(ss, 'managers', params.id); break;
      case 'resetManager': result = updateField(ss, 'managers', params.id, 'password', params.password); break;
      case 'addClient':    result = addClient(ss, params.name); break;
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
      case 'deleteLead':   result = deleteRow(ss, 'leads', params.id); break;
      case 'setExecCode':  result = setConfig(ss, 'execCode', params.value); break;
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

function getAll(ss) {
  const tz = ss.getSpreadsheetTimeZone();
  const staff = sheetToObjects(ss.getSheetByName('staff')).map(s => ({ ...s, createdAt: dateCellToString(s.createdAt, tz) }));
  const managers = sheetToObjects(ss.getSheetByName('managers')).map(m => ({ ...m, createdAt: dateCellToString(m.createdAt, tz) }));
  const clients = sheetToObjects(ss.getSheetByName('clients')).map(c => c.name);
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
  const config = sheetToObjects(ss.getSheetByName('config'));
  const execCode = (config.find(c => c.key === 'execCode') || {}).value || DEFAULT_EXEC_CODE;
  return { staff, managers, clients, entries, codes, reports, tasks, leads, execCode: String(execCode) };
}

function addRow(ss, sheetName, obj) {
  const sh = ss.getSheetByName(sheetName);
  const headers = sheetHeaders(sh, sheetName);
  const row = headers.map(h => {
    let v = obj[h];
    if ((h === 'rows' || h === 'breakdown') && (Array.isArray(v) || typeof v === 'object')) v = JSON.stringify(v);
    if (v === undefined || v === null) v = '';
    return v;
  });
  sh.appendRow(row);
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
  return addRow(ss, 'entries', obj);
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

function addClient(ss, name) {
  const sh = ss.getSheetByName('clients');
  const existing = sheetToObjects(sh).map(c => c.name);
  if (existing.includes(name)) return { ok: false, error: 'exists' };
  sh.appendRow([name]);
  return { ok: true };
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
