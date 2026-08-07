/* ═══════════════════════════════════════════════════════════════
   ██  KREW — REPAIR the staff sheet header and the rows it broke
   ██  Paste as a NEW file in the Krew Apps Script project.
   ██  Run repairCheck() first (changes nothing), then repairApply().
   ██  Delete this file afterwards.
   ══════════════════════════════════════════════════════════════
   What happened: cell A1 of the "staff" tab was overwritten from
   "id" to " an", so the backend stopped seeing the id column and
   appended an empty one. Every staff member then read as id = "",
   which made them share one identity — same logs, same attendance.
   ═══════════════════════════════════════════════════════════════ */

var EXPECTED_STAFF = ['id','name','username','password','createdAt'];

function repairCheck() { run_(true); }
function repairApply() { run_(false); }

function run_(dryRun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(dryRun ? '=== DRY RUN — nothing will be written ===' : '=== APPLYING REPAIRS ===');
  Logger.log('Spreadsheet: "%s"', ss.getName());

  var sh = ss.getSheetByName('staff');
  if (!sh) { Logger.log('STOP — no "staff" tab. Wrong spreadsheet?'); return; }

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  Logger.log('staff headers now: %s', JSON.stringify(headers));

  // ── 1. find the column that really holds the ids ──────────────
  // It is the one whose values look like s1778587409190, whatever its
  // header currently says.
  var lastRow = sh.getLastRow();
  var idCol = -1;
  for (var c = 0; c < headers.length; c++) {
    if (lastRow < 2) break;
    var vals = sh.getRange(2, c + 1, lastRow - 1, 1).getValues().map(function (r) { return String(r[0]); });
    var looksLikeIds = vals.length && vals.every(function (v) { return /^s\d{6,}$/.test(v); });
    if (looksLikeIds) { idCol = c; break; }
  }
  if (idCol < 0) { Logger.log('STOP — could not find a column of staff ids. Nothing changed.'); return; }
  Logger.log('id values live in column %s (header is "%s")', idCol + 1, headers[idCol]);

  // ── 2. any OTHER column called "id" that is completely empty is the
  //       duplicate the backend appended — it must go. ─────────────
  var junkCols = [];
  for (var c2 = 0; c2 < headers.length; c2++) {
    if (c2 === idCol) continue;
    if (headers[c2].trim() !== 'id') continue;
    var empty = lastRow < 2 || sh.getRange(2, c2 + 1, lastRow - 1, 1).getValues()
      .every(function (r) { return String(r[0]).trim() === ''; });
    if (empty) junkCols.push(c2);
  }

  Logger.log('PLAN: rename column %s header "%s" -> "id"', idCol + 1, headers[idCol]);
  Logger.log('PLAN: delete %s empty duplicate "id" column(s): %s',
             junkCols.length, junkCols.map(function (c) { return c + 1; }).join(', ') || '(none)');

  if (!dryRun) {
    sh.getRange(1, idCol + 1).setValue('id');
    // Delete right-to-left so earlier indexes stay valid.
    junkCols.sort(function (a, b) { return b - a; }).forEach(function (c) { sh.deleteColumn(c + 1); });
    Logger.log('staff headers after: %s',
      JSON.stringify(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]));
  }

  // ── 3. rebuild name -> id and backfill every row that lost its id ──
  var staff = sheetToObjects(ss.getSheetByName('staff'));
  var byName = {};
  staff.forEach(function (s) { if (s.name) byName[String(s.name).trim()] = s.id; });
  Logger.log('staff map: %s', JSON.stringify(byName));

  ['entries', 'attendance', 'incMembers', 'incWork', 'incReports', 'leads'].forEach(function (tab) {
    var t = ss.getSheetByName(tab);
    if (!t) return;
    var h = t.getRange(1, 1, 1, t.getLastColumn()).getValues()[0].map(String);
    var ci = h.indexOf('staffId'), cn = h.indexOf('staffName');
    if (ci < 0 || cn < 0) return;
    var rows = t.getLastRow();
    var fixed = 0, unmatched = [];
    for (var r = 2; r <= rows; r++) {
      var cur = String(t.getRange(r, ci + 1).getValue()).trim();
      if (cur !== '') continue;
      var nm = String(t.getRange(r, cn + 1).getValue()).trim();
      var id = byName[nm];
      if (!id) { if (nm) unmatched.push(nm); continue; }
      if (!dryRun) { var cell = t.getRange(r, ci + 1); cell.setNumberFormat('@'); cell.setValue(id); }
      fixed++;
    }
    Logger.log('%s: %s row(s) with a blank staffId %s%s', tab, fixed,
      dryRun ? 'would be repaired' : 'repaired',
      unmatched.length ? ' | no match for: ' + unmatched.join(', ') : '');
  });

  Logger.log(dryRun
    ? 'Dry run complete. Run repairApply() to write these changes.'
    : 'Repairs written. Staff should reload the app.');
}
