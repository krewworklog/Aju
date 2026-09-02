/* ═══════════════════════════════════════════════════════════════
   ██  ONE-OFF — delete June + July 2026 work-log entries
   ██  Paste as a NEW file in the Krew Apps Script project.
   ██  Run previewDelete() first (changes nothing), then runDelete().
   ██  Delete this file afterwards.
   ══════════════════════════════════════════════════════════════
   Why a script and not the app: deleting 259 rows one HTTP request
   at a time would burn ~30 minutes of Apps Script runtime and trip
   the daily quota. This does it in a single execution.
   A backup of these entries is on the Desktop:
     krew-entries-june-july-2026-backup.csv / .json
   ═══════════════════════════════════════════════════════════════ */

var DELETE_MONTHS = ['2026-06', '2026-07'];

function previewDelete() { runJuneJuly_(true); }
function runDelete()     { runJuneJuly_(false); }

function runJuneJuly_(dryRun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(dryRun ? '=== PREVIEW — nothing will be deleted ===' : '=== DELETING ===');
  Logger.log('Spreadsheet: "%s"', ss.getName());

  var sh = ss.getSheetByName('entries');
  if (!sh) { Logger.log('STOP — no "entries" tab. Wrong spreadsheet.'); return; }

  var tz = ss.getSpreadsheetTimeZone();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var dateCol = headers.indexOf('date');
  var nameCol = headers.indexOf('staffName');
  if (dateCol < 0) { Logger.log('STOP — no "date" column.'); return; }

  var data = sh.getDataRange().getValues();
  var doomed = [];                       // row numbers, 1-based
  var byMonth = {}, byStaff = {};

  for (var i = 1; i < data.length; i++) {
    var ds = String(dateCellToString(data[i][dateCol], tz)).slice(0, 7);
    if (DELETE_MONTHS.indexOf(ds) === -1) continue;
    doomed.push(i + 1);
    byMonth[ds] = (byMonth[ds] || 0) + 1;
    var who = nameCol >= 0 ? String(data[i][nameCol]) : '?';
    byStaff[who] = (byStaff[who] || 0) + 1;
  }

  Logger.log('Rows in entries (excluding header): %s', data.length - 1);
  Logger.log('Matching %s: %s rows', DELETE_MONTHS.join(' + '), doomed.length);
  Object.keys(byMonth).sort().forEach(function (m) { Logger.log('   %s : %s', m, byMonth[m]); });
  Object.keys(byStaff).forEach(function (n) { Logger.log('   %s : %s', n, byStaff[n]); });
  Logger.log('Rows that will remain: %s', (data.length - 1) - doomed.length);

  if (dryRun) {
    Logger.log('Preview only. Run runDelete() to actually remove these.');
    return;
  }

  // Delete from the bottom up so earlier row numbers stay valid.
  doomed.sort(function (a, b) { return b - a; });
  for (var j = 0; j < doomed.length; j++) sh.deleteRow(doomed[j]);

  Logger.log('Deleted %s rows. Entries remaining: %s', doomed.length, sh.getLastRow() - 1);
  Logger.log('Done. Staff should reload the app.');
}
