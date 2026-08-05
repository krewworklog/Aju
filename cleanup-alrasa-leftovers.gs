/* ═══════════════════════════════════════════════════════════════
   ██  ONE-OFF CLEANUP — run from the Krew WorkLog Apps Script editor
   ██  Removes the Al Rasa tabs that were created in the Krew database
   ██  by mistake. Run checkLeftovers() first, then deleteLeftovers().
   ██  Delete this file afterwards; it is not part of the app.
   ═══════════════════════════════════════════════════════════════ */

// Tabs the Krew app owns. This list is the safety net: nothing in it can be
// deleted, no matter what else this script says.
var KREW_TABS = ['staff','managers','clients','entries','codes','reports',
  'tasks','leads','attendance','incMembers','deliverables','incWork',
  'incReports','config'];

// Tabs that only Al Rasa ever creates. 'config' is deliberately absent —
// both apps use a tab of that name and Krew's must survive.
var ALRASA_TABS = ['users','sales','invoices','invoiceLog','collection','cheques'];

function checkLeftovers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Spreadsheet: "%s"', ss.getName());

  // Refuse to run against the wrong file.
  if (!ss.getSheetByName('entries')) {
    Logger.log('STOP — no "entries" tab. This is not the Krew database. Nothing checked.');
    return;
  }

  Logger.log('--- every tab in this spreadsheet ---');
  ss.getSheets().forEach(function (sh) {
    var rows = Math.max(0, sh.getLastRow() - 1);   // minus the header row
    var owner = KREW_TABS.indexOf(sh.getName()) >= 0 ? 'KREW — keep'
              : (ALRASA_TABS.indexOf(sh.getName()) >= 0 ? 'AL RASA — will be deleted'
              : 'unknown — will be KEPT, delete by hand if you want it gone');
    Logger.log('%s  (%s data rows)  [%s]', sh.getName(), rows, owner);
  });

  Logger.log('--- Krew config keys (these must stay) ---');
  var cfg = ss.getSheetByName('config');
  if (cfg) {
    cfg.getDataRange().getValues().forEach(function (r) {
      if (String(r[0]) && String(r[0]) !== 'key') Logger.log('%s = %s', r[0], r[1]);
    });
  }
  Logger.log('Nothing has been deleted. Run deleteLeftovers() when the list above looks right.');
}

function deleteLeftovers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName('entries')) {
    Logger.log('STOP — no "entries" tab. This is not the Krew database. Nothing deleted.');
    return;
  }

  var removed = [];
  ALRASA_TABS.forEach(function (name) {
    if (KREW_TABS.indexOf(name) >= 0) return;      // belt and braces
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    ss.deleteSheet(sh);
    removed.push(name);
  });

  // The Al Rasa code also writes its own settings into the shared config tab.
  var cfg = ss.getSheetByName('config');
  var cfgRemoved = [];
  if (cfg) {
    var vals = cfg.getDataRange().getValues();
    for (var i = vals.length - 1; i >= 1; i--) {
      var k = String(vals[i][0]);
      if (k === 'openingBalance' || k === 'openingDate') {
        cfg.deleteRow(i + 1);
        cfgRemoved.push(k);
      }
    }
  }

  Logger.log('Deleted tabs: %s', removed.length ? removed.join(', ') : '(none found)');
  Logger.log('Deleted config rows: %s', cfgRemoved.length ? cfgRemoved.join(', ') : '(none found)');
  Logger.log('Remaining tabs: %s', ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
}
