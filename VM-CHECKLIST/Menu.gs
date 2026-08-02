// Menu.gs
// Builds the "VM Tools" menu and its supporting actions, including
// summary helpers and test data tools used during development.

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('VM Tools')
    .addItem('⚙️ Setup System (first time only)', 'setup')
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('📊 Today\'s Quick Summary', 'showTodaySummary')
    .addItem('🕒 View Stores Not Visited This Month', 'showPendingStores')
    .addSeparator()
    .addItem('🧪 Seed 15 Sample Visits (testing)', 'seedTestData')
    .addItem('🗑️ Reset (clear all responses + test data)', 'clearAllResponses')
    .addSeparator()
    .addItem('🖼️ Apply/Refresh Logo', 'applyLogo')
    .addItem('🔗 Get Form Link', 'getFormLink')
    .addItem('📋 Rebuild Checklist_Master', 'rebuildChecklistMaster')
    .addToUi();
}

function clearAllResponses() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Clear all responses?',
    'This permanently deletes every row in Master Log, Issue Log, Item Score Log, and Form Responses (headers kept), then refreshes the dashboard. Continue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.MASTER));
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.ISSUES));
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.SCORES));
  ss.getSheets().forEach(function (s) {
    if (s.getName().indexOf('Form Responses') === 0) clearSheetKeepHeader_(s);
  });
  refreshDashboard();
  ui.alert('Cleared. Dashboard refreshed.');
}

// Generates dummy visits with randomized scores (including some
// deliberate low scores) so the dashboard has data to render during testing.
function seedTestData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(SHEETS.MASTER);
  var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
  var scoreSheet = ss.getSheetByName(SHEETS.SCORES);
  var today = new Date();
  var count = 15;
  var issueRows = [], scoreRows = [];

  for (var i = 0; i < count; i++) {
    var d = new Date(today.getFullYear(), today.getMonth(), 1 + Math.floor(Math.random() * 27));
    var store = CONFIG.STORES[Math.floor(Math.random() * CONFIG.STORES.length)];
    var sum = 0, max = 0, rated = 0;

    SECTIONS.forEach(function (sec) {
      sec.items.forEach(function (item) {
        if (Math.random() < 0.1) return; // occasionally Not Applicable
        var score = 1 + Math.floor(Math.random() * 5);
        sum += score; max += 5; rated += 1;
        scoreRows.push([new Date(), d, store, 'Test VM ' + ((i % 3) + 1), sec.name, item, score]);
        if (score <= CONFIG.LOW_SCORE_THRESHOLD) {
          issueRows.push([new Date(), d, store, 'Test VM ' + ((i % 3) + 1), sec.name, item, score, 'Open', '', '', '', '']);
        }
      });
    });

    var adherencePct = max > 0 ? Math.round(sum / max * 100) : '';
    master.appendRow([new Date(), d, store, 'Test SM ' + ((i % 3) + 1), 'Test VM ' + ((i % 3) + 1), 'Store Visit',
      rated, sum, max, adherencePct, i % 4 === 0 ? 'Sample section remark' : '', i % 5 === 0 ? 'Sample overall remark' : '']);
  }
  if (issueRows.length) issueSheet.getRange(issueSheet.getLastRow() + 1, 1, issueRows.length, issueRows[0].length).setValues(issueRows);
  if (scoreRows.length) scoreSheet.getRange(scoreSheet.getLastRow() + 1, 1, scoreRows.length, scoreRows[0].length).setValues(scoreRows);
  refreshDashboard();
  SpreadsheetApp.getUi().alert('Seeded ' + count + ' sample visits and refreshed the dashboard.');
}

function showTodaySummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(SHEETS.MASTER);
  var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
  var today = dayKey_(new Date());
  var mRows = (master && master.getLastRow() > 1) ? master.getRange(2, 1, master.getLastRow() - 1, 12).getValues() : [];
  var iRows = (issueSheet && issueSheet.getLastRow() > 1) ? issueSheet.getRange(2, 1, issueSheet.getLastRow() - 1, 12).getValues() : [];
  var todayRows = mRows.filter(function (r) { return dayKey_(r[1]) === today; });
  var visitedStores = uniqueCount_(todayRows.map(function (r) { return r[2]; }));
  var avgToday = avgOf_(todayRows, 9);
  var openIssues = iRows.filter(function (r) { return r[7] !== 'Closed'; }).length;

  SpreadsheetApp.getUi().alert('Today\'s Summary',
    'Stores visited today: ' + visitedStores + ' / ' + CONFIG.STORES.length +
    '\nAvg adherence % today: ' + avgToday + '%' +
    '\nOpen issues (all-time unresolved): ' + openIssues,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function showPendingStores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(SHEETS.MASTER);
  var monthPrefix = monthKey_(new Date());
  var mRows = (master && master.getLastRow() > 1) ? master.getRange(2, 1, master.getLastRow() - 1, 12).getValues() : [];
  var visitedThisMonth = mRows.filter(function (r) { return monthKey_(r[1]) === monthPrefix; }).map(function (r) { return r[2]; });
  var pending = CONFIG.STORES.filter(function (s) { return visitedThisMonth.indexOf(s) === -1; });

  SpreadsheetApp.getUi().alert('Stores Not Visited This Month',
    pending.length ? pending.join(', ') : 'None — every store has been visited this month! 🎉',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function rebuildChecklistMaster() {
  buildChecklistMaster_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Checklist_Master rebuilt from Config.gs.');
}

function getFormLink() {
  var id = PropertiesService.getDocumentProperties().getProperty('FORM_ID');
  if (!id) { Logger.log('No form found yet. Run setup() first.'); return; }
  var form = FormApp.openById(id);
  var msg = 'Fill link (share with VMs):\n' + form.getPublishedUrl() +
    '\n\nEdit link (keep for yourself only):\n' + form.getEditUrl();
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) {}
}
