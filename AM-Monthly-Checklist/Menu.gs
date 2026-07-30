// Menu.gs
// Builds the "AM Tools" menu and its supporting actions, including the
// test data helpers used during development.

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AM Tools')
    .addItem('⚙️ Setup System (first time only)', 'setup')
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('🧪 Seed 15 Sample Visits (testing)', 'seedTestData')
    .addItem('🗑️ Reset (clear all responses + test data)', 'clearAllResponses')
    .addSeparator()
    .addItem('🖼️ Apply/Refresh Logo', 'applyLogo')
    .addItem('🔗 Get Form Link', 'getFormLink')
    .addToUi();
}

function clearAllResponses() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Clear all responses?',
    'This permanently deletes every row in Master Log, Issue Log, and Form Responses (headers kept), then refreshes the dashboard. Continue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.MASTER));
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.ISSUES));
  ss.getSheets().forEach(function (s) {
    if (s.getName().indexOf('Form Responses') === 0) clearSheetKeepHeader_(s);
  });
  refreshDashboard();
  ui.alert('Cleared. Dashboard refreshed.');
}

// Generates dummy visits, including some deliberate failures, so every
// KPI card, chart, ranking, and issue list has data to render during testing.
function seedTestData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(SHEETS.MASTER);
  var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
  var today = new Date();
  var count = 15;
  var issueRows = [];
  var totalItems = SECTIONS.reduce(function (s, sec) { return s + sec.items.length; }, 0);

  for (var i = 0; i < count; i++) {
    var d = new Date(today.getFullYear(), today.getMonth(), 1 + Math.floor(Math.random() * 27));
    var store = CONFIG.STORES[Math.floor(Math.random() * CONFIG.STORES.length)];
    var zeros = Math.floor(Math.random() * 12);
    var ones = totalItems - zeros;
    var scorePct = Math.round(ones / totalItems * 100);

    master.appendRow([new Date(), d, store, 'Test AM ' + ((i % 3) + 1), 'Cluster ' + ((i % 2) + 1),
      totalItems, ones, zeros, scorePct,
      i % 4 === 0 ? 'Sample section remark' : '', i % 5 === 0 ? 'Sample overall remark' : '']);

    for (var z = 0; z < zeros; z++) {
      var sec = SECTIONS[Math.floor(Math.random() * SECTIONS.length)];
      var item = sec.items[Math.floor(Math.random() * sec.items.length)];
      issueRows.push([new Date(), d, store, 'Test AM ' + ((i % 3) + 1), sec.name, item, 'Open', '', '', '', '']);
    }
  }
  if (issueRows.length) {
    issueSheet.getRange(issueSheet.getLastRow() + 1, 1, issueRows.length, issueRows[0].length).setValues(issueRows);
  }
  refreshDashboard();
  SpreadsheetApp.getUi().alert('Seeded ' + count + ' sample visits and refreshed the dashboard.');
}

function getFormLink() {
  var id = PropertiesService.getDocumentProperties().getProperty('FORM_ID');
  if (!id) { Logger.log('No form found yet. Run setup() first.'); return; }
  var form = FormApp.openById(id);
  var msg = 'Fill link (share with Area Managers):\n' + form.getPublishedUrl() +
    '\n\nEdit link (keep for yourself only):\n' + form.getEditUrl();
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) { /* expected when run from the script editor */ }
}
