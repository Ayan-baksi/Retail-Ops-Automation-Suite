// Menu.gs
// Builds the "SM Tools" menu and its supporting actions, including
// data-repair tools, summary helpers, and test data tools used during development.

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SM Tools')
    .addItem('⚙️ Setup System (first time only)', 'setup')
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('📊 Today\'s Quick Summary', 'showTodaySummary')
    .addItem('🕒 View Pending Stores Today', 'showPendingStores')
    .addSeparator()
    .addItem('🧪 Seed 15 Sample Audits (testing)', 'seedTestData')
    .addItem('🗑️ Reset (clear all responses + test data)', 'clearAllResponses')
    .addSeparator()
    .addItem('🖼️ Apply/Refresh Logo', 'applyLogo')
    .addItem('🔗 Get Form Link', 'getFormLink')
    .addItem('📋 Rebuild Checklist_Master', 'rebuildChecklistMaster')
    .addItem('📸 Rebuild Photo Log from All Responses', 'rebuildPhotoLog')
    .addItem('🔍 Audit: Find Non-Required Questions', 'auditFormRequiredStatus')
    .addItem('🗑️ Delete One Visit (by Timestamp)', 'deleteVisitByTimestamp')
    .addToUi();
}

// Removes one specific visit's rows from Master Log, Issue Log, and
// Photo Log together, matched by Timestamp. Manual row deletion in a
// single sheet leaves orphaned rows in the other two, so this keeps
// all three in sync.
function deleteVisitByTimestamp() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Delete one visit',
    'Paste the exact Timestamp from Master Log column A (e.g. copy the cell value) for the visit you want removed. This will delete matching rows from Master Log, Issue Log, and Photo Log together.',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var tsInput = resp.getResponseText().trim();
  if (!tsInput) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var removed = { 'Master Log': 0, 'Issue Log': 0, 'Photo Log': 0 };

  [SHEETS.MASTER, SHEETS.ISSUES, SHEETS.PHOTOS].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues(); // Timestamp column only
    for (var i = data.length - 1; i >= 0; i--) {
      var cellTs = data[i][0];
      var cellTsStr = cellTs instanceof Date ? cellTs.toString() : String(cellTs);
      if (cellTsStr === tsInput || cellTsStr.indexOf(tsInput) === 0) {
        sh.deleteRow(i + 2);
        removed[name]++;
      }
    }
  });

  refreshDashboard();
  ui.alert('Removed rows — Master Log: ' + removed['Master Log'] +
    ', Issue Log: ' + removed['Issue Log'] +
    ', Photo Log: ' + removed['Photo Log'] + '.\nDashboard refreshed.');
}

// Scans the live form and reports every question not marked required.
// Useful when submissions are missing mandatory items or photos — often
// caused by a question losing its "Required" toggle during manual edits.
function auditFormRequiredStatus() {
  var ui = SpreadsheetApp.getUi();
  var id = PropertiesService.getDocumentProperties().getProperty('FORM_ID');
  if (!id) { ui.alert('No form found yet. Run setup() first.'); return; }

  var form = FormApp.openById(id);
  var items = form.getItems();
  var notRequired = [];

  items.forEach(function (item) {
    var type = item.getType();
    var required = null;
    try {
      if (type === FormApp.ItemType.FILE_UPLOAD) required = item.asFileUploadItem().isRequired();
      else if (type === FormApp.ItemType.MULTIPLE_CHOICE) required = item.asMultipleChoiceItem().isRequired();
      else if (type === FormApp.ItemType.TEXT) required = item.asTextItem().isRequired();
      else if (type === FormApp.ItemType.LIST) required = item.asListItem().isRequired();
      else if (type === FormApp.ItemType.DATE) required = item.asDateItem().isRequired();
    } catch (err) { return; } // page breaks, section headers, etc. have no "required" concept

    if (required === false) notRequired.push(item.getTitle() + '  [' + type + ']');
  });

  var msg = notRequired.length
    ? 'Found ' + notRequired.length + ' question(s) that are NOT required:\n\n' + notRequired.join('\n')
    : 'Every question in the form is correctly marked required. The gap is likely something else — send me the exact stores/dates and I\'ll dig further.';
  Logger.log(msg);
  ui.alert('Required-Status Audit', msg, ui.ButtonSet.OK);
}

// Re-scans every historical form response and rebuilds Photo Log from
// scratch using the current detection logic. Only touches Photo Log.
// Use this if photos are missing for some stores — those responses were
// likely submitted while an older detection method was still in use.
function rebuildPhotoLog() {
  var ui = SpreadsheetApp.getUi();
  var id = PropertiesService.getDocumentProperties().getProperty('FORM_ID');
  if (!id) { ui.alert('No form found yet. Run setup() first.'); return; }

  var form = FormApp.openById(id);
  var responses = form.getResponses();
  var itemTextToSection = buildItemSectionLookup_();
  var allPhotoRows = [];

  responses.forEach(function (r) {
    var ts = r.getTimestamp();
    var sm = '', store = '', visitDate = null, auditType = '';
    var lastItemTitle = '', lastItemSection = '';

    r.getItemResponses().forEach(function (ir) {
      var item = ir.getItem(), title = item.getTitle(), ans = ir.getResponse();
      if (title === Q.SM) { sm = ans; return; }
      if (title === Q.STORE) { store = ans; return; }
      if (title === Q.DATE) { visitDate = new Date(ans); return; }
      if (title === Q.AUDIT_TYPE) { auditType = ans; return; }
      if (item.getType() === FormApp.ItemType.FILE_UPLOAD) {
        var fileIds = ans || [];
        fileIds.forEach(function (fileId) {
          allPhotoRows.push([ts, visitDate || ts, store, sm, auditType, lastItemSection, lastItemTitle,
            'https://drive.google.com/file/d/' + fileId + '/view']);
        });
        return;
      }
      if (itemTextToSection[title]) { lastItemTitle = title; lastItemSection = itemTextToSection[title]; }
    });
  });

  var photoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PHOTOS);
  clearSheetKeepHeader_(photoSheet);
  if (allPhotoRows.length) {
    photoSheet.getRange(2, 1, allPhotoRows.length, allPhotoRows[0].length).setValues(allPhotoRows);
  }
  refreshDashboard();
  ui.alert('Rebuilt Photo Log from ' + responses.length + ' total form responses.\nFound ' + allPhotoRows.length + ' photos.');
}

// Quick glance at today's numbers without opening the Dashboard tab.
function showTodaySummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(SHEETS.MASTER);
  var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
  var today = dayKey_(new Date());

  var mRows = (master && master.getLastRow() > 1) ? master.getRange(2, 1, master.getLastRow() - 1, 12).getValues() : [];
  var iRows = (issueSheet && issueSheet.getLastRow() > 1) ? issueSheet.getRange(2, 1, issueSheet.getLastRow() - 1, 12).getValues() : [];
  var todayRows = mRows.filter(function (r) { return dayKey_(r[1]) === today; });
  var openingDone = uniqueCount_(todayRows.filter(function (r) { return r[4] === 'Opening'; }).map(function (r) { return r[2]; }));
  var closingDone = uniqueCount_(todayRows.filter(function (r) { return r[4] === 'Closing'; }).map(function (r) { return r[2]; }));
  var openIssues = iRows.filter(function (r) { return r[7] !== 'Closed'; }).length;
  var avgScoreToday = avgOf_(todayRows, 9);

  SpreadsheetApp.getUi().alert('Today\'s Summary',
    'Opening completed: ' + openingDone + ' / ' + CONFIG.STORES.length +
    '\nClosing completed: ' + closingDone + ' / ' + CONFIG.STORES.length +
    '\nAvg score today: ' + avgScoreToday + '%' +
    '\nOpen issues (all-time unresolved): ' + openIssues,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// Lists which stores still haven't done Opening/Closing today.
function showPendingStores() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(SHEETS.MASTER);
  var today = dayKey_(new Date());
  var mRows = (master && master.getLastRow() > 1) ? master.getRange(2, 1, master.getLastRow() - 1, 12).getValues() : [];
  var todayRows = mRows.filter(function (r) { return dayKey_(r[1]) === today; });
  var openedStores = todayRows.filter(function (r) { return r[4] === 'Opening'; }).map(function (r) { return r[2]; });
  var closedStores = todayRows.filter(function (r) { return r[4] === 'Closing'; }).map(function (r) { return r[2]; });
  var pendingOpening = CONFIG.STORES.filter(function (s) { return openedStores.indexOf(s) === -1; });
  var pendingClosing = CONFIG.STORES.filter(function (s) { return closedStores.indexOf(s) === -1; });

  SpreadsheetApp.getUi().alert('Pending Stores Today',
    'Pending Opening (' + pendingOpening.length + '): ' + (pendingOpening.join(', ') || 'None — all done! 🎉') +
    '\n\nPending Closing (' + pendingClosing.length + '): ' + (pendingClosing.join(', ') || 'None — all done! 🎉'),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// Re-syncs Checklist_Master with the current contents of Config.gs.
function rebuildChecklistMaster() {
  buildChecklistMaster_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('Checklist_Master rebuilt from Config.gs.');
}

function clearAllResponses() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Clear all responses?',
    'This permanently deletes every row in Master Log, Issue Log, Photo Log, and Form Responses (headers kept), then refreshes the dashboard. Continue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.MASTER));
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.ISSUES));
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.PHOTOS));
  ss.getSheets().forEach(function (s) {
    if (s.getName().indexOf('Form Responses') === 0) clearSheetKeepHeader_(s);
  });
  refreshDashboard();
  ui.alert('Cleared. Dashboard refreshed.');
}

function seedTestData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(SHEETS.MASTER);
  var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
  var today = new Date();
  var count = 15;
  var issueRows = [];
  var openingTotal = SECTIONS_OPENING.reduce(function (s, sec) { return s + sec.items.length; }, 0);
  var closingTotal = SECTIONS_CLOSING.reduce(function (s, sec) { return s + sec.items.length; }, 0);

  for (var i = 0; i < count; i++) {
    var d = new Date(today.getFullYear(), today.getMonth(), 1 + Math.floor(Math.random() * 27));
    var store = CONFIG.STORES[Math.floor(Math.random() * CONFIG.STORES.length)];
    var type = i % 2 === 0 ? 'Opening' : 'Closing';
    var total = type === 'Opening' ? openingTotal : closingTotal;
    var no = Math.floor(Math.random() * 5);
    var na = Math.floor(Math.random() * 2);
    var yes = total - no - na;
    var scorePct = (yes + no) > 0 ? Math.round(yes / (yes + no) * 100) : '';

    master.appendRow([new Date(), d, store, 'Test SM ' + ((i % 3) + 1), type,
      total, yes, no, na, scorePct, i % 4 === 0 ? 'Sample section remark' : '', i % 5 === 0 ? 'Sample overall remark' : '']);

    var sections = type === 'Opening' ? SECTIONS_OPENING : SECTIONS_CLOSING;
    for (var z = 0; z < no; z++) {
      var sec = sections[Math.floor(Math.random() * sections.length)];
      var item = sec.items[Math.floor(Math.random() * sec.items.length)];
      issueRows.push([new Date(), d, store, 'Test SM ' + ((i % 3) + 1), type, sec.name, item.text, 'Open', '', '', '', '']);
    }
  }
  if (issueRows.length) issueSheet.getRange(issueSheet.getLastRow() + 1, 1, issueRows.length, issueRows[0].length).setValues(issueRows);
  refreshDashboard();
  SpreadsheetApp.getUi().alert('Seeded ' + count + ' sample audits and refreshed the dashboard.');
}

function getFormLink() {
  var id = PropertiesService.getDocumentProperties().getProperty('FORM_ID');
  if (!id) { Logger.log('No form found yet. Run setup() first.'); return; }
  var form = FormApp.openById(id);
  var msg = 'Fill link (share with Store Managers):\n' + form.getPublishedUrl() +
    '\n\nEdit link (keep for yourself only):\n' + form.getEditUrl();
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) {}
}
