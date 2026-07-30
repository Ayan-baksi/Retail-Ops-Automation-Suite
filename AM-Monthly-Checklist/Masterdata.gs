// MasterData.gs
// Builds the reference "master" sheets: Store, Employee, Checklist, and
// Settings. These support future Cluster/Region/RM/VM filtering,
// checklist item weighting, and non-technical settings changes. The
// filter/weighting logic in Dashboard.gs does not consume this data yet,
// but the data model is in place for when it does.

function buildAllMasterData_(ss) {
  buildStoreMaster_(ss);
  buildEmployeeMaster_(ss);
  buildChecklistMaster_(ss);
  buildSettings_(ss);
}

// One row per store. Cluster/Region/RM/VM start blank for manual fill-in.
function buildStoreMaster_(ss) {
  var sh = ss.getSheetByName(SHEETS.STORE_MASTER) || ss.insertSheet(SHEETS.STORE_MASTER);
  var h = ['Store Code', 'Store Name', 'Cluster', 'Region', 'State',
    'Area Manager', 'Regional Manager', 'Visual Merchandiser', 'Status'];
  if (sh.getLastRow() > 0) return; // don't overwrite existing data
  sh.getRange(1, 1, 1, h.length).setValues([h])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  var rows = CONFIG.STORES.map(function (code) {
    return [code, code, '', '', '', '', '', '', 'Active'];
  });
  sh.getRange(2, 1, rows.length, h.length).setValues(rows);
  sh.setColumnWidth(1, 100); sh.setColumnWidth(2, 160);
  sh.getRange('A1').setNote('Store Name defaults to the store code — replace with the real store name whenever convenient.');
}

// Blank template — one row per Area/Regional Manager or Visual Merchandiser.
function buildEmployeeMaster_(ss) {
  var sh = ss.getSheetByName(SHEETS.EMPLOYEE_MASTER) || ss.insertSheet(SHEETS.EMPLOYEE_MASTER);
  if (sh.getLastRow() > 0) return;
  var h = ['Name', 'Email', 'Role', 'Assigned Cluster', 'Assigned Region'];
  sh.getRange(1, 1, 1, h.length).setValues([h])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  var dv = SpreadsheetApp.newDataValidation()
    .requireValueInList(['AM', 'RM', 'VM', 'MAR'], true).setAllowInvalid(false).build();
  sh.getRange('C2:C200').setDataValidation(dv);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 200);
}

// One row per checklist item, auto-populated from SECTIONS in Config.gs.
// Weight defaults to 1 (equal weighting) for future use.
function buildChecklistMaster_(ss) {
  var sh = ss.getSheetByName(SHEETS.CHECKLIST_MASTER) || ss.insertSheet(SHEETS.CHECKLIST_MASTER);
  sh.clear(); // safe to rebuild — mirrors Config.gs, no manual data lives here
  var h = ['Section', 'Item', 'Weight', 'Active'];
  sh.getRange(1, 1, 1, h.length).setValues([h])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  var rows = [];
  SECTIONS.forEach(function (sec) {
    sec.items.forEach(function (item) { rows.push([sec.name, item, 1, 'Y']); });
  });
  sh.getRange(2, 1, rows.length, h.length).setValues(rows);
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 420);
  sh.getRange('A1').setNote('This mirrors SECTIONS in Config.gs. Editing Weight here has no effect yet — weighting is not wired into the dashboard until enabled.');
}

function buildSettings_(ss) {
  var sh = ss.getSheetByName(SHEETS.SETTINGS) || ss.insertSheet(SHEETS.SETTINGS);
  if (sh.getLastRow() > 0) return;
  var h = ['Setting', 'Value'];
  sh.getRange(1, 1, 1, h.length).setValues([h])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  var rows = [
    ['Org Name', CONFIG.ORG_NAME],
    ['App Version', CONFIG.APP_VERSION],
    ['Overdue Days Threshold', CONFIG.OVERDUE_DAYS_DEFAULT],
    ['Target Score %', 90]
  ];
  sh.getRange(2, 1, rows.length, h.length).setValues(rows);
  sh.setColumnWidth(1, 200);
  sh.getRange('A1').setNote('Change values here — Dashboard reads Overdue Days Threshold live via getSetting_().');
}
