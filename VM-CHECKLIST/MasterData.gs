// MasterData.gs
// Builds Store_Master, Employee_Master, Checklist_Master, and Settings.

function buildAllMasterData_(ss) {
  buildStoreMaster_(ss);
  buildEmployeeMaster_(ss);
  buildChecklistMaster_(ss);
  buildSettings_(ss);
}

function buildStoreMaster_(ss) {
  var sh = ss.getSheetByName(SHEETS.STORE_MASTER) || ss.insertSheet(SHEETS.STORE_MASTER);
  if (sh.getLastRow() > 0) return;
  var h = ['Store Code', 'Store Name', 'Cluster', 'Region', 'State', 'Store Manager', 'Assigned VM', 'Status'];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  var rows = CONFIG.STORES.map(function (code) { return [code, code, '', '', '', '', '', 'Active']; });
  sh.getRange(2, 1, rows.length, h.length).setValues(rows);
  sh.setColumnWidth(1, 100); sh.setColumnWidth(2, 160);
}

function buildEmployeeMaster_(ss) {
  var sh = ss.getSheetByName(SHEETS.EMPLOYEE_MASTER) || ss.insertSheet(SHEETS.EMPLOYEE_MASTER);
  if (sh.getLastRow() > 0) return;
  var h = ['Name', 'Email', 'Role', 'Assigned Store(s)'];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  var dv = SpreadsheetApp.newDataValidation().requireValueInList(['VM', 'SM', 'RM'], true).setAllowInvalid(false).build();
  sh.getRange('C2:C200').setDataValidation(dv);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 200);
}

function buildChecklistMaster_(ss) {
  var sh = ss.getSheetByName(SHEETS.CHECKLIST_MASTER) || ss.insertSheet(SHEETS.CHECKLIST_MASTER);
  sh.clear();
  var h = ['Section', 'Item', 'Max Score', 'Weight', 'Active'];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  var rows = [];
  SECTIONS.forEach(function (sec) {
    sec.items.forEach(function (item) { rows.push([sec.name, item, 5, 1, 'Y']); });
  });
  sh.getRange(2, 1, rows.length, h.length).setValues(rows);
  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 420);
}

function buildSettings_(ss) {
  var sh = ss.getSheetByName(SHEETS.SETTINGS) || ss.insertSheet(SHEETS.SETTINGS);
  if (sh.getLastRow() > 0) return;
  var h = ['Setting', 'Value'];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  var rows = [
    ['Org Name', CONFIG.ORG_NAME],
    ['App Version', CONFIG.APP_VERSION],
    ['Overdue Days Threshold', CONFIG.OVERDUE_DAYS_DEFAULT],
    ['Low Score Threshold', CONFIG.LOW_SCORE_THRESHOLD]
  ];
  sh.getRange(2, 1, rows.length, h.length).setValues(rows);
  sh.setColumnWidth(1, 200);
}
