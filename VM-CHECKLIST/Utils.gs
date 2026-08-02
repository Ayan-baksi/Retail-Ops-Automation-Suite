// Utils.gs
// Shared helper functions used across multiple files.

function dayKey_(d) { return Utilities.formatDate(new Date(d), TZ, 'yyyy-MM-dd'); }
function monthKey_(d) { return Utilities.formatDate(new Date(d), TZ, 'yyyy-MM'); }

function uniqueCount_(arr) {
  var s = {};
  arr.forEach(function (v) { if (v) s[v] = true; });
  return Object.keys(s).length;
}

function clearSheetKeepHeader_(sh) {
  if (!sh) return;
  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
}

function colorStatusCol_(sh, startRow, n, col) {
  for (var i = 0; i < n; i++) {
    var cell = sh.getRange(startRow + i, col), v = String(cell.getValue());
    if (v.indexOf('🔴') === 0) cell.setBackground('#fde2e1');
    else if (v.indexOf('🟡') === 0) cell.setBackground('#fff4d6');
    else if (v.indexOf('🟢') === 0) cell.setBackground('#e3f4e1');
  }
}

// Reads a value from the Settings sheet (Setting | Value columns),
// falling back to a default if the sheet or row doesn't exist.
function getSetting_(name, fallback) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.SETTINGS);
  if (!sh) return fallback;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === name) return data[i][1];
  }
  return fallback;
}
