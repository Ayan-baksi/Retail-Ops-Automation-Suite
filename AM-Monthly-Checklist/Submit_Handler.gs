// SubmitHandler.gs
// Runs on every Form submission. Parses the response, writes one row to
// Master Log per visit, and zero or more rows to Issue Log (one per
// failed item). Also builds those two sheets, since their structure is
// defined entirely by what gets written here.

function onFormSubmitHandler(e) {
  var r = e.response, ts = r.getTimestamp();
  var am = '', cluster = '', store = '', visitDate = null, overallRemarks = '';
  var remarksBySection = {};
  var ones = 0, zeros = 0;
  var failedRows = [];

  r.getItemResponses().forEach(function (ir) {
    var item = ir.getItem(), title = item.getTitle(), ans = ir.getResponse();
    if (title === Q.AM) am = ans;
    else if (title === Q.CLUSTER) cluster = ans;
    else if (title === Q.STORE) store = ans;
    else if (title === Q.DATE) visitDate = new Date(ans);
    else if (title === Q.FINAL_REMARKS) overallRemarks = ans || '';
    else if (title.indexOf('— Remarks') > -1) {
      if (ans) remarksBySection[title.replace(' — Remarks (optional)', '')] = ans;
    } else if (item.getType() === FormApp.ItemType.GRID) {
      var rows = item.asGridItem().getRows();
      var sectionName = title.replace(' — checklist', '');
      for (var i = 0; i < rows.length; i++) {
        var a = ans ? ans[i] : null;
        if (a === '1') ones++;
        else if (a === '0') { zeros++; failedRows.push([sectionName, rows[i]]); }
      }
    }
  });

  if (!visitDate) visitDate = ts;
  var scorePct = (ones + zeros) > 0 ? Math.round(ones / (ones + zeros) * 100) : '';
  var sectionRemarksCombined = Object.keys(remarksBySection).map(function (k) {
    return k + ': ' + remarksBySection[k];
  }).join(' | ');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(SHEETS.MASTER).appendRow([
    ts, visitDate, store, am, cluster, ones + zeros, ones, zeros, scorePct,
    sectionRemarksCombined, overallRemarks
  ]);

  if (failedRows.length) {
    var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
    var issueRows = failedRows.map(function (f) {
      // Status, Assigned To, Due Date, Closed Date, Severity are left
      // blank — reserved for a future issue-resolution workflow.
      return [ts, visitDate, store, am, f[0], f[1], 'Open', '', '', '', ''];
    });
    issueSheet.getRange(issueSheet.getLastRow() + 1, 1, issueRows.length, issueRows[0].length).setValues(issueRows);
  }

  refreshDashboard();
}

function buildMasterLogSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.MASTER) || ss.insertSheet(SHEETS.MASTER);
  sh.clear();
  var h = ['Timestamp', 'Date', 'Store', 'AM Name', 'Cluster', 'Items Scored',
    'Score = 1', 'Score = 0', 'Score %', 'Section Remarks', 'Overall Remarks'];
  sh.getRange(1, 1, 1, h.length).setValues([h])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 100); sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 130); sh.setColumnWidth(10, 260); sh.setColumnWidth(11, 260);
}

// One row per failed item. Includes Status/Assigned To/Due Date/Closed
// Date/Severity columns, left blank for now and reserved for a future
// issue-resolution workflow.
function buildIssueLogSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.ISSUES) || ss.insertSheet(SHEETS.ISSUES);
  sh.clear();
  var h = ['Timestamp', 'Date', 'Store', 'AM Name', 'Section', 'Item (Scored 0)',
    'Status', 'Assigned To', 'Due Date', 'Closed Date', 'Severity'];
  sh.getRange(1, 1, 1, h.length).setValues([h])
    .setFontWeight('bold').setBackground('#7f1d1d').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 100); sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 170); sh.setColumnWidth(6, 380);
}

function hideFormResponsesSheet_(ss) {
  ss.getSheets().forEach(function (s) {
    if (s.getName().indexOf('Form Responses') === 0) {
      try { s.hideSheet(); } catch (err) { /* not possible if it's the only visible sheet */ }
    }
  });
}
