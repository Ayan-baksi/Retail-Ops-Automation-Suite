// SubmitHandler.gs
// Parses VM Form submissions into Master Log (one row per visit, with
// adherence % computed the same way as the source workbook's Summary
// sheet) and Issue Log (one row per item scored at or below the Low
// Score Threshold).

function onFormSubmitHandler(e) {
  var r = e.response, ts = r.getTimestamp();
  var vm = '', store = '', sm = '', visitDate = null, activity = '', overallRemarks = '';
  var remarksBySection = {};
  var sumScore = 0, maxPossible = 0, totalRated = 0;
  var lowScoreRows = [];

  var itemTextToSection = buildItemSectionLookup_();
  var lowThreshold = Number(getSetting_('Low Score Threshold', CONFIG.LOW_SCORE_THRESHOLD));
  var allScoreRows = []; // feeds Item Score Log (every item, for per-area rollups)

  r.getItemResponses().forEach(function (ir) {
    var item = ir.getItem(), title = item.getTitle(), ans = ir.getResponse();

    if (title === Q.VM) { vm = ans; return; }
    if (title === Q.STORE) { store = ans; return; }
    if (title === Q.SM) { sm = ans; return; }
    if (title === Q.DATE) { visitDate = new Date(ans); return; }
    if (title === Q.ACTIVITY) { activity = ans; return; }
    if (title === Q.FINAL_REMARKS) { overallRemarks = ans || ''; return; }
    if (title.indexOf('— Remarks') > -1) {
      if (ans) remarksBySection[title.replace(' — Remarks (optional)', '')] = ans;
      return;
    }
    // Otherwise: a checklist score item (1-5 or Not Applicable)
    if (ans === 'Not Applicable') return;
    var score = Number(ans);
    sumScore += score;
    maxPossible += 5;
    totalRated += 1;
    var itemSection = itemTextToSection[title] || '';
    allScoreRows.push([ts, visitDate, store, vm, itemSection, title, score]);
    if (score <= lowThreshold) {
      lowScoreRows.push([itemSection, title, score]);
    }
  });

  if (!visitDate) visitDate = ts;
  var adherencePct = maxPossible > 0 ? Math.round(sumScore / maxPossible * 100) : '';
  var sectionRemarksCombined = Object.keys(remarksBySection).map(function (k) {
    return k + ': ' + remarksBySection[k];
  }).join(' | ');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(SHEETS.MASTER).appendRow([
    ts, visitDate, store, sm, vm, activity, totalRated, sumScore, maxPossible, adherencePct,
    sectionRemarksCombined, overallRemarks
  ]);

  if (lowScoreRows.length) {
    var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
    var issueRows = lowScoreRows.map(function (f) {
      return [ts, visitDate, store, vm, f[0], f[1], f[2], 'Open', '', '', '', ''];
    });
    issueSheet.getRange(issueSheet.getLastRow() + 1, 1, issueRows.length, issueRows[0].length).setValues(issueRows);
  }

  if (allScoreRows.length) {
    var scoreSheet = ss.getSheetByName(SHEETS.SCORES);
    scoreSheet.getRange(scoreSheet.getLastRow() + 1, 1, allScoreRows.length, allScoreRows[0].length).setValues(allScoreRows);
  }

  refreshDashboard();
}

function buildItemSectionLookup_() {
  var map = {};
  SECTIONS.forEach(function (sec) {
    sec.items.forEach(function (item) { map[item] = sec.name; });
  });
  return map;
}

function buildMasterLogSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.MASTER) || ss.insertSheet(SHEETS.MASTER);
  sh.clear();
  var h = ['Timestamp', 'Date', 'Store', 'Store Manager Name', 'VM Name', 'Activity',
    'Items Rated', 'Sum Score', 'Max Possible', 'Adherence %', 'Section Remarks', 'Overall Remarks'];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.setColumnWidth(1, 140); sh.setColumnWidth(4, 150); sh.setColumnWidth(5, 130); sh.setColumnWidth(11, 260); sh.setColumnWidth(12, 260);
}

function buildIssueLogSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.ISSUES) || ss.insertSheet(SHEETS.ISSUES);
  sh.clear();
  var h = ['Timestamp', 'Date', 'Store', 'VM Name', 'Section', 'Item (Low Score)', 'Score',
    'Status', 'Assigned To', 'Due Date', 'Closed Date', 'Severity'];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#7f1d1d').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 220); sh.setColumnWidth(6, 380);
}

function buildItemScoreLogSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.SCORES) || ss.insertSheet(SHEETS.SCORES);
  sh.clear();
  var h = ['Timestamp', 'Date', 'Store', 'VM Name', 'Section', 'Item', 'Score'];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 220); sh.setColumnWidth(6, 400);
}

function hideFormResponsesSheet_(ss) {
  ss.getSheets().forEach(function (s) {
    if (s.getName().indexOf('Form Responses') === 0) {
      try { s.hideSheet(); } catch (err) {}
    }
  });
}
