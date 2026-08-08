// SubmitHandler.gs
// Runs on every Form submission. Parses the response — works for either
// the Opening or Closing branch automatically, since it reads whichever
// questions are actually present. Writes one row to Master Log, one row
// per "No" answer to Issue Log, and one row per uploaded photo to Photo Log.

function onFormSubmitHandler(e) {
  var r = e.response, ts = r.getTimestamp();
  var sm = '', store = '', visitDate = null, auditType = '', overallRemarks = '';
  var remarksBySection = {};
  var yes = 0, no = 0, na = 0;
  var failedRows = []; // feeds Issue Log
  var photoRows = [];  // feeds Photo Log

  var itemTextToSection = buildItemSectionLookup_();

  r.getItemResponses().forEach(function (ir) {
    var item = ir.getItem(), title = item.getTitle(), ans = ir.getResponse();

    if (title === Q.SM) { sm = ans; return; }
    if (title === Q.STORE) { store = ans; return; }
    if (title === Q.DATE) { visitDate = new Date(ans); return; }
    if (title === Q.AUDIT_TYPE) { auditType = ans; return; }
    if (title === Q.FINAL_REMARKS) { overallRemarks = ans || ''; return; }
    if (title.indexOf('— Remarks') > -1) {
      if (ans) remarksBySection[title.replace(' — Remarks (optional)', '')] = ans;
      return;
    }
    if (title.indexOf(' — Photo Link') > -1) {
      var baseTitle = title.replace(/ — Photo Link.*$/, '');
      var section = itemTextToSection[baseTitle] || '';
      if (ans) photoRows.push([ts, visitDate, store, sm, auditType, section, baseTitle, ans]);
      return;
    }
    // Otherwise: a Yes / No / Not Applicable checklist question
    if (ans === 'Yes') yes++;
    else if (ans === 'No') {
      no++;
      failedRows.push([itemTextToSection[title] || '', title]);
    } else if (ans === 'Not Applicable') na++;
  });

  if (!visitDate) visitDate = ts;
  var scorePct = (yes + no) > 0 ? Math.round(yes / (yes + no) * 100) : '';
  var sectionRemarksCombined = Object.keys(remarksBySection).map(function (k) {
    return k + ': ' + remarksBySection[k];
  }).join(' | ');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(SHEETS.MASTER).appendRow([
    ts, visitDate, store, sm, auditType, yes + no + na, yes, no, na, scorePct,
    sectionRemarksCombined, overallRemarks
  ]);

  if (failedRows.length) {
    var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
    var issueRows = failedRows.map(function (f) {
      return [ts, visitDate, store, sm, auditType, f[0], f[1], 'Open', '', '', '', ''];
    });
    issueSheet.getRange(issueSheet.getLastRow() + 1, 1, issueRows.length, issueRows[0].length).setValues(issueRows);
  }

  if (photoRows.length) {
    var photoSheet = ss.getSheetByName(SHEETS.PHOTOS);
    photoSheet.getRange(photoSheet.getLastRow() + 1, 1, photoRows.length, photoRows[0].length).setValues(photoRows);
  }

  refreshDashboard();
}

// Maps item text to section name across both Opening and Closing, so
// Issue Log / Photo Log rows know which section an item belongs to
// regardless of which branch the responder took.
function buildItemSectionLookup_() {
  var map = {};
  SECTIONS_OPENING.concat(SECTIONS_CLOSING).forEach(function (sec) {
    sec.items.forEach(function (item) { map[item.text] = sec.name; });
  });
  return map;
}

function buildMasterLogSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.MASTER) || ss.insertSheet(SHEETS.MASTER);
  sh.clear();
  var h = ['Timestamp', 'Date', 'Store', 'Store Manager Name', 'Opening/Closing',
    'Total Questions', 'Yes Count', 'No Count', 'N/A Count', 'Score %', 'Section Remarks', 'Overall Remarks'];
  sh.getRange(1, 1, 1, h.length).setValues([h])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.setColumnWidth(1, 140); sh.setColumnWidth(4, 150); sh.setColumnWidth(11, 260); sh.setColumnWidth(12, 260);
}

function buildIssueLogSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.ISSUES) || ss.insertSheet(SHEETS.ISSUES);
  sh.clear();
  var h = ['Timestamp', 'Date', 'Store', 'Store Manager', 'Opening/Closing', 'Section', 'Item (Answered No)',
    'Status', 'Assigned To', 'Due Date', 'Closed Date', 'Severity'];
  sh.getRange(1, 1, 1, h.length).setValues([h])
    .setFontWeight('bold').setBackground('#7f1d1d').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.setColumnWidth(4, 150); sh.setColumnWidth(6, 170); sh.setColumnWidth(7, 380);
}

function buildPhotoLogSheet_(ss) {
  var sh = ss.getSheetByName(SHEETS.PHOTOS) || ss.insertSheet(SHEETS.PHOTOS);
  sh.clear();
  var h = ['Timestamp', 'Date', 'Store', 'Store Manager', 'Opening/Closing', 'Section', 'Item', 'Photo URL'];
  sh.getRange(1, 1, 1, h.length).setValues([h])
    .setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
  sh.setColumnWidth(4, 150); sh.setColumnWidth(6, 170); sh.setColumnWidth(7, 320); sh.setColumnWidth(8, 300);
}

function hideFormResponsesSheet_(ss) {
  ss.getSheets().forEach(function (s) {
    if (s.getName().indexOf('Form Responses') === 0) {
      try { s.hideSheet(); } catch (err) { /* not possible if it's the only visible sheet */ }
    }
  });
}
