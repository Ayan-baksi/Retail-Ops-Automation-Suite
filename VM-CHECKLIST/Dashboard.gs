// Dashboard.gs
// VM live dashboard, following the same patterns as the AM/SM systems
// (content-based pairing, checkbox-driven controls, self-healing
// checkbox sweep, dynamic Export-to-PDF), adapted for the 1-5 adherence
// % scoring model and the source workbook's per-area rollup structure.

var MAIN_COL = 4;
var PAIR_GAP = 1;
var STAGING_COL = 40;

function buildDashboardShell_(ss) {
  var sh = ss.getSheetByName(SHEETS.DASH) || ss.insertSheet(SHEETS.DASH);
  sh.clear();
  ss.setActiveSheet(sh); ss.moveActiveSheet(1);
}

function panelBorder_(range) {
  range.setBorder(true, true, true, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
}

function statusPill_(range, text, color) {
  range.setValue(text).setBackground(color + '22').setFontColor(color).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function tableHeader_(dash, row, col, headers, tint) {
  var bg = tint === 'danger' ? '#fef2f2' : '#f1f5f9';
  var fg = tint === 'danger' ? '#991b1b' : '#334155';
  dash.getRange(row, col, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground(bg).setFontColor(fg);
}

// Clears stray checkbox validations left behind by previous refreshes
// (see the equivalent function in the AM system's Dashboard.gs for details).
function clearAllCheckboxValidations_(dash) {
  var lastKnownRow = Number(PropertiesService.getDocumentProperties().getProperty('EXPORT_BTN_ROW')) || 300;
  var maxRow = Math.min(lastKnownRow + 20, dash.getMaxRows());
  var maxCol = Math.min(STAGING_COL + 10, dash.getMaxColumns());
  var rules = dash.getRange(1, 1, Math.max(maxRow, 1), Math.max(maxCol, 1)).getDataValidations();
  for (var i = 0; i < rules.length; i++) {
    for (var j = 0; j < rules[i].length; j++) {
      var rule = rules[i][j];
      if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
        dash.getRange(i + 1, j + 1).clearDataValidations();
      }
    }
  }
}

function onEdit(e) {
  var sh = e.range.getSheet();
  if (sh.getName() !== SHEETS.DASH) return;
  var a1 = e.range.getA1Notation();

  if (a1 === REFRESH_CHECKBOX_CELL) {
    if (e.range.getValue() === true) { refreshDashboard(); sh.getRange(REFRESH_CHECKBOX_CELL).setValue(false); }
    return;
  }
  if (a1 === FILTER_CELL || a1 === MONTH_FILTER_CELL) { refreshDashboard(); return; }

  var props = PropertiesService.getDocumentProperties();
  var exRow = Number(props.getProperty('EXPORT_BTN_ROW'));
  var exCol = Number(props.getProperty('EXPORT_BTN_COL'));
  if (exRow && e.range.getRow() === exRow && e.range.getColumn() === exCol) {
    if (e.range.getValue() === true) { sh.getRange(exRow, exCol).setValue(false); exportDashboardToPdf(); }
  }
}

// Distinct yyyy-MM values from Master Log, newest first. The current
// month is always included even if it has no data yet.
function getAvailableMonths_() {
  var master = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MASTER);
  var set = {};
  set[monthKey_(new Date())] = true;
  if (master && master.getLastRow() > 1) {
    var dates = master.getRange(2, 2, master.getLastRow() - 1, 1).getValues();
    dates.forEach(function (row) { if (row[0]) set[monthKey_(row[0])] = true; });
  }
  return Object.keys(set).sort().reverse();
}

function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(SHEETS.DASH) || ss.insertSheet(SHEETS.DASH);
  var master = ss.getSheetByName(SHEETS.MASTER);
  var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
  var scoreSheet = ss.getSheetByName(SHEETS.SCORES);

  var filterVal = 'All Stores', monthVal = monthKey_(new Date());
  try { filterVal = dash.getRange(FILTER_CELL).getValue() || 'All Stores'; } catch (err) {}
  try { monthVal = dash.getRange(MONTH_FILTER_CELL).getValue() || monthVal; } catch (err) {}

  dash.getCharts().forEach(function (c) { dash.removeChart(c); });
  dash.setConditionalFormatRules([]);
  dash.clear();
  dash.setHiddenGridlines(true);
  clearAllCheckboxValidations_(dash);

  var now = new Date();
  var today = dayKey_(now);

  var mRowsAll = (master && master.getLastRow() > 1) ? master.getRange(2, 1, master.getLastRow() - 1, 12).getValues() : [];
  // Columns: 0 ts, 1 date, 2 store, 3 sm, 4 vm, 5 activity, 6 rated, 7 sum, 8 max, 9 adherencePct, 10 sectionRemarks, 11 overallRemarks
  var iRowsAll = (issueSheet && issueSheet.getLastRow() > 1) ? issueSheet.getRange(2, 1, issueSheet.getLastRow() - 1, 12).getValues() : [];
  // Columns: 0 ts, 1 date, 2 store, 3 vm, 4 section, 5 item, 6 score, 7 status, 8 assignedTo, 9 dueDate, 10 closedDate, 11 severity
  var sRowsAll = (scoreSheet && scoreSheet.getLastRow() > 1) ? scoreSheet.getRange(2, 1, scoreSheet.getLastRow() - 1, 7).getValues() : [];
  // Columns: 0 ts, 1 date, 2 store, 3 vm, 4 section, 5 item, 6 score

  var byStore = function (rows) { return filterVal !== 'All Stores' ? rows.filter(function (r) { return r[2] === filterVal; }) : rows; };
  var mRows = byStore(mRowsAll), iRows = byStore(iRowsAll), sRows = byStore(sRowsAll);
  var storesInScope = filterVal === 'All Stores' ? CONFIG.STORES : [filterVal];

  var selectedMonthRows = mRows.filter(function (r) { return monthKey_(r[1]) === monthVal; });
  var monthIssues = iRows.filter(function (r) { return monthKey_(r[1]) === monthVal; });
  var monthScores = sRows.filter(function (r) { return monthKey_(r[1]) === monthVal; });
  var todayRows = mRows.filter(function (r) { return dayKey_(r[1]) === today; });
  var openIssuesAll = iRows.filter(function (r) { return r[7] !== 'Closed'; });

  var lastMonthPrefix = monthKey_(new Date(Number(monthVal.slice(0, 4)), Number(monthVal.slice(5, 7)) - 2, 1));
  var lastMonthRows = mRows.filter(function (r) { return monthKey_(r[1]) === lastMonthPrefix; });

  // --- Header ---
  var headerLastCol = MAIN_COL + 4 * 3 - 1;
  for (var c = 1; c <= headerLastCol; c++) { dash.getRange(1, c).setBackground('#0f172a'); dash.getRange(2, c).setBackground('#0f172a'); }
  dash.getRange(1, 3, 2, 2).merge();
  dash.getRange(1, 5, 1, 5).merge().setValue('VM AUDIT — LIVE DASHBOARD')
    .setFontColor('#ffffff').setFontWeight('bold').setFontSize(16).setVerticalAlignment('middle');
  dash.getRange(2, 5, 1, 5).merge().setValue('Real-time overview of store visual merchandising adherence')
    .setFontColor('#94a3b8').setFontSize(9).setVerticalAlignment('middle');
  dash.getRange(1, 10, 1, headerLastCol - 9).merge()
    .setValue('Last Updated: ' + Utilities.formatDate(now, TZ, 'dd MMM yyyy, hh:mm a'))
    .setFontColor('#94a3b8').setFontSize(9).setHorizontalAlignment('right').setVerticalAlignment('middle');
  dash.getRange(2, 10, 1, headerLastCol - 12).merge()
    .setValue('🔄  Refresh Data').setBackground('#ffffff').setFontColor('#0f172a')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.getRange(REFRESH_CHECKBOX_CELL).insertCheckboxes().setValue(false).setBackground('#ffffff');
  dash.setRowHeight(1, 38); dash.setRowHeight(2, 28);

  // --- Sidebar ---
  dash.getRange('A4:B4').merge().setValue('STORE').setFontWeight('bold').setFontSize(9).setFontColor('#64748b');
  var storeDv = SpreadsheetApp.newDataValidation().requireValueInList(['All Stores'].concat(CONFIG.STORES), true).setAllowInvalid(false).build();
  var currentStore = 'All Stores';
  try { currentStore = dash.getRange(FILTER_CELL).getValue() || 'All Stores'; } catch (err) {}
  dash.getRange('A5:B5').merge().setDataValidation(storeDv).setValue(currentStore)
    .setBackground('#ffffff').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a').setVerticalAlignment('middle');
  dash.setRowHeight(5, 26);
  dash.getRange('A5:B5').setBorder(true, true, true, true, false, false, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  dash.getRange('A7:B7').merge().setValue('MONTH').setFontWeight('bold').setFontSize(9).setFontColor('#64748b');
  var months = getAvailableMonths_();
  var monthDv = SpreadsheetApp.newDataValidation().requireValueInList(months, true).setAllowInvalid(false).build();
  var currentMonth = monthKey_(now);
  try { var em = dash.getRange(MONTH_FILTER_CELL).getValue(); if (em && months.indexOf(em) > -1) currentMonth = em; } catch (err) {}
  dash.getRange('A8:B8').merge().setDataValidation(monthDv).setValue(currentMonth)
    .setBackground('#ffffff').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a').setVerticalAlignment('middle');
  dash.setRowHeight(8, 26);
  dash.getRange('A8:B8').setBorder(true, true, true, true, false, false, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  dash.getRange('A10:B10').merge().setValue('VM SNAPSHOT').setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  dash.getRange('A11:B11').merge().setValue('Today Overview').setBackground('#1e293b').setFontColor('#cbd5e1').setFontSize(8);
  var visitedToday = uniqueCount_(todayRows.map(function (r) { return r[2]; }));
  var snapshotRows = [
    ['Visits Today', visitedToday + ' / ' + storesInScope.length],
    ['Avg Adherence Today', avgOf_(todayRows, 9) + '%'],
    ['Open Issues', String(openIssuesAll.length)],
    ['Visited This Month', uniqueCount_(selectedMonthRows.map(function (r) { return r[2]; })) + ' / ' + storesInScope.length]
  ];
  snapshotRows.forEach(function (row, i) {
    dash.getRange(12 + i, 1).setValue(row[0]).setFontColor('#64748b').setFontSize(9);
    dash.getRange(12 + i, 2).setValue(row[1]).setFontWeight('bold').setFontColor('#0f172a').setHorizontalAlignment('right').setFontSize(9);
  });
  panelBorder_(dash.getRange('A10:B15'));

  var sidebarR = 17;
  dash.getRange(sidebarR, 1, 1, 2).merge().setValue('DAILY ADHERENCE TREND').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
  sidebarR += 1;
  tableHeader_(dash, sidebarR, 1, ['Date', 'Avg %']);
  sidebarR += 1;
  var byDay = {};
  selectedMonthRows.forEach(function (row) {
    var d = dayKey_(row[1]);
    if (!byDay[d]) byDay[d] = [];
    if (row[9] !== '') byDay[d].push(Number(row[9]));
  });
  var trendOut = Object.keys(byDay).sort().map(function (d) {
    var vals = byDay[d];
    var avg = vals.length ? Math.round(vals.reduce(function (a, b) { return a + b; }, 0) / vals.length) : 0;
    return [Utilities.formatDate(new Date(d), TZ, 'dd MMM'), avg];
  });
  var sidebarTrend = trendOut.slice(-10);
  var trendTableTop = sidebarR;
  if (sidebarTrend.length) {
    dash.getRange(sidebarR, 1, sidebarTrend.length, 2).setValues(sidebarTrend);
    panelBorder_(dash.getRange(trendTableTop - 1, 1, sidebarTrend.length + 1, 2));
    sidebarR += sidebarTrend.length;
  } else {
    dash.getRange(sidebarR, 1, 1, 2).merge().setValue('No data for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
    panelBorder_(dash.getRange(trendTableTop - 1, 1, 2, 2));
    sidebarR += 1;
  }
  sidebarR += 2;

  dash.getRange(sidebarR, 1, 1, 2).merge().setValue('ℹ️  HOW TO USE').setBackground('#eff6ff').setFontColor('#1e40af').setFontWeight('bold').setFontSize(10);
  sidebarR += 1;
  var tips = ['Use filters to view specific store or month performance.',
    'Pick a store above to see its full deep-dive at the bottom.',
    'Data refreshes automatically when filters change.'];
  var tipsTop = sidebarR;
  tips.forEach(function (tip) {
    dash.getRange(sidebarR, 1, 1, 2).merge().setValue('•  ' + tip).setBackground('#eff6ff').setFontColor('#1e3a8a').setFontSize(8).setWrap(true);
    dash.setRowHeight(sidebarR, 30);
    sidebarR += 1;
  });
  panelBorder_(dash.getRange(tipsTop - 1, 1, tips.length + 1, 2));

  dash.setColumnWidth(1, 150); dash.setColumnWidth(2, 90); dash.setColumnWidth(3, 20);
  for (var mc = MAIN_COL; mc <= MAIN_COL + 19; mc++) dash.setColumnWidth(mc, 95);

  // --- KPI cards (2 rows of 4) ---
  var adherenceMonth = avgOf_(selectedMonthRows, 9);
  var adherenceLastMonth = lastMonthRows.length ? avgOf_(lastMonthRows, 9) : null;
  var delta = adherenceLastMonth !== null ? adherenceMonth - adherenceLastMonth : null;
  var deltaTxt = delta === null ? '' : (delta >= 0 ? ' ▲+' + delta : ' ▼' + delta) + '%';
  var highAdherence = maxOf_(selectedMonthRows, 9);
  var lowAdherence = minOf_(selectedMonthRows, 9);
  var storesVisitedMonth = uniqueCount_(selectedMonthRows.map(function (r) { return r[2]; }));
  var perfectStores = storesInScope.filter(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    return vs.length && avgOf_(vs, 9) === 100;
  }).length;
  var criticalStores = storesInScope.filter(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    return vs.length && avgOf_(vs, 9) <= 60;
  }).length;
  var monthBand = vmBand_(adherenceMonth);

  var cards = [
    { icon: '📍', label: 'VISITS TODAY', value: String(todayRows.length), sub: 'Store visits', color: '#2563eb', tint: '#eff6ff' },
    { icon: '📊', label: 'AVG ADHERENCE (MTD)' + deltaTxt, value: adherenceMonth + '%', sub: monthBand.label, color: monthBand.color, tint: '#f8fafc' },
    { icon: '🏆', label: 'HIGHEST ADHERENCE', value: highAdherence + '%', sub: 'Best visit this month', color: '#16a34a', tint: '#f0fdf4' },
    { icon: '🚨', label: 'LOWEST ADHERENCE', value: lowAdherence + '%', sub: 'Worst visit this month', color: '#dc2626', tint: '#fef2f2' },
    { icon: '🏬', label: 'STORES COVERED', value: storesVisitedMonth + ' / ' + storesInScope.length, sub: 'This month', color: '#7c3aed', tint: '#faf5ff' },
    { icon: '⚠️', label: 'OPEN ISSUES', value: String(openIssuesAll.length), sub: 'Unresolved', color: '#dc2626', tint: '#fef2f2' },
    { icon: '✅', label: 'PERFECT STORES', value: perfectStores + ' / ' + storesInScope.length, sub: '100% this month', color: '#16a34a', tint: '#f0fdf4' },
    { icon: '🔻', label: 'CRITICAL STORES', value: String(criticalStores), sub: '≤ 60% this month', color: '#dc2626', tint: '#fef2f2' }
  ];
  var cardW = 2;
  for (var ci = 0; ci < cards.length; ci++) {
    var row = ci < 4 ? 4 : 8;
    var idx = ci % 4;
    var card = cards[ci];
    var c0 = MAIN_COL + idx * (cardW + 1);
    dash.getRange(row, c0, 1, cardW).merge().setBackground(card.tint)
      .setValue(card.icon + '  ' + card.label).setFontSize(9).setFontWeight('bold').setFontColor(card.color).setVerticalAlignment('middle');
    dash.getRange(row + 1, c0, 1, cardW).merge().setBackground(card.tint)
      .setValue(card.value).setFontSize(20).setFontWeight('bold').setFontColor('#0f172a').setVerticalAlignment('middle');
    dash.getRange(row + 2, c0, 1, cardW).merge().setBackground(card.tint)
      .setValue(card.sub).setFontSize(8).setFontColor('#475569').setVerticalAlignment('middle');
    var block = dash.getRange(row, c0, 3, cardW);
    block.setBorder(true, false, true, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(row, c0, 3, 1).setBorder(false, true, false, false, false, false, card.color, SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
  dash.setRowHeight(4, 22); dash.setRowHeight(5, 30); dash.setRowHeight(6, 20);
  dash.setRowHeight(8, 22); dash.setRowHeight(9, 30); dash.setRowHeight(10, 20);

  var r = 13;

  // --- Charts: Adherence Trend | Adherence by Store | Top 5 Low-Score Items ---
  var chartColSpan = 5;
  var chartsRightCol = MAIN_COL + chartColSpan + PAIR_GAP;
  var top5Col = chartsRightCol + chartColSpan + PAIR_GAP;
  dash.getRange(r, MAIN_COL).setValue('ADHERENCE % TREND — ' + monthVal).setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, chartsRightCol).setValue('ADHERENCE % BY STORE (' + monthVal + ')').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, top5Col).setValue('TOP 5 LOW-SCORE ITEMS — ' + monthVal).setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  var chartsRow = r + 1;

  var neededCols = STAGING_COL + 8;
  if (dash.getMaxColumns() < neededCols) dash.insertColumnsAfter(dash.getMaxColumns(), neededCols - dash.getMaxColumns());

  var chartWidthPx = chartColSpan * 95;
  var trendStaging = [['Day', 'Avg Adherence %']].concat(trendOut);
  if (trendOut.length) {
    dash.getRange(chartsRow, STAGING_COL, trendStaging.length, 2).setValues(trendStaging);
    var trendChart = dash.newChart().asLineChart()
      .addRange(dash.getRange(chartsRow, STAGING_COL, trendStaging.length, 2))
      .setPosition(chartsRow, MAIN_COL, 0, 0)
      .setOption('title', null).setOption('legend', { position: 'none' })
      .setOption('width', chartWidthPx).setOption('height', 250).setOption('colors', ['#2563eb'])
      .build();
    dash.insertChart(trendChart);
  } else {
    dash.getRange(chartsRow, MAIN_COL, 8, chartColSpan).merge()
      .setValue('No adherence data yet for ' + monthVal + '.').setBackground('#f8fafc').setFontColor('#94a3b8').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    panelBorder_(dash.getRange(chartsRow, MAIN_COL, 8, chartColSpan));
  }

  var storeScoreOut = storesInScope.map(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    return [store, vs.length ? avgOf_(vs, 9) : 0];
  });
  var storeStaging = [['Store', 'Adherence %']].concat(storeScoreOut);
  dash.getRange(chartsRow, STAGING_COL + 3, storeStaging.length, 2).setValues(storeStaging);
  var storeChart = dash.newChart().asColumnChart()
    .addRange(dash.getRange(chartsRow, STAGING_COL + 3, storeStaging.length, 2))
    .setPosition(chartsRow, chartsRightCol, 0, 0)
    .setOption('title', null).setOption('legend', { position: 'none' })
    .setOption('width', chartWidthPx).setOption('height', 250).setOption('colors', ['#7c3aed'])
    .build();
  dash.insertChart(storeChart);

  tableHeader_(dash, chartsRow - 1, top5Col, ['Rank', 'Section', 'Item', 'Score'], 'danger');
  var lowScored = monthIssues.slice().sort(function (a, b) { return a[6] - b[6]; }).slice(0, 5)
    .map(function (x, idx) { return [idx + 1, x[4], x[5], x[6]]; });
  if (lowScored.length) dash.getRange(chartsRow, top5Col, lowScored.length, 4).setValues(lowScored);
  else dash.getRange(chartsRow, top5Col, 1, 4).merge().setValue('No low-scored items for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
  panelBorder_(dash.getRange(chartsRow - 1, top5Col, Math.max(lowScored.length, 1) + 1, 4));

  r = chartsRow + 12;

  // --- Area-wise adherence | Store ranking ---
  var rankCol = MAIN_COL + 3 + PAIR_GAP;
  dash.getRange(r, MAIN_COL).setValue('AREA-WISE ADHERENCE — ' + monthVal +
    (filterVal !== 'All Stores' ? ' (' + filterVal + ')' : ' (ALL STORES AVG)')).setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, rankCol).setValue('STORE RANKING (' + monthVal + ')').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  r += 1;
  tableHeader_(dash, r, MAIN_COL, ['Section', 'Sum Score', 'Max Possible', 'Adherence %', 'Rating']);
  tableHeader_(dash, r, rankCol, ['Rank', 'Store', 'Adherence %']);
  var areaTop = r + 1;

  var areaOut = SECTIONS.map(function (sec) {
    var rows = monthScores.filter(function (x) { return x[4] === sec.name; });
    var sum = rows.reduce(function (s, x) { return s + Number(x[6]); }, 0);
    var max = rows.length * 5;
    var pct = max > 0 ? Math.round(sum / max * 100) : 0;
    return [sec.name, sum, max, pct, vmBand_(pct).label];
  });
  areaOut.forEach(function (row, i) {
    dash.getRange(areaTop + i, MAIN_COL, 1, 5).setValues([row]);
    statusPill_(dash.getRange(areaTop + i, MAIN_COL + 4), row[4], vmBand_(row[3]).color);
  });
  panelBorder_(dash.getRange(areaTop - 1, MAIN_COL, areaOut.length + 1, 5));

  var ranked = storeScoreOut.slice().sort(function (a, b) { return b[1] - a[1]; })
    .map(function (row, idx) { return [idx + 1, row[0], row[1] + '%']; });
  if (ranked.length) dash.getRange(areaTop, rankCol, ranked.length, 3).setValues(ranked);
  else dash.getRange(areaTop, rankCol, 1, 3).merge().setValue('No visits yet.').setFontSize(9).setFontColor('#94a3b8');
  panelBorder_(dash.getRange(areaTop - 1, rankCol, Math.max(ranked.length, 1) + 1, 3));

  r = areaTop + Math.max(areaOut.length, ranked.length, 1) + 2;

  // --- Heat map | Open issues ---
  var heatHeaders = ['Store'].concat(SECTIONS.map(function (s) { return s.name; })).concat(['Total Issues']);
  var heatRightCol = MAIN_COL + heatHeaders.length + PAIR_GAP;
  dash.getRange(r, MAIN_COL).setValue('HEAT MAP — LOW SCORES BY STORE & SECTION (' + monthVal + ')').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, heatRightCol).setValue('OPEN ISSUES').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  r += 1;
  tableHeader_(dash, r, MAIN_COL, heatHeaders);
  tableHeader_(dash, r, heatRightCol, ['Store', 'Section', 'Item', 'Score'], 'danger');
  var heatTop = r + 1;

  var heatOut = storesInScope.map(function (store) {
    var counts = SECTIONS.map(function (sec) {
      return monthIssues.filter(function (x) { return x[2] === store && x[4] === sec.name; }).length;
    });
    var total = counts.reduce(function (a, b) { return a + b; }, 0);
    return [store].concat(counts).concat([total]);
  });
  dash.getRange(heatTop, MAIN_COL, heatOut.length, heatHeaders.length).setValues(heatOut);
  var heatRange = dash.getRange(heatTop, MAIN_COL + 1, heatOut.length, heatHeaders.length - 1);
  var maxIssue = Math.max.apply(null, heatOut.map(function (row) { return Math.max.apply(null, row.slice(1)); }).concat([1]));
  var heatRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue('#dc2626', SpreadsheetApp.InterpolationType.NUMBER, String(maxIssue))
    .setRanges([heatRange]).build();
  dash.setConditionalFormatRules(dash.getConditionalFormatRules().concat([heatRule]));
  dash.getRange(heatTop, MAIN_COL, heatOut.length, heatHeaders.length).setHorizontalAlignment('center');
  panelBorder_(dash.getRange(heatTop - 1, MAIN_COL, heatOut.length + 1, heatHeaders.length));

  if (openIssuesAll.length) {
    var openList = openIssuesAll.slice(0, 30);
    openList.forEach(function (x, i) {
      dash.getRange(heatTop + i, heatRightCol, 1, 4).setValues([[x[2], x[4], x[5], x[6]]]);
    });
  } else {
    dash.getRange(heatTop, heatRightCol, 1, 4).merge().setValue('No open issues 🎉').setFontSize(9).setFontColor('#94a3b8');
  }
  panelBorder_(dash.getRange(heatTop - 1, heatRightCol, Math.max(Math.min(openIssuesAll.length, 30), 1) + 1, 4));

  r = heatTop + Math.max(heatOut.length, Math.min(openIssuesAll.length, 30), 1) + 2;

  // --- Drill-through: Store Deep-Dive ---
  if (filterVal !== 'All Stores') {
    dash.getRange(r, MAIN_COL, 1, 6).merge().setValue('🔍  STORE DEEP-DIVE — ' + filterVal)
      .setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12).setVerticalAlignment('middle');
    dash.setRowHeight(r, 26);
    r += 2;

    var deepRightCol = MAIN_COL + 4 + PAIR_GAP;
    dash.getRange(r, MAIN_COL).setValue('Visit History').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
    dash.getRange(r, deepRightCol).setValue('Open Issues').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
    r += 1;
    tableHeader_(dash, r, MAIN_COL, ['Date', 'VM', 'Adherence %', 'Rating']);
    tableHeader_(dash, r, deepRightCol, ['Date', 'Section', 'Item', 'Score'], 'danger');
    var deepTop = r + 1;

    var visitHistory = selectedMonthRows.slice().sort(function (a, b) { return b[1] - a[1]; })
      .map(function (x) { return [Utilities.formatDate(new Date(x[1]), TZ, 'dd MMM'), x[4], x[9] + '%', vmBand_(Number(x[9]) || 0).label]; });
    if (visitHistory.length) dash.getRange(deepTop, MAIN_COL, visitHistory.length, 4).setValues(visitHistory);
    else dash.getRange(deepTop, MAIN_COL, 1, 4).merge().setValue('No visits for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
    panelBorder_(dash.getRange(deepTop - 1, MAIN_COL, Math.max(visitHistory.length, 1) + 1, 4));

    var storeIssues = monthIssues.filter(function (x) { return x[7] !== 'Closed'; })
      .map(function (x) { return [Utilities.formatDate(new Date(x[1]), TZ, 'dd MMM'), x[4], x[5], x[6]]; });
    if (storeIssues.length) dash.getRange(deepTop, deepRightCol, storeIssues.length, 4).setValues(storeIssues);
    else dash.getRange(deepTop, deepRightCol, 1, 4).merge().setValue('No open issues. 🎉').setFontSize(9).setFontColor('#94a3b8');
    panelBorder_(dash.getRange(deepTop - 1, deepRightCol, Math.max(storeIssues.length, 1) + 1, 4));

    r = deepTop + Math.max(visitHistory.length, storeIssues.length, 1) + 2;
  }

  // --- Export to PDF ---
  r += 1;
  dash.getRange(r, MAIN_COL, 2, 4).merge().setValue('⬇  EXPORT TO PDF')
    .setBackground('#dc2626').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  var exportCheckboxCol = MAIN_COL + 4;
  dash.getRange(r, exportCheckboxCol, 2, 1).merge().insertCheckboxes().setValue(false)
    .setBackground('#fee2e2').setVerticalAlignment('middle').setHorizontalAlignment('center');
  dash.setRowHeight(r, 26); dash.setRowHeight(r + 1, 26);
  PropertiesService.getDocumentProperties().setProperty('EXPORT_BTN_ROW', String(r));
  PropertiesService.getDocumentProperties().setProperty('EXPORT_BTN_COL', String(exportCheckboxCol));

  // --- Final polish ---
  dash.getRange(4, 3, Math.max(r - 3, 1), 1).setBorder(false, false, false, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
  dash.hideColumns(STAGING_COL, 8);
  dash.setFrozenRows(2);
  ensureLogoOnDashboard_(dash);
}

function avgOf_(rows, col) {
  var vals = rows.filter(function (r) { return r[col] !== ''; }).map(function (r) { return Number(r[col]); });
  return vals.length ? Math.round(vals.reduce(function (a, b) { return a + b; }, 0) / vals.length) : 0;
}
function maxOf_(rows, col) {
  var vals = rows.filter(function (r) { return r[col] !== ''; }).map(function (r) { return Number(r[col]); });
  return vals.length ? Math.max.apply(null, vals) : 0;
}
function minOf_(rows, col) {
  var vals = rows.filter(function (r) { return r[col] !== ''; }).map(function (r) { return Number(r[col]); });
  return vals.length ? Math.min.apply(null, vals) : 0;
}
