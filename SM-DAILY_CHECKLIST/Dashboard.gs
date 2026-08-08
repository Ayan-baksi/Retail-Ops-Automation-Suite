// Dashboard.gs
// SM Store Manager live dashboard, following the same patterns as the
// AM system (content-based pairing, checkbox-driven controls,
// self-healing checkbox sweep, dynamic Export-to-PDF position), adapted
// for Opening/Closing KPIs, tables, and a Store Deep-Dive that includes photos.

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

function statusPill_(range, text, kind) {
  var bg = kind === 'good' ? '#dcfce7' : kind === 'fair' ? '#fef3c7' : kind === 'poor' ? '#fee2e2' : '#f1f5f9';
  var fg = kind === 'good' ? '#166534' : kind === 'fair' ? '#92400e' : kind === 'poor' ? '#991b1b' : '#475569';
  range.setValue(text).setBackground(bg).setFontColor(fg).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function tableHeader_(dash, row, col, headers, tint) {
  var bg = tint === 'danger' ? '#fef2f2' : '#f1f5f9';
  var fg = tint === 'danger' ? '#991b1b' : '#334155';
  dash.getRange(row, col, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground(bg).setFontColor(fg);
}

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
    if (e.range.getValue() === true) {
      refreshDashboard();
      sh.getRange(REFRESH_CHECKBOX_CELL).setValue(false);
    }
    return;
  }
  if (a1 === FILTER_CELL || a1 === MONTH_FILTER_CELL || a1 === TYPE_FILTER_CELL || a1 === DAY_FILTER_CELL) {
    refreshDashboard();
    return;
  }

  var props = PropertiesService.getDocumentProperties();
  var exRow = Number(props.getProperty('EXPORT_BTN_ROW'));
  var exCol = Number(props.getProperty('EXPORT_BTN_COL'));
  if (exRow && e.range.getRow() === exRow && e.range.getColumn() === exCol) {
    if (e.range.getValue() === true) {
      sh.getRange(exRow, exCol).setValue(false);
      exportDashboardToPdf();
    }
  }
}

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

// Every calendar day in the given month (yyyy-MM), newest first, with
// "Today" pinned at the top as a shortcut. Includes days with no
// submissions too, since those may be exactly what needs checking.
function getAvailableDays_(monthVal) {
  var year = Number(monthVal.slice(0, 4));
  var month = Number(monthVal.slice(5, 7)); // 1-12
  var daysInMonth = new Date(year, month, 0).getDate();
  var days = [];
  for (var d = daysInMonth; d >= 1; d--) {
    days.push(Utilities.formatDate(new Date(year, month - 1, d), TZ, 'yyyy-MM-dd'));
  }
  return ['Today'].concat(days);
}

function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(SHEETS.DASH) || ss.insertSheet(SHEETS.DASH);
  var master = ss.getSheetByName(SHEETS.MASTER);
  var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
  var photoSheet = ss.getSheetByName(SHEETS.PHOTOS);

  var filterVal = 'All Stores', monthVal = monthKey_(new Date()), typeVal = 'Both', dayVal = 'Today';
  try { filterVal = dash.getRange(FILTER_CELL).getValue() || 'All Stores'; } catch (err) {}
  try {
    var mv = dash.getRange(MONTH_FILTER_CELL).getValue();
    monthVal = (mv instanceof Date) ? monthKey_(mv) : (mv || monthVal);
  } catch (err) {}
  try { typeVal = dash.getRange(TYPE_FILTER_CELL).getValue() || 'Both'; } catch (err) {}
  try {
    var dv = dash.getRange(DAY_FILTER_CELL).getValue();
    dayVal = (dv instanceof Date) ? dayKey_(dv) : (dv || 'Today');
  } catch (err) {}

  dash.getCharts().forEach(function (c) { dash.removeChart(c); });
  dash.setConditionalFormatRules([]);
  dash.clear();
  dash.setHiddenGridlines(true);
  clearAllCheckboxValidations_(dash);

  var now = new Date();
  var today = dayVal === 'Today' ? dayKey_(now) : dayVal;
  var dayLabel = dayVal === 'Today' ? 'TODAY' : Utilities.formatDate(new Date(today), TZ, 'dd MMM yyyy').toUpperCase();
  var overdueDays = Number(getSetting_('Overdue Days Threshold', CONFIG.OVERDUE_DAYS_DEFAULT));
  var targetScore = Number(getSetting_('Target Score %', CONFIG.TARGET_SCORE_DEFAULT));
  var criticalThreshold = 50;

  var mRowsAll = (master && master.getLastRow() > 1)
    ? master.getRange(2, 1, master.getLastRow() - 1, 12).getValues() : [];
  // Columns: 0 ts, 1 date, 2 store, 3 sm, 4 type, 5 total, 6 yes, 7 no, 8 na, 9 scorePct, 10 sectionRemarks, 11 overallRemarks
  var iRowsAll = (issueSheet && issueSheet.getLastRow() > 1)
    ? issueSheet.getRange(2, 1, issueSheet.getLastRow() - 1, 12).getValues() : [];
  // Columns: 0 ts, 1 date, 2 store, 3 sm, 4 type, 5 section, 6 item, 7 status, 8 assignedTo, 9 dueDate, 10 closedDate, 11 severity
  var pRowsAll = (photoSheet && photoSheet.getLastRow() > 1)
    ? photoSheet.getRange(2, 1, photoSheet.getLastRow() - 1, 8).getValues() : [];
  // Columns: 0 ts, 1 date, 2 store, 3 sm, 4 type, 5 section, 6 item, 7 url

  var byStore = function (rows) { return filterVal !== 'All Stores' ? rows.filter(function (r) { return r[2] === filterVal; }) : rows; };
  var byType = function (rows, typeCol) { return typeVal !== 'Both' ? rows.filter(function (r) { return r[typeCol] === typeVal; }) : rows; };

  var mRows = byType(byStore(mRowsAll), 4);
  var iRows = byType(byStore(iRowsAll), 4);
  var pRows = byType(byStore(pRowsAll), 4);
  var storesInScope = filterVal === 'All Stores' ? CONFIG.STORES : [filterVal];

  var selectedMonthRows = mRows.filter(function (r) { return monthKey_(r[1]) === monthVal; });
  var monthIssues = iRows.filter(function (r) { return monthKey_(r[1]) === monthVal; });
  var todayRows = mRows.filter(function (r) { return dayKey_(r[1]) === today; });
  var openIssuesAll = iRows.filter(function (r) { return r[7] !== 'Closed'; });

  var openingTodayRows = todayRows.filter(function (r) { return r[4] === 'Opening'; });
  var closingTodayRows = todayRows.filter(function (r) { return r[4] === 'Closing'; });
  var openingMonthRows = selectedMonthRows.filter(function (r) { return r[4] === 'Opening'; });
  var closingMonthRows = selectedMonthRows.filter(function (r) { return r[4] === 'Closing'; });

  // --- Header ---
  var headerLastCol = MAIN_COL + 5 * 3 - 1;
  for (var c = 1; c <= headerLastCol; c++) { dash.getRange(1, c).setBackground('#0f172a'); dash.getRange(2, c).setBackground('#0f172a'); }
  dash.getRange(1, 3, 2, 2).merge();
  dash.getRange(1, 5, 1, 6).merge().setValue('STORE MANAGER — LIVE AUDIT DASHBOARD')
    .setFontColor('#ffffff').setFontWeight('bold').setFontSize(16).setVerticalAlignment('middle');
  dash.getRange(2, 5, 1, 6).merge().setValue('Real-time overview of daily opening/closing checklists')
    .setFontColor('#94a3b8').setFontSize(9).setVerticalAlignment('middle');
  dash.getRange(1, 12, 1, headerLastCol - 11).merge()
    .setValue('Last Updated: ' + Utilities.formatDate(now, TZ, 'dd MMM yyyy, hh:mm a'))
    .setFontColor('#94a3b8').setFontSize(9).setHorizontalAlignment('right').setVerticalAlignment('middle');
  dash.getRange(2, 12, 1, headerLastCol - 15).merge()
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
  var currentMonth = months.indexOf(monthVal) > -1 ? monthVal : monthKey_(now);
  dash.getRange('A8:B8').setNumberFormat('@');
  dash.getRange('A8:B8').merge().setDataValidation(monthDv).setValue(currentMonth)
    .setBackground('#ffffff').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a').setVerticalAlignment('middle');
  dash.setRowHeight(8, 26);
  dash.getRange('A8:B8').setBorder(true, true, true, true, false, false, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  dash.getRange('A10:B10').merge().setValue('AUDIT TYPE').setFontWeight('bold').setFontSize(9).setFontColor('#64748b');
  var typeDv = SpreadsheetApp.newDataValidation().requireValueInList(['Both', 'Opening', 'Closing'], true).setAllowInvalid(false).build();
  var currentType = 'Both';
  try { currentType = dash.getRange(TYPE_FILTER_CELL).getValue() || 'Both'; } catch (err) {}
  dash.getRange('A11:B11').merge().setDataValidation(typeDv).setValue(currentType)
    .setBackground('#ffffff').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a').setVerticalAlignment('middle');
  dash.setRowHeight(11, 26);
  dash.getRange('A11:B11').setBorder(true, true, true, true, false, false, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  dash.getRange('A13:B13').merge().setValue('DAY').setFontWeight('bold').setFontSize(9).setFontColor('#64748b');
  var days = getAvailableDays_(monthVal);
  var dayDv = SpreadsheetApp.newDataValidation().requireValueInList(days, true).setAllowInvalid(false).build();
  var currentDay = days.indexOf(dayVal) > -1 ? dayVal : 'Today';
  dash.getRange('A14:B14').setNumberFormat('@');
  dash.getRange('A14:B14').merge().setDataValidation(dayDv).setValue(currentDay)
    .setBackground('#ffffff').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a').setVerticalAlignment('middle');
  dash.setRowHeight(14, 26);
  dash.getRange('A14:B14').setBorder(true, true, true, true, false, false, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  dash.getRange('A16:B16').merge().setValue('SM SNAPSHOT').setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  dash.getRange('A17:B17').merge().setValue(dayLabel + ' OVERVIEW').setBackground('#1e293b').setFontColor('#cbd5e1').setFontSize(8);
  var openingDoneToday = uniqueCount_(openingTodayRows.map(function (r) { return r[2]; }));
  var closingDoneToday = uniqueCount_(closingTodayRows.map(function (r) { return r[2]; }));
  var snapshotRows = [
    ['Opening Done', openingDoneToday + ' / ' + storesInScope.length],
    ['Closing Done', closingDoneToday + ' / ' + storesInScope.length],
    ['Open Issues', String(openIssuesAll.length)],
    ['Photos ' + dayLabel, String(pRows.filter(function (p) { return dayKey_(p[1]) === today; }).length)]
  ];
  snapshotRows.forEach(function (row, i) {
    dash.getRange(18 + i, 1).setValue(row[0]).setFontColor('#64748b').setFontSize(9);
    dash.getRange(18 + i, 2).setValue(row[1]).setFontWeight('bold').setFontColor('#0f172a').setHorizontalAlignment('right').setFontSize(9);
  });
  panelBorder_(dash.getRange('A16:B21'));

  var sidebarR = 23;
  dash.getRange(sidebarR, 1, 1, 2).merge().setValue('DAILY SCORE TREND').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
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
  var tips = ['Use filters to view specific store, month, or audit type.',
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

  // --- KPI cards (2 rows of 5) ---
  var avgOpeningToday = avgOf_(openingTodayRows, 9);
  var avgClosingToday = avgOf_(closingTodayRows, 9);
  var avgOpeningMonth = avgOf_(openingMonthRows, 9);
  var avgClosingMonth = avgOf_(closingMonthRows, 9);
  var monthlyAvg = avgOf_(selectedMonthRows, 9);
  var storesPendingOpening = storesInScope.length - openingDoneToday;
  var storesPendingClosing = storesInScope.length - closingDoneToday;
  var perfectStores = storesInScope.filter(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    return vs.length && avgOf_(vs, 9) === 100;
  }).length;
  var criticalStores = storesInScope.filter(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    return vs.length && avgOf_(vs, 9) < criticalThreshold;
  }).length;

  var cards = [
    { icon: '🌅', label: dayLabel + "'S OPENING", value: openingDoneToday + ' / ' + storesInScope.length, sub: 'Completed', color: '#2563eb', tint: '#eff6ff' },
    { icon: '🌆', label: dayLabel + "'S CLOSING", value: closingDoneToday + ' / ' + storesInScope.length, sub: 'Completed', color: '#7c3aed', tint: '#faf5ff' },
    { icon: '✅', label: 'AVG OPENING SCORE', value: avgOpeningMonth + '%', sub: 'This month', color: '#16a34a', tint: '#f0fdf4' },
    { icon: '✅', label: 'AVG CLOSING SCORE', value: avgClosingMonth + '%', sub: 'This month', color: '#16a34a', tint: '#f0fdf4' },
    { icon: '📅', label: 'MONTHLY AVG SCORE', value: monthlyAvg + '%', sub: monthVal, color: '#0891b2', tint: '#ecfeff' },
    { icon: '⚠️', label: 'OPEN ISSUES', value: String(openIssuesAll.length), sub: 'Unresolved', color: '#dc2626', tint: '#fef2f2' },
    { icon: '🕒', label: 'PENDING OPENING', value: String(storesPendingOpening), sub: 'Stores today', color: '#b45309', tint: '#fffbeb' },
    { icon: '🕒', label: 'PENDING CLOSING', value: String(storesPendingClosing), sub: 'Stores today', color: '#b45309', tint: '#fffbeb' },
    { icon: '🏆', label: 'PERFECT STORES', value: perfectStores + ' / ' + storesInScope.length, sub: '100% this month', color: '#7c3aed', tint: '#faf5ff' },
    { icon: '🚨', label: 'CRITICAL STORES', value: String(criticalStores), sub: '< ' + criticalThreshold + '% this month', color: '#dc2626', tint: '#fef2f2' }
  ];
  var cardW = 2;
  for (var ci = 0; ci < cards.length; ci++) {
    var row = ci < 5 ? 4 : 8;
    var idx = ci % 5;
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

  // --- Charts: Opening vs Closing Trend | Avg Score by Store | Top 5 Issues ---
  var chartColSpan = 5;
  var chartsRightCol = MAIN_COL + chartColSpan + PAIR_GAP;
  var top5Col = chartsRightCol + chartColSpan + PAIR_GAP;
  dash.getRange(r, MAIN_COL).setValue('OPENING vs CLOSING SCORE TREND — ' + monthVal).setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, chartsRightCol).setValue('AVG SCORE BY STORE (' + monthVal + ')').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, top5Col).setValue('TOP 5 RECURRING ISSUES — ' + monthVal).setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  var chartsRow = r + 1;

  var neededCols = STAGING_COL + 8;
  if (dash.getMaxColumns() < neededCols) dash.insertColumnsAfter(dash.getMaxColumns(), neededCols - dash.getMaxColumns());

  var byDayOpen = {}, byDayClose = {};
  selectedMonthRows.forEach(function (row) {
    var d = dayKey_(row[1]);
    if (row[9] === '') return;
    if (row[4] === 'Opening') { if (!byDayOpen[d]) byDayOpen[d] = []; byDayOpen[d].push(Number(row[9])); }
    if (row[4] === 'Closing') { if (!byDayClose[d]) byDayClose[d] = []; byDayClose[d].push(Number(row[9])); }
  });
  var allDays = Object.keys(byDay).sort();
  var avgArr = function (arr) { return arr && arr.length ? Math.round(arr.reduce(function (a, b) { return a + b; }, 0) / arr.length) : null; };
  var comboStaging = [['Day', 'Opening %', 'Closing %']].concat(allDays.map(function (d) {
    return [Utilities.formatDate(new Date(d), TZ, 'dd MMM'), avgArr(byDayOpen[d]), avgArr(byDayClose[d])];
  }));
  var chartWidthPx = chartColSpan * 95;
  if (allDays.length) {
    dash.getRange(chartsRow, STAGING_COL, comboStaging.length, 3).setValues(comboStaging);
    var trendChart = dash.newChart().asLineChart()
      .addRange(dash.getRange(chartsRow, STAGING_COL, comboStaging.length, 3))
      .setPosition(chartsRow, MAIN_COL, 0, 0)
      .setOption('title', null).setOption('legend', { position: 'top' })
      .setOption('width', chartWidthPx).setOption('height', 250)
      .setOption('colors', ['#2563eb', '#7c3aed'])
      .build();
    dash.insertChart(trendChart);
  } else {
    dash.getRange(chartsRow, MAIN_COL, 8, chartColSpan).merge()
      .setValue('No score data yet for ' + monthVal + '.').setBackground('#f8fafc').setFontColor('#94a3b8').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    panelBorder_(dash.getRange(chartsRow, MAIN_COL, 8, chartColSpan));
  }

  var storeScoreOut = storesInScope.map(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    return [store, vs.length ? avgOf_(vs, 9) : 0];
  });
  var storeStaging = [['Store', 'Avg Score %']].concat(storeScoreOut);
  dash.getRange(chartsRow, STAGING_COL + 4, storeStaging.length, 2).setValues(storeStaging);
  var storeChart = dash.newChart().asColumnChart()
    .addRange(dash.getRange(chartsRow, STAGING_COL + 4, storeStaging.length, 2))
    .setPosition(chartsRow, chartsRightCol, 0, 0)
    .setOption('title', null).setOption('legend', { position: 'none' })
    .setOption('width', chartWidthPx).setOption('height', 250).setOption('colors', ['#dc2626'])
    .build();
  dash.insertChart(storeChart);

  tableHeader_(dash, chartsRow - 1, top5Col, ['Rank', 'Section', 'Item', 'Times Failed'], 'danger');
  var counts5 = {};
  monthIssues.forEach(function (x) {
    var key = x[5] + ' || ' + x[6];
    counts5[key] = (counts5[key] || 0) + 1;
  });
  var top5 = Object.keys(counts5).map(function (k) {
    var parts = k.split(' || ');
    return [parts[0], parts[1], counts5[k]];
  }).sort(function (a, b) { return b[2] - a[2]; }).slice(0, 5).map(function (row, idx) { return [idx + 1, row[0], row[1], row[2]]; });
  if (top5.length) dash.getRange(chartsRow, top5Col, top5.length, 4).setValues(top5);
  else dash.getRange(chartsRow, top5Col, 1, 4).merge().setValue('No failed items for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
  panelBorder_(dash.getRange(chartsRow - 1, top5Col, Math.max(top5.length, 1) + 1, 4));

  r = chartsRow + 12;

  // --- Today's Opening | Today's Closing | Store Ranking ---
  var closingCol = MAIN_COL + 4 + PAIR_GAP;
  var rankCol = closingCol + 4 + PAIR_GAP;
  dash.getRange(r, MAIN_COL).setValue(dayLabel + "'S OPENING STATUS").setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, closingCol).setValue(dayLabel + "'S CLOSING STATUS").setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, rankCol).setValue('STORE RANKING (' + monthVal + ')').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  r += 1;
  tableHeader_(dash, r, MAIN_COL, ['Store', 'SM', 'Score %', 'Status']);
  tableHeader_(dash, r, closingCol, ['Store', 'SM', 'Score %', 'Status']);
  tableHeader_(dash, r, rankCol, ['Rank', 'Store', 'Score %']);
  var statusTop = r + 1;

  var statusRow_ = function (rows, store) {
    var v = rows.filter(function (x) { return x[2] === store; }).pop();
    if (!v) return [store, '—', '—', 'NOT_DONE'];
    var kind = v[9] >= targetScore ? 'good' : v[9] >= 70 ? 'fair' : 'poor';
    return [store, v[3], v[9] + '%', kind];
  };
  var openingOut = storesInScope.map(function (s) { return statusRow_(openingTodayRows, s); });
  var closingOut = storesInScope.map(function (s) { return statusRow_(closingTodayRows, s); });

  // Batch all value cells into one write each, instead of one write per row.
  dash.getRange(statusTop, MAIN_COL, openingOut.length, 3)
    .setValues(openingOut.map(function (row) { return [row[0], row[1], row[2]]; }));
  openingOut.forEach(function (row, i) {
    var kind = row[3] === 'NOT_DONE' ? 'neutral' : row[3];
    var txt = row[3] === 'NOT_DONE' ? 'Pending' : (row[3] === 'good' ? 'Good' : row[3] === 'fair' ? 'Fair' : 'Poor');
    statusPill_(dash.getRange(statusTop + i, MAIN_COL + 3), txt, kind);
  });
  panelBorder_(dash.getRange(statusTop - 1, MAIN_COL, openingOut.length + 1, 4));

  dash.getRange(statusTop, closingCol, closingOut.length, 3)
    .setValues(closingOut.map(function (row) { return [row[0], row[1], row[2]]; }));
  closingOut.forEach(function (row, i) {
    var kind = row[3] === 'NOT_DONE' ? 'neutral' : row[3];
    var txt = row[3] === 'NOT_DONE' ? 'Pending' : (row[3] === 'good' ? 'Good' : row[3] === 'fair' ? 'Fair' : 'Poor');
    statusPill_(dash.getRange(statusTop + i, closingCol + 3), txt, kind);
  });
  panelBorder_(dash.getRange(statusTop - 1, closingCol, closingOut.length + 1, 4));

  var ranked = storeScoreOut.slice().sort(function (a, b) { return b[1] - a[1]; })
    .map(function (row, idx) { return [idx + 1, row[0], row[1] + '%']; });
  if (ranked.length) dash.getRange(statusTop, rankCol, ranked.length, 3).setValues(ranked);
  else dash.getRange(statusTop, rankCol, 1, 3).merge().setValue('No visits yet.').setFontSize(9).setFontColor('#94a3b8');
  panelBorder_(dash.getRange(statusTop - 1, rankCol, Math.max(ranked.length, 1) + 1, 3));

  r = statusTop + Math.max(openingOut.length, closingOut.length, ranked.length, 1) + 2;

  // --- Heat map | Open issues ---
  var heatHeaders = ['Store'].concat(SECTIONS_OPENING.concat(SECTIONS_CLOSING).map(function (s) { return s.name; })).concat(['Total Failures']);
  // De-duplicate section names in case a name repeats between Opening/Closing lists.
  var seenNames = {};
  heatHeaders = heatHeaders.filter(function (n, i) { if (i === 0 || i === heatHeaders.length - 1) return true; if (seenNames[n]) return false; seenNames[n] = true; return true; });
  var heatRightCol = MAIN_COL + heatHeaders.length + PAIR_GAP;
  dash.getRange(r, MAIN_COL).setValue('HEAT MAP — FAILURES BY STORE & SECTION (' + monthVal + ')').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, heatRightCol).setValue('OPEN ISSUES').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  r += 1;
  tableHeader_(dash, r, MAIN_COL, heatHeaders);
  tableHeader_(dash, r, heatRightCol, ['Store', 'Section', 'Item', 'Status'], 'danger');
  var heatTop = r + 1;
  var heatSectionNames = heatHeaders.slice(1, -1);
  var heatOut = storesInScope.map(function (store) {
    var counts = heatSectionNames.map(function (name) {
      return monthIssues.filter(function (x) { return x[2] === store && x[5] === name; }).length;
    });
    var total = counts.reduce(function (a, b) { return a + b; }, 0);
    return [store].concat(counts).concat([total]);
  });
  dash.getRange(heatTop, MAIN_COL, heatOut.length, heatHeaders.length).setValues(heatOut);
  var heatRange = dash.getRange(heatTop, MAIN_COL + 1, heatOut.length, heatHeaders.length - 1);
  var maxFailure = Math.max.apply(null, heatOut.map(function (row) { return Math.max.apply(null, row.slice(1)); }).concat([1]));
  var heatRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue('#dc2626', SpreadsheetApp.InterpolationType.NUMBER, String(maxFailure))
    .setRanges([heatRange]).build();
  dash.setConditionalFormatRules(dash.getConditionalFormatRules().concat([heatRule]));
  dash.getRange(heatTop, MAIN_COL, heatOut.length, heatHeaders.length).setHorizontalAlignment('center');
  panelBorder_(dash.getRange(heatTop - 1, MAIN_COL, heatOut.length + 1, heatHeaders.length));

  if (openIssuesAll.length) {
    var openList = openIssuesAll.slice(0, 30);
    dash.getRange(heatTop, heatRightCol, openList.length, 3)
      .setValues(openList.map(function (x) { return [x[2], x[5], x[6]]; }));
    openList.forEach(function (x, i) {
      statusPill_(dash.getRange(heatTop + i, heatRightCol + 3), x[7] || 'Open', x[7] === 'Closed' ? 'good' : 'poor');
    });
  } else {
    dash.getRange(heatTop, heatRightCol, 1, 4).merge().setValue('No open issues 🎉').setFontSize(9).setFontColor('#94a3b8');
  }
  panelBorder_(dash.getRange(heatTop - 1, heatRightCol, Math.max(Math.min(openIssuesAll.length, 30), 1) + 1, 4));

  r = heatTop + Math.max(heatOut.length, Math.min(openIssuesAll.length, 30), 1) + 2;

  // --- Monthly performance ---
  dash.getRange(r, MAIN_COL).setValue('MONTHLY PERFORMANCE — ' + monthVal).setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  r += 1;
  tableHeader_(dash, r, MAIN_COL, ['Store', 'Opening Visits', 'Closing Visits', 'Avg Opening %', 'Avg Closing %', 'Overall Avg %']);
  var perfTop = r + 1;
  var perfOut = storesInScope.map(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    var opens = vs.filter(function (x) { return x[4] === 'Opening'; });
    var closes = vs.filter(function (x) { return x[4] === 'Closing'; });
    return [store, opens.length, closes.length, opens.length ? avgOf_(opens, 9) + '%' : '—',
      closes.length ? avgOf_(closes, 9) + '%' : '—', vs.length ? avgOf_(vs, 9) + '%' : '—'];
  });
  dash.getRange(perfTop, MAIN_COL, perfOut.length, 6).setValues(perfOut);
  panelBorder_(dash.getRange(perfTop - 1, MAIN_COL, perfOut.length + 1, 6));
  r = perfTop + perfOut.length + 2;

  // --- Drill-through: Store Deep-Dive ---
  if (filterVal !== 'All Stores') {
    dash.getRange(r, MAIN_COL, 1, 6).merge().setValue('🔍  STORE DEEP-DIVE — ' + filterVal)
      .setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12).setVerticalAlignment('middle');
    dash.setRowHeight(r, 26);
    r += 2;

    var deepRightCol = MAIN_COL + 4 + PAIR_GAP;
    dash.getRange(r, MAIN_COL).setValue('Opening History').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
    dash.getRange(r, deepRightCol).setValue('Closing History').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
    r += 1;
    tableHeader_(dash, r, MAIN_COL, ['Date', 'SM', 'Score %', 'Issues']);
    tableHeader_(dash, r, deepRightCol, ['Date', 'SM', 'Score %', 'Issues']);
    var histTop = r + 1;

    var histRow_ = function (rows) {
      return rows.slice().sort(function (a, b) { return b[1] - a[1]; })
        .map(function (x) { return [Utilities.formatDate(new Date(x[1]), TZ, 'dd MMM'), x[3], x[9] + '%', x[7]]; });
    };
    var openHist = histRow_(openingMonthRows), closeHist = histRow_(closingMonthRows);
    if (openHist.length) dash.getRange(histTop, MAIN_COL, openHist.length, 4).setValues(openHist);
    else dash.getRange(histTop, MAIN_COL, 1, 4).merge().setValue('No Opening visits for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
    panelBorder_(dash.getRange(histTop - 1, MAIN_COL, Math.max(openHist.length, 1) + 1, 4));
    if (closeHist.length) dash.getRange(histTop, deepRightCol, closeHist.length, 4).setValues(closeHist);
    else dash.getRange(histTop, deepRightCol, 1, 4).merge().setValue('No Closing visits for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
    panelBorder_(dash.getRange(histTop - 1, deepRightCol, Math.max(closeHist.length, 1) + 1, 4));

    r = histTop + Math.max(openHist.length, closeHist.length, 1) + 2;

    dash.getRange(r, MAIN_COL).setValue('Open Issues').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
    dash.getRange(r, deepRightCol).setValue('Uploaded Photos').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
    r += 1;
    tableHeader_(dash, r, MAIN_COL, ['Date', 'Type', 'Section', 'Item'], 'danger');
    tableHeader_(dash, r, deepRightCol, ['Date', 'Type', 'Item', 'Link']);
    var d2Top = r + 1;

    var storeIssues = monthIssues.filter(function (x) { return x[7] !== 'Closed'; })
      .map(function (x) { return [Utilities.formatDate(new Date(x[1]), TZ, 'dd MMM'), x[4], x[5], x[6]]; });
    if (storeIssues.length) dash.getRange(d2Top, MAIN_COL, storeIssues.length, 4).setValues(storeIssues);
    else dash.getRange(d2Top, MAIN_COL, 1, 4).merge().setValue('No open issues. 🎉').setFontSize(9).setFontColor('#94a3b8');
    panelBorder_(dash.getRange(d2Top - 1, MAIN_COL, Math.max(storeIssues.length, 1) + 1, 4));

    var storePhotos = pRows.filter(function (x) { return monthKey_(x[1]) === monthVal; })
      .map(function (x) { return [Utilities.formatDate(new Date(x[1]), TZ, 'dd MMM'), x[4], x[6], x[7]]; });
    if (storePhotos.length) {
      storePhotos.forEach(function (row, i) {
        var rr = d2Top + i;
        dash.getRange(rr, deepRightCol, 1, 3).setValues([[row[0], row[1], row[2]]]);
        dash.getRange(rr, deepRightCol + 3).setFormula('=HYPERLINK("' + row[3] + '","View Photo")');
      });
    } else {
      dash.getRange(d2Top, deepRightCol, 1, 4).merge().setValue('No photos uploaded for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
    }
    panelBorder_(dash.getRange(d2Top - 1, deepRightCol, Math.max(storePhotos.length, 1) + 1, 4));

    r = d2Top + Math.max(storeIssues.length, storePhotos.length, 1) + 2;
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
