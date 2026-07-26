// Dashboard.gs
// Builds and refreshes the "Dashboard" sheet. Layout:
//   - Navy header band: logo, title, last updated, refresh control
//   - Left sidebar (fixed rows): store/month filters, audit snapshot,
//     daily trend mini-table, how-to-use tips
//   - Main content (dynamic rows): KPI cards, two charts, paired tables
//   - Store deep-dive drill-through when a specific store is selected
//   - Export-to-PDF control at the bottom of whatever content exists
//
// Note: Sheets has no native rounded cards, shadows, or pill badges —
// these are approximated with colored cell backgrounds, and icons are emoji.

var MAIN_COL = 4;      // main content starts at column D
var PAIR_GAP = 1;      // spacing between a left item and its paired right item
var STAGING_COL = 40;  // hidden columns backing the two charts' data

function buildDashboardShell_(ss) {
  var sh = ss.getSheetByName(SHEETS.DASH) || ss.insertSheet(SHEETS.DASH);
  sh.clear();
  ss.setActiveSheet(sh); ss.moveActiveSheet(1);
}

// Draws a light border around a range so it reads as a distinct card.
function panelBorder_(range) {
  range.setBorder(true, true, true, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
}

function statusPill_(range, text, kind) {
  // kind: 'good' | 'fair' | 'poor' | 'neutral'
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

// Clears stray checkbox validations left behind by previous refreshes.
// The Export-to-PDF checkbox's row position shifts as content grows or
// shrinks, and dash.clear() does not remove data validation rules, so
// old positions would otherwise persist as blank checkboxes. The two
// active checkboxes (Refresh, Export) are redrawn immediately after this runs.
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

// Fires on manual edit. Handles the refresh checkbox, both filter
// dropdowns, and the Export-to-PDF checkbox (looked up from Document
// Properties since its position is dynamic — see refreshDashboard()).
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
  if (a1 === FILTER_CELL || a1 === MONTH_FILTER_CELL) {
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

function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(SHEETS.DASH) || ss.insertSheet(SHEETS.DASH);
  var master = ss.getSheetByName(SHEETS.MASTER);
  var issueSheet = ss.getSheetByName(SHEETS.ISSUES);

  var filterVal = 'All Stores';
  var monthVal = monthKey_(new Date());
  try { filterVal = dash.getRange(FILTER_CELL).getValue() || 'All Stores'; } catch (err) {}
  try { monthVal = dash.getRange(MONTH_FILTER_CELL).getValue() || monthVal; } catch (err) {}

  dash.getCharts().forEach(function (c) { dash.removeChart(c); });
  dash.setConditionalFormatRules([]);
  dash.clear();
  dash.setHiddenGridlines(true);
  clearAllCheckboxValidations_(dash);

  var now = new Date();
  var today = dayKey_(now);
  var yesterday = dayKey_(new Date(now.getTime() - 86400000));
  var overdueDays = Number(getSetting_('Overdue Days Threshold', CONFIG.OVERDUE_DAYS_DEFAULT));
  var targetScore = Number(getSetting_('Target Score %', 90));

  // At current store counts this full-table scan is instant. At much
  // larger scale (200+ stores, years of history), this should move to a
  // nightly trigger that pre-computes a summary sheet instead of scanning
  // raw Master Log / Issue Log on every refresh.
  var mRowsAll = (master && master.getLastRow() > 1)
    ? master.getRange(2, 1, master.getLastRow() - 1, 11).getValues() : [];
  // Columns: 0 ts, 1 date, 2 store, 3 am, 4 cluster, 5 itemsScored, 6 ones, 7 zeros, 8 scorePct, 9 sectionRemarks, 10 overallRemarks
  var iRowsAll = (issueSheet && issueSheet.getLastRow() > 1)
    ? issueSheet.getRange(2, 1, issueSheet.getLastRow() - 1, 11).getValues() : [];
  // Columns: 0 ts, 1 date, 2 store, 3 am, 4 section, 5 item, 6 status, 7 assignedTo, 8 dueDate, 9 closedDate, 10 severity

  var mRows = filterVal !== 'All Stores' ? mRowsAll.filter(function (r) { return r[2] === filterVal; }) : mRowsAll.slice();
  var iRows = filterVal !== 'All Stores' ? iRowsAll.filter(function (r) { return r[2] === filterVal; }) : iRowsAll.slice();
  var storesInScope = filterVal === 'All Stores' ? CONFIG.STORES : [filterVal];

  var selectedMonthRows = mRows.filter(function (r) { return monthKey_(r[1]) === monthVal; });
  var monthIssues = iRows.filter(function (r) { return monthKey_(r[1]) === monthVal; });
  var todayRows = mRows.filter(function (r) { return dayKey_(r[1]) === today; });
  var yesterdayRows = mRows.filter(function (r) { return dayKey_(r[1]) === yesterday; });
  var todaysIssues = iRows.filter(function (r) { return dayKey_(r[1]) === today; });

  var lastMonthPrefix = monthKey_(new Date(Number(monthVal.slice(0, 4)), Number(monthVal.slice(5, 7)) - 2, 1));
  var lastMonthRows = mRows.filter(function (r) { return monthKey_(r[1]) === lastMonthPrefix; });

  // --- Header band (rows 1-2, navy, spans the KPI row width) ---
  var headerLastCol = MAIN_COL + 6 * 3 - 1;
  for (var c = 1; c <= headerLastCol; c++) {
    dash.getRange(1, c).setBackground('#0f172a');
    dash.getRange(2, c).setBackground('#0f172a');
  }
  // Logo occupies columns 1-2 (see ensureLogoOnDashboard_ in Branding.gs).
  dash.getRange(1, 3, 2, 2).merge();
  dash.getRange(1, 5, 1, 6).merge().setValue('LIVE AUDIT DASHBOARD')
    .setFontColor('#ffffff').setFontWeight('bold').setFontSize(16).setVerticalAlignment('middle');
  dash.getRange(2, 5, 1, 6).merge().setValue('Real-time overview of store audit performance')
    .setFontColor('#94a3b8').setFontSize(9).setVerticalAlignment('middle');
  dash.getRange(1, 12, 1, headerLastCol - 11).merge()
    .setValue('Last Updated: ' + Utilities.formatDate(now, TZ, 'dd MMM yyyy, hh:mm a'))
    .setFontColor('#94a3b8').setFontSize(9).setHorizontalAlignment('right').setVerticalAlignment('middle');
  dash.getRange(2, 12, 1, headerLastCol - 15).merge()
    .setValue('🔄  Refresh Data').setBackground('#ffffff').setFontColor('#0f172a')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.getRange(REFRESH_CHECKBOX_CELL).insertCheckboxes().setValue(false).setBackground('#ffffff');
  dash.setRowHeight(1, 38); dash.setRowHeight(2, 28);

  // --- Left sidebar (fixed rows) ---
  dash.getRange('A4:B4').merge().setValue('STORE').setFontWeight('bold').setFontSize(9).setFontColor('#64748b');
  var storeDv = SpreadsheetApp.newDataValidation()
    .requireValueInList(['All Stores'].concat(CONFIG.STORES), true).setAllowInvalid(false).build();
  var currentStore = 'All Stores';
  try { currentStore = dash.getRange(FILTER_CELL).getValue() || 'All Stores'; } catch (err) {}
  dash.getRange('A5:B5').merge().setDataValidation(storeDv).setValue(currentStore)
    .setBackground('#ffffff').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a')
    .setVerticalAlignment('middle');
  dash.setRowHeight(5, 26);
  dash.getRange('A5:B5').setBorder(true, true, true, true, false, false, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  dash.getRange('A7:B7').merge().setValue('MONTH').setFontWeight('bold').setFontSize(9).setFontColor('#64748b');
  var months = getAvailableMonths_();
  var monthDv = SpreadsheetApp.newDataValidation()
    .requireValueInList(months, true).setAllowInvalid(false).build();
  var currentMonth = monthKey_(now);
  try {
    var existing = dash.getRange(MONTH_FILTER_CELL).getValue();
    if (existing && months.indexOf(existing) > -1) currentMonth = existing;
  } catch (err) {}
  dash.getRange('A8:B8').merge().setDataValidation(monthDv).setValue(currentMonth)
    .setBackground('#ffffff').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a')
    .setVerticalAlignment('middle');
  dash.setRowHeight(8, 26);
  dash.getRange('A8:B8').setBorder(true, true, true, true, false, false, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Audit Snapshot card
  dash.getRange('A10:B10').merge().setValue('AUDIT SNAPSHOT').setBackground('#0f172a')
    .setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  dash.getRange('A11:B11').merge().setValue('Today Overview').setBackground('#1e293b')
    .setFontColor('#cbd5e1').setFontSize(8);
  var storesVisitedToday = uniqueCount_(todayRows.map(function (r) { return r[2]; }));
  var snapshotRows = [
    ['Visits Today', String(todayRows.length)],
    ['Stores Visited', String(storesVisitedToday)],
    ['Avg Score Today', avgOf_(todayRows, 8) + '%'],
    ['Issues Raised', String(todaysIssues.length)]
  ];
  snapshotRows.forEach(function (row, i) {
    dash.getRange(12 + i, 1).setValue(row[0]).setFontColor('#64748b').setFontSize(9);
    dash.getRange(12 + i, 2).setValue(row[1]).setFontWeight('bold').setFontColor('#0f172a')
      .setHorizontalAlignment('right').setFontSize(9);
  });
  panelBorder_(dash.getRange('A10:B15'));

  // Daily Score Trend mini-table (most recent 10 days with data)
  var sidebarR = 17;
  dash.getRange(sidebarR, 1, 1, 2).merge().setValue('DAILY SCORE TREND')
    .setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
  sidebarR += 1;
  tableHeader_(dash, sidebarR, 1, ['Date', 'Avg Score %']);
  sidebarR += 1;
  var byDay = {};
  selectedMonthRows.forEach(function (row) {
    var d = dayKey_(row[1]);
    if (!byDay[d]) byDay[d] = [];
    if (row[8] !== '') byDay[d].push(Number(row[8]));
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

  // How To Use card
  dash.getRange(sidebarR, 1, 1, 2).merge().setValue('ℹ️  HOW TO USE')
    .setBackground('#eff6ff').setFontColor('#1e40af').setFontWeight('bold').setFontSize(10);
  sidebarR += 1;
  var tips = [
    'Use filters to view specific store or month performance.',
    'Pick a store above to see its full deep-dive at the bottom.',
    'Data refreshes automatically when filters change.'
  ];
  var tipsTop = sidebarR;
  tips.forEach(function (tip) {
    dash.getRange(sidebarR, 1, 1, 2).merge().setValue('•  ' + tip)
      .setBackground('#eff6ff').setFontColor('#1e3a8a').setFontSize(8).setWrap(true);
    dash.setRowHeight(sidebarR, 30);
    sidebarR += 1;
  });
  panelBorder_(dash.getRange(tipsTop - 1, 1, tips.length + 1, 2));

  dash.setColumnWidth(1, 150); dash.setColumnWidth(2, 90); dash.setColumnWidth(3, 20);
  for (var mc = MAIN_COL; mc <= MAIN_COL + 19; mc++) dash.setColumnWidth(mc, 95);

  // --- KPI cards (fixed rows 4-6, aligned with header/sidebar top) ---
  var scoresMonth = avgOf_(selectedMonthRows, 8);
  var scoresLastMonth = lastMonthRows.length ? avgOf_(lastMonthRows, 8) : null;
  var mtdDelta = scoresLastMonth !== null ? scoresMonth - scoresLastMonth : null;
  var mtdDeltaTxt = mtdDelta === null ? '—' : (mtdDelta >= 0 ? '▲ +' + mtdDelta : '▼ ' + mtdDelta) + '%';

  var avgToday = avgOf_(todayRows, 8);
  var avgYesterday = yesterdayRows.length ? avgOf_(yesterdayRows, 8) : null;
  var todayDelta = avgYesterday !== null ? avgToday - avgYesterday : null;
  var todayDeltaTxt = todayDelta === null ? '—' : (todayDelta >= 0 ? '▲ +' + todayDelta : '▼ ' + todayDelta) + '%';

  var monthlyVisitCount = selectedMonthRows.filter(function (r) { return r[8] !== ''; }).length;
  var meetingTarget = selectedMonthRows.filter(function (r) { return r[8] !== '' && Number(r[8]) >= targetScore; }).length;
  var targetAchievementPct = monthlyVisitCount ? Math.round(meetingTarget / monthlyVisitCount * 100) : 0;

  var perfectStores = storesInScope.filter(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    return vs.length && avgOf_(vs, 8) === 100;
  }).length;

  var highPriorityToday = todaysIssues.filter(function (x) { return x[10] === 'High'; }).length;

  var cards = [
    { icon: '🔵', label: 'VISITS TODAY', value: String(uniqueCount_(todayRows.map(function (r) { return r[2]; }))), sub: 'Stores Visited', color: '#2563eb', tint: '#eff6ff' },
    { icon: '✅', label: 'AVG SCORE TODAY', value: avgToday + '%', sub: 'vs Yesterday: ' + todayDeltaTxt, color: '#16a34a', tint: '#f0fdf4' },
    { icon: '📅', label: 'AVG SCORE (MTD)', value: scoresMonth + '%', sub: 'vs Last Month: ' + mtdDeltaTxt, color: '#16a34a', tint: '#f0fdf4' },
    { icon: '🎯', label: 'TARGET ACHIEVEMENT', value: targetAchievementPct + '%', sub: 'Target: ' + targetScore + '%', color: '#dc2626', tint: '#fef2f2' },
    { icon: '🏆', label: 'PERFECT SCORES', value: perfectStores + ' / ' + storesInScope.length, sub: 'Stores', color: '#7c3aed', tint: '#faf5ff' },
    { icon: '⚠️', label: 'OPEN ISSUES TODAY', value: String(todaysIssues.length), sub: 'High Priority: ' + highPriorityToday, color: '#dc2626', tint: '#fef2f2' }
  ];
  var cardRow = 4, cardW = 2;
  cards.forEach(function (card, i) {
    var c0 = MAIN_COL + i * (cardW + 1);
    dash.getRange(cardRow, c0, 1, cardW).merge().setBackground(card.tint)
      .setValue(card.icon + '  ' + card.label).setFontSize(9).setFontWeight('bold').setFontColor(card.color)
      .setVerticalAlignment('middle');
    dash.getRange(cardRow + 1, c0, 1, cardW).merge().setBackground(card.tint)
      .setValue(card.value).setFontSize(22).setFontWeight('bold').setFontColor('#0f172a')
      .setVerticalAlignment('middle');
    dash.getRange(cardRow + 2, c0, 1, cardW).merge().setBackground(card.tint)
      .setValue(card.sub).setFontSize(8).setFontColor('#475569').setVerticalAlignment('middle');
    var block = dash.getRange(cardRow, c0, 3, cardW);
    block.setBorder(true, false, true, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(cardRow, c0, 3, 1).setBorder(false, true, false, false, false, false, card.color, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  dash.setRowHeight(cardRow, 22); dash.setRowHeight(cardRow + 1, 32); dash.setRowHeight(cardRow + 2, 20);

  var r = cardRow + 4;

  // --- Two charts side by side: Score % Trend | Avg Score by Store ---
  var chartColSpan = 5;
  var chartsRightCol = MAIN_COL + chartColSpan + PAIR_GAP;
  dash.getRange(r, MAIN_COL).setValue('SCORE % TREND — ' + monthVal).setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, chartsRightCol).setValue('AVG SCORE BY STORE (' + monthVal + ')').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  var chartsRow = r + 1;

  // A brand-new sheet defaults to 26 columns, short of STAGING_COL (40),
  // so the sheet must be widened before writing to or hiding those columns.
  var neededCols = STAGING_COL + 6;
  if (dash.getMaxColumns() < neededCols) {
    dash.insertColumnsAfter(dash.getMaxColumns(), neededCols - dash.getMaxColumns());
  }

  // Hidden staging data backs both charts without cluttering the visible
  // layout (columns 40+ are hidden at the end of this function).
  var chartWidthPx = chartColSpan * 95;
  var trendStaging = [['Day', 'Avg Score %']].concat(trendOut);
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
      .setValue('No score data yet for ' + monthVal + '.')
      .setBackground('#f8fafc').setFontColor('#94a3b8').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    panelBorder_(dash.getRange(chartsRow, MAIN_COL, 8, chartColSpan));
  }

  var storeScoreOut = storesInScope.map(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    return [store, vs.length ? avgOf_(vs, 8) : 0];
  });
  var storeStaging = [['Store', 'Avg Score %']].concat(storeScoreOut);
  dash.getRange(chartsRow, STAGING_COL + 3, storeStaging.length, 2).setValues(storeStaging);
  var storeChart = dash.newChart().asColumnChart()
    .addRange(dash.getRange(chartsRow, STAGING_COL + 3, storeStaging.length, 2))
    .setPosition(chartsRow, chartsRightCol, 0, 0)
    .setOption('title', null).setOption('legend', { position: 'none' })
    .setOption('width', chartWidthPx).setOption('height', 250).setOption('colors', ['#dc2626'])
    .build();
  dash.insertChart(storeChart);

  // Top 5 Recurring Issues fills the space next to the second chart.
  var top5Col = chartsRightCol + chartColSpan + PAIR_GAP;
  dash.getRange(r, top5Col).setValue('TOP 5 RECURRING ISSUES — ' + monthVal)
    .setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  tableHeader_(dash, chartsRow - 1, top5Col, ['Rank', 'Section', 'Item', 'Times Failed'], 'danger');
  var counts5 = {};
  monthIssues.forEach(function (x) {
    var key = x[4] + ' || ' + x[5];
    counts5[key] = (counts5[key] || 0) + 1;
  });
  var top5 = Object.keys(counts5).map(function (k) {
    var parts = k.split(' || ');
    return [parts[0], parts[1], counts5[k]];
  }).sort(function (a, b) { return b[2] - a[2]; }).slice(0, 5)
    .map(function (row, idx) { return [idx + 1, row[0], row[1], row[2]]; });
  if (top5.length) {
    dash.getRange(chartsRow, top5Col, top5.length, 4).setValues(top5);
  } else {
    dash.getRange(chartsRow, top5Col, 1, 4).merge()
      .setValue('No failed items for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
  }
  panelBorder_(dash.getRange(chartsRow - 1, top5Col, Math.max(top5.length, 1) + 1, 4));

  r = chartsRow + 12; // clears the chart's visual height (250px ≈ 12 rows)

  // --- Row band: Today | Selected Month by Store | Store Ranking ---
  var todayRightCol = MAIN_COL + 5 + PAIR_GAP;
  var rankCol = todayRightCol + 5 + PAIR_GAP;
  dash.getRange(r, MAIN_COL).setValue('TODAY  (' + Utilities.formatDate(now, TZ, 'dd MMM yyyy') + ')')
    .setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, todayRightCol).setValue('SELECTED MONTH — AVG SCORE % BY STORE (' + monthVal + ')')
    .setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, rankCol).setValue('RANKING (' + monthVal + ')')
    .setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  r += 1;

  tableHeader_(dash, r, MAIN_COL, ['Store', 'Visit By', 'Cluster', 'Score %', 'Status']);
  tableHeader_(dash, r, todayRightCol, ['Store', 'Visits', 'Avg Score %', 'Total Score', 'Days ≥ Target']);
  tableHeader_(dash, r, rankCol, ['Rank', 'Store', 'Score %']);
  var pairTop = r + 1;

  var todayOut = storesInScope.map(function (store) {
    var v = todayRows.filter(function (x) { return x[2] === store; }).pop();
    if (!v) return [store, '—', '—', '—', 'NOT_VISITED'];
    var kind = v[8] >= targetScore ? 'good' : v[8] >= 70 ? 'fair' : 'poor';
    return [store, v[3], v[4], v[8] + '%', kind];
  });
  todayOut.forEach(function (row, i) {
    var rr = pairTop + i;
    dash.getRange(rr, MAIN_COL, 1, 4).setValues([[row[0], row[1], row[2], row[3]]]);
    var pillKind = row[4] === 'NOT_VISITED' ? 'neutral' : row[4];
    var pillText = row[4] === 'NOT_VISITED' ? 'Not Visited' : (row[4] === 'good' ? 'Good' : row[4] === 'fair' ? 'Fair' : 'Poor');
    statusPill_(dash.getRange(rr, MAIN_COL + 4), pillText, pillKind);
  });
  panelBorder_(dash.getRange(pairTop - 1, MAIN_COL, todayOut.length + 1, 5));

  var outB = storesInScope.map(function (store) {
    var vs = selectedMonthRows.filter(function (x) { return x[2] === store; });
    var visits = vs.length;
    var avg = visits ? avgOf_(vs, 8) : 0;
    var totalScore = vs.reduce(function (s, x) { return s + (Number(x[8]) || 0); }, 0);
    var daysAbove = vs.filter(function (x) { return Number(x[8]) >= targetScore; }).length;
    return [store, visits, visits ? avg + '%' : '—', totalScore, daysAbove];
  });
  dash.getRange(pairTop, todayRightCol, outB.length, 5).setValues(outB);
  panelBorder_(dash.getRange(pairTop - 1, todayRightCol, outB.length + 1, 5));

  var ranked = outB.filter(function (row) { return row[2] !== '—'; })
    .map(function (row) { return [row[0], Number(String(row[2]).replace('%', ''))]; })
    .sort(function (a, b) { return b[1] - a[1]; })
    .map(function (row, idx) { return [idx + 1, row[0], row[1] + '%']; });
  if (ranked.length) {
    dash.getRange(pairTop, rankCol, ranked.length, 3).setValues(ranked);
  } else {
    dash.getRange(pairTop, rankCol, 1, 3).merge()
      .setValue('No visits yet for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
  }
  panelBorder_(dash.getRange(pairTop - 1, rankCol, Math.max(ranked.length, 1) + 1, 3));

  r = pairTop + Math.max(todayOut.length, outB.length, ranked.length, 1) + 2;

  // --- Pair: Heat Map (left) | Today's Open Issues (right) ---
  var heatHeaders = ['Store'].concat(SECTIONS.map(function (s) { return s.name; })).concat(['Total Failures']);
  var heatRightCol = MAIN_COL + heatHeaders.length + PAIR_GAP;
  dash.getRange(r, MAIN_COL).setValue('HEAT MAP — FAILURE BY STORE & SECTION (' + monthVal + ')')
    .setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, heatRightCol).setValue("TODAY'S OPEN ISSUES")
    .setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  r += 1;

  tableHeader_(dash, r, MAIN_COL, heatHeaders);
  tableHeader_(dash, r, heatRightCol, ['Store', 'Section', 'Issue', 'Priority'], 'danger');
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
  var maxFailure = Math.max.apply(null, heatOut.map(function (row) { return Math.max.apply(null, row.slice(1)); }).concat([1]));
  var heatRule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#ffffff', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpointWithValue('#dc2626', SpreadsheetApp.InterpolationType.NUMBER, String(maxFailure))
    .setRanges([heatRange]).build();
  dash.setConditionalFormatRules(dash.getConditionalFormatRules().concat([heatRule]));
  dash.getRange(heatTop, MAIN_COL, heatOut.length, heatHeaders.length).setHorizontalAlignment('center');
  panelBorder_(dash.getRange(heatTop - 1, MAIN_COL, heatOut.length + 1, heatHeaders.length));

  if (todaysIssues.length) {
    todaysIssues.forEach(function (x, i) {
      var rr = heatTop + i;
      dash.getRange(rr, heatRightCol, 1, 3).setValues([[x[2], x[4], x[5]]]);
      var sev = x[10] || 'Low';
      var kind = sev === 'High' ? 'poor' : sev === 'Medium' ? 'fair' : 'good';
      statusPill_(dash.getRange(rr, heatRightCol + 3), sev, kind);
    });
  } else {
    dash.getRange(heatTop, heatRightCol, 1, 4).merge()
      .setValue('No open issues today 🎉').setFontSize(9).setFontColor('#94a3b8');
  }
  panelBorder_(dash.getRange(heatTop - 1, heatRightCol, Math.max(todaysIssues.length, 1) + 1, 4));

  r = heatTop + Math.max(heatOut.length, todaysIssues.length, 1) + 2;

  // --- Drill-through: Store Deep-Dive (only when one store is selected) ---
  var deepRightCol = MAIN_COL + 3 + PAIR_GAP;
  if (filterVal !== 'All Stores') {
    dash.getRange(r, MAIN_COL, 1, 6).merge().setValue('🔍  STORE DEEP-DIVE — ' + filterVal)
      .setBackground('#0f172a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12)
      .setVerticalAlignment('middle');
    dash.setRowHeight(r, 26);
    r += 2;

    dash.getRange(r, MAIN_COL).setValue('Visit History').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
    dash.getRange(r, deepRightCol).setValue('Open Issues').setFontWeight('bold').setFontSize(10).setFontColor('#0f172a');
    r += 1;
    tableHeader_(dash, r, MAIN_COL, ['Date', 'Visit By', 'Score %']);
    tableHeader_(dash, r, deepRightCol, ['Section', 'Open Issues']);
    var deepTop = r + 1;

    var visitHistory = selectedMonthRows.slice().sort(function (a, b) { return b[1] - a[1]; })
      .map(function (x) { return [Utilities.formatDate(new Date(x[1]), TZ, 'dd MMM yyyy'), x[3], x[8] + '%']; });
    if (visitHistory.length) {
      dash.getRange(deepTop, MAIN_COL, visitHistory.length, 3).setValues(visitHistory);
    } else {
      dash.getRange(deepTop, MAIN_COL, 1, 3).merge()
        .setValue('No visits for ' + monthVal + '.').setFontSize(9).setFontColor('#94a3b8');
    }
    panelBorder_(dash.getRange(deepTop - 1, MAIN_COL, Math.max(visitHistory.length, 1) + 1, 3));

    var bySection = SECTIONS.map(function (sec) {
      return [sec.name, monthIssues.filter(function (x) { return x[4] === sec.name; }).length];
    }).filter(function (row) { return row[1] > 0; });
    if (bySection.length) {
      dash.getRange(deepTop, deepRightCol, bySection.length, 2).setValues(bySection);
    } else {
      dash.getRange(deepTop, deepRightCol, 1, 2).merge()
        .setValue('No open issues. 🎉').setFontSize(9).setFontColor('#94a3b8');
    }
    var totalOpenRow = deepTop + Math.max(bySection.length, 1);
    dash.getRange(totalOpenRow, deepRightCol).setValue('Total Open Issues').setFontWeight('bold').setFontSize(9);
    dash.getRange(totalOpenRow, deepRightCol + 1).setValue(monthIssues.length).setFontWeight('bold').setFontSize(9);
    panelBorder_(dash.getRange(deepTop - 1, deepRightCol, Math.max(bySection.length, 1) + 2, 2));

    r = deepTop + Math.max(visitHistory.length, bySection.length + 1, 1) + 2;
  }

  // --- Export to PDF (position is dynamic — stored for onEdit() to find) ---
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

  // --- Final polish: divider so the sidebar reads as one continuous panel ---
  dash.getRange(4, 3, Math.max(r - 3, 1), 1)
    .setBorder(false, false, false, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

  dash.hideColumns(STAGING_COL, 6);
  dash.setFrozenRows(2);
  ensureLogoOnDashboard_(dash);
}

// --- Stats helpers ---
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
