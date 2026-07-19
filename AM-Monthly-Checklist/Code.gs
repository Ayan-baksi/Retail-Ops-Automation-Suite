/*************************************************************************
* Appsscript.json
*************************************************************************/
{
  "timeZone": "Asia/Kolkata",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/forms",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}

/*************************************************************************
 * Branding.gs
 * Applies a custom logo to the Google Form header and the Dashboard sheet.
 * Safe to re-run — will not create duplicate logos.
 *
 * Setup: Paste your logo's Google Drive file ID into CONFIG.LOGO_FILE_ID
 * (in Config.gs), then run "🖼️ Apply/Refresh Logo" from the AM Tools menu.
 *************************************************************************/

function applyLogo() {
  var ui = SpreadsheetApp.getUi();
  if (!CONFIG.LOGO_FILE_ID) {
    ui.alert('No logo file ID set. Paste your logo\'s Drive file ID into CONFIG.LOGO_FILE_ID in Config.gs, then run this again.');
    return;
  }
  var formId = PropertiesService.getDocumentProperties().getProperty('FORM_ID');
  if (!formId) {
    ui.alert('Run "Setup System" first — no form exists yet.');
    return;
  }

  var blob;
  try {
    blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
  } catch (err) {
    ui.alert('Could not open that file. Check the file ID in Config.gs and confirm the file still exists in Drive.');
    return;
  }

  FormApp.openById(formId).setImage(blob);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName(SHEETS.DASH);
  dash.getImages().forEach(function (img) { img.remove(); });
  var img = dash.insertImage(blob, 1, 1);
  img.setWidth(50).setHeight(50);

  ui.alert('Done — logo applied to the Form header and Dashboard.');
}

// Re-applies the logo after a full dashboard rebuild, without duplicating it.
function ensureLogoOnDashboard_(dash) {
  if (!CONFIG.LOGO_FILE_ID) return;
  if (dash.getImages().length > 0) return;
  try {
    var blob = DriveApp.getFileById(CONFIG.LOGO_FILE_ID).getBlob();
    var img = dash.insertImage(blob, 1, 1);
    img.setWidth(50).setHeight(50);
  } catch (err) {
    // Logo not configured yet, or file inaccessible — dashboard still works without it.
  }
}

/*************************************************************************
 * Config.gs — DEEHUB LIFESTYLE RETAIL AUDIT SYSTEM
 * -----------------------------------------------------------------------
 * Single source of truth for everything the other files need. Nothing in
 * this file talks to Forms/Sheets directly — it's pure configuration.
 *************************************************************************/

var CONFIG = {
  ORG_NAME: 'DeeHub Lifestyle',
  APP_VERSION: 'v1.0',

  FORM_TITLE: 'DeeHub Lifestyle — Area Manager Store Audit',
  FORM_DESC: 'Welcome. This checklist takes about 10–15 minutes per store visit. ' +
    'Your responses are recorded instantly and feed the live performance dashboard automatically — ' +
    'no separate reporting needed.',
  CONFIRMATION_MSG: 'Thank you — your audit has been recorded and the DeeHub dashboard has been updated.',

  // >>> Paste your logo's Google Drive file ID here, then run "Apply/Refresh Logo" from AM Tools <<<
  // Get it from: right-click the file in Drive → Get link → the ID is the long string between /d/ and /view
  LOGO_FILE_ID: '',

  // >>> Edit this list if your store footprint changes <<<
  STORES: [
    'STORE-001', 'STORE-002', 'STORE-003', 'STORE-004', 'STORE-005',
    'STORE-006', 'STORE-007', 'STORE-008', 'STORE-009', 'STORE-010'
  ],

  SCORE_COLS: ['1', '0'],          // 1 = compliant, 0 = not compliant, blank = N/A
  OVERDUE_DAYS_DEFAULT: 7          // fallback if Settings sheet value is missing
};

// Sheet names — change here once, every file picks it up
var SHEETS = {
  MASTER: 'Master Log',
  ISSUES: 'Issue Log',
  DASH: 'Dashboard',
  STORE_MASTER: 'Store_Master',
  EMPLOYEE_MASTER: 'Employee_Master',
  CHECKLIST_MASTER: 'Checklist_Master',
  SETTINGS: 'Settings'
};

// Form question titles — used both when building the form and when
// parsing submissions, so they always stay in sync.
var Q = {
  AM: 'AM Name',
  CLUSTER: 'Cluster',
  STORE: 'Store',
  DATE: 'Date of Visit',
  FINAL_REMARKS: 'Overall Remarks for This Visit'
};

var TZ = Session.getScriptTimeZone();

// Dashboard control cells (see Dashboard.gs)
var FILTER_CELL = 'A5';           // Store dropdown, fixed in the left sidebar
var MONTH_FILTER_CELL = 'A8';     // Month dropdown, fixed in the left sidebar
var REFRESH_CHECKBOX_CELL = 'R2'; // Refresh trigger, sits inside the navy header band

/* ===================== CHECKLIST DEFINITION =====================
 * Exactly the items from your original AM Checklist workbook — same
 * wording, same order, nothing added or removed. Each section gets a
 * short standard instruction line when the Form is built (see
 * FormBuilder.gs) rather than repeating similar text 7 times here.
 * =================================================================== */
var SECTIONS = [
  { name: 'BRANDING', items: [
    'Façade clean & no faulty lights',
    'Internal/External branding - Pillar, Stairs, Directory etc - in good condition',
    'Consumer touchpoints maintained as per standard',
    'Branding - Pillar, Stairs, Directory etc - in good condition',
    'Baggae Counter - ensure token available, neat & clean , dust free'
  ]},
  { name: 'PRODUCT & DISPLAY', items: [
    'Full options are in front - wall/browsers & tables',
    'Full size set styles to be displayed option wise NOT size wise',
    'Assorted merchandise - Department >colour>size wise',
    'All tables are 100% stocked & options are merchandised as per barcode list',
    '100% received styles are on display ?',
    'Promo & TSO docket 100% adherence',
    'Excess/base stocks are kept below same hanging/folding',
    'Same style different colours/fabric/embroidery etc are kept together - hanging/folding',
    'Designs to be merchandised together - solids/checks/stripes/prints etc',
    'Price point flow from left to right',
    'Slat wall -  all impulse hanging & 100% stocked',
    'Are proper signage in the section maintained as per offers',
    'Staircase is free of bins & merchandise',
    'No apparels to be kept in bins',
    'Docket maintained - Impulse/lingerie/GM/tables &  current season apparels wall elevation',
    'Broken- One size per Fixture- (Browsers one size per arm/I way Two size max) with size indicators',
    'Hygiene Check: Dusting on shelves, hangers, under fixture cleanliness',
    '100% hangers as per SOP - if NO then mention timeline',
    'Floor fixtures placement & wall as per layout',
    'Every two hours team to check base stocks for replenishment',
    'Every hour merchandise to be arranged as per option in 100% fixture',
    'High Week cover- Action plan executed',
    'Low week cover-  Intransit status',
    'Low/Zero sell thru options - action plan executed',
    'Ageing inventory- action plan executed',
    'Layout - Departments & fixtures as per layout'
  ]},
  { name: 'Customers Engagaement', items: [
    'Tele calling coverage as per target- minimum 50 cusomers per day',
    'Marketing activity adherence as per plan- weak stores',
    'Google review - minimum 50% of the bills everyday- QR code placed as per guidelines',
    'Instagram QR code palced in trial room'
  ]},
  { name: 'PROCESS', items: [
    'SM walk adherence (random check - checklist vs floor)',
    '100% tagging & pc to pc GRC',
    'Cash Reconciliation, Banking, manual bills,  & Cash Expense Details',
    'Customer Return and gate pass entry as per norms',
    'Sensormatic Gate & Tags Verification',
    'Global Count (System vs physical) of High Value items - 8-10 departments > 0.3%',
    'Check for availability of HK/Security as per budget; Quality of standards',
    'SM doing Business Review with Staff & Managers- review minutes/action plan',
    'No merchandise and shopping bags lying behind cash tills and in trial rooms('
  ]},
  { name: 'STORE WH', items: [
    'WH setup as per defined SOP- Season>Division>Section>Department>Options>Assorted (size wise)',
    'Replenishment qty as per sales vs WH stocks- 100% check sample styles',
    'Proper labelling and segregation of  merchandise',
    'Damage/ Defective List Review % of overall value',
    'IST plan vs actual executed 100%'
  ]},
  { name: 'PEOPLE', items: [
    'Total Manpower Strength- MPP vs actual- 95%',
    'Staff Grooming as per defined Standards',
    'Target Awareness of each staff - Monthly/Weekly; Value & Qty',
    'Staff Bill mix% (low /zero staff sales review)- 75%',
    'Training Conducted - SOP/ Product/ Business/ Process for 10-15 mins',
    'Daily customer swagat & staff briefing happening'
  ]},
  { name: 'HYGIENE & MAINTENANCE', items: [
    'Store parking area & entrace - neat & clean',
    'DG area - neat & clean and clutter free',
    'Staircase at main entrance & inside must be dust free.',
    'Store toilet cleaing checklist is in place and cleaning happening every hours',
    'All light are in working condition and if not then escalate the same to maintenance team with timeline',
    'AC is properly running and store temperature is pleasant(if not then escalate the same to maintenance team with timeline)'
  ]}
];

/*************************************************************************
 * Dashboard.gs — builds and refreshes the "Dashboard" tab, styled to
 * match MASTER_DASH_BOARD_FOR_DEEHUB.png as closely as Google Sheets
 * can actually render:
 *   - Navy header band (logo, title, last updated, refresh control)
 *   - Left sidebar (Store/Month filters, Audit Snapshot, Daily Trend
 *     mini-table, How To Use tips) — all FIXED rows, since this content
 *     never changes size
 *   - Main content area (KPI cards, two charts side by side, then
 *     paired tables) — position is DYNAMIC below the KPI cards, since
 *     table height depends on how much data exists
 *   - Store Deep-Dive drill-through at the very bottom when a specific
 *     store is selected
 *   - An Export-to-PDF control at the true bottom of whatever content
 *     exists that refresh (see exportDashboardToPdf in Export.gs)
 *
 * HONEST LIMITATIONS: Sheets can't do rounded cards, drop shadows, or
 * true pill-shaped badges — these are approximated with colored cell
 * backgrounds. Icons are emoji, not vector graphics.
 *************************************************************************/

var MAIN_COL = 4;               // main content starts at column D
var PAIR_GAP = 1;                // columns of breathing room between a left item and its paired right item
var STAGING_COL = 40;           // far-right hidden columns that back the two charts' data

function buildDashboardShell_(ss) {
  var sh = ss.getSheetByName(SHEETS.DASH) || ss.insertSheet(SHEETS.DASH);
  sh.clear();
  ss.setActiveSheet(sh); ss.moveActiveSheet(1);
}

// Light gray box around a block of cells so each section reads as a
// distinct card rather than plain sheet rows.
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

// Removes every checkbox-type data validation on the sheet. Needed
// because the Export-to-PDF checkbox's position moves every refresh
// (it sits at the dynamic "true bottom" of the content) — dash.clear()
// does NOT remove data validation rules, so without this sweep, every
// past position becomes a permanent stray blank checkbox. The two real
// checkboxes (Refresh + Export) get redrawn fresh immediately after this
// runs, so clearing everything first is safe.
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

// Simple trigger — fires on any manual edit. Catches the refresh
// checkbox, both filter dropdowns, and the Export-to-PDF checkbox
// (whose position is dynamic, so it's looked up from Document
// Properties rather than a fixed constant — see refreshDashboard()).
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
  clearAllCheckboxValidations_(dash); // sweeps up any stray checkbox left behind by a past refresh, since the Export button's position moves as data grows/shrinks

  var now = new Date();
  var today = dayKey_(now);
  var yesterday = dayKey_(new Date(now.getTime() - 86400000));
  var overdueDays = Number(getSetting_('Overdue Days Threshold', CONFIG.OVERDUE_DAYS_DEFAULT));
  var targetScore = Number(getSetting_('Target Score %', 90));

  // NOTE ON SCALE: at 10 stores this full-table read+scan is instant.
  // Once you're at 200+ stores with years of history, switch this to a
  // nightly trigger that pre-computes a small "Daily Summary" sheet
  // instead of scanning raw Master Log/Issue Log every time.
  var mRowsAll = (master && master.getLastRow() > 1)
    ? master.getRange(2, 1, master.getLastRow() - 1, 11).getValues() : [];
  // 0 ts,1 date,2 store,3 am,4 cluster,5 itemsScored,6 ones,7 zeros,8 scorePct,9 sectionRemarks,10 overallRemarks
  var iRowsAll = (issueSheet && issueSheet.getLastRow() > 1)
    ? issueSheet.getRange(2, 1, issueSheet.getLastRow() - 1, 11).getValues() : [];
  // 0 ts,1 date,2 store,3 am,4 section,5 item,6 status,7 assignedTo,8 dueDate,9 closedDate,10 severity

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

  /* =====================================================================
   * HEADER BAND (rows 1–2, navy, spans the full width of the KPI row)
   * ===================================================================== */
  var headerLastCol = MAIN_COL + 6 * 3 - 1; // matches the KPI card row width
  for (var c = 1; c <= headerLastCol; c++) {
    dash.getRange(1, c).setBackground('#0f172a');
    dash.getRange(2, c).setBackground('#0f172a');
  }
  // Logo lives in columns 1–2 (see ensureLogoOnDashboard_ in Branding.gs) —
  // its wordmark already says "deehub lifestyle", so no separate text here.
  dash.getRange(1, 3, 2, 2).merge(); // reserved blank space for the logo to float over
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

  /* =====================================================================
   * LEFT SIDEBAR (fixed rows — this content never changes size)
   * ===================================================================== */
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

  /* =====================================================================
   * MAIN CONTENT — KPI CARDS (fixed rows 4–6, aligned with header/sidebar top)
   * ===================================================================== */
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

  var r = cardRow + 4; // one blank row after the cards

  /* =====================================================================
   * TWO CHARTS SIDE BY SIDE: Score % Trend  |  Avg Score by Store
   * ===================================================================== */
  var chartColSpan = 5;                                  // columns of grid-space each chart visually occupies
  var chartsRightCol = MAIN_COL + chartColSpan + PAIR_GAP; // right chart starts right after the left one
  dash.getRange(r, MAIN_COL).setValue('SCORE % TREND — ' + monthVal).setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  dash.getRange(r, chartsRightCol).setValue('AVG SCORE BY STORE (' + monthVal + ')').setFontWeight('bold').setFontSize(11).setFontColor('#0f172a');
  var chartsRow = r + 1;

  // A brand-new sheet only has 26 columns by default — STAGING_COL (40)
  // is beyond that, so it must be widened first or every write/hide call
  // out there throws "columns are out of bounds".
  var neededCols = STAGING_COL + 6;
  if (dash.getMaxColumns() < neededCols) {
    dash.insertColumnsAfter(dash.getMaxColumns(), neededCols - dash.getMaxColumns());
  }

  // Hidden staging data so both charts have a real backing range without
  // cluttering the visible layout (columns 40+ are hidden at the end).
  var chartWidthPx = chartColSpan * 95; // matches the 95px column width set above
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

  // Top 5 Recurring Issues fills the empty space to the right of the
  // second chart, instead of leaving that whole area blank.
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

  r = chartsRow + 12; // tight row-advance to clear the chart's visual height (250px ≈ 12 rows)

  /* =====================================================================
   * ROW BAND: TODAY | SELECTED MONTH BY STORE | STORE RANKING
   * ===================================================================== */
  var todayRightCol = MAIN_COL + 5 + PAIR_GAP; // Today table is 5 columns wide
  var rankCol = todayRightCol + 5 + PAIR_GAP;  // Selected Month table is also 5 columns wide
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

  /* =====================================================================
   * PAIR: HEAT MAP (left)  |  TODAY'S OPEN ISSUES (right)
   * ===================================================================== */
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

  /* =====================================================================
   * DRILL-THROUGH: STORE DEEP-DIVE (only when one store is selected)
   * ===================================================================== */
  var deepRightCol = MAIN_COL + 3 + PAIR_GAP; // visit history table is 3 columns wide
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

  /* =====================================================================
   * EXPORT TO PDF — sits at the true bottom of whatever content exists,
   * left-aligned under the main content (not off in a distant column).
   * Position is dynamic, so it's stored in Document Properties for
   * onEdit() to find (see top of this file).
   * ===================================================================== */
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

  /* =====================================================================
   * FINAL POLISH
   * ===================================================================== */
  // A thin divider running the full content height so the sidebar reads
  // as one continuous panel instead of visibly stopping short of the
  // (much longer) main content next to it.
  dash.getRange(4, 3, Math.max(r - 3, 1), 1)
    .setBorder(false, false, false, true, false, false, '#e2e8f0', SpreadsheetApp.BorderStyle.SOLID);

  dash.hideColumns(STAGING_COL, 6);
  dash.setFrozenRows(2);
  ensureLogoOnDashboard_(dash);
}

/* ============================ small stats helpers ============================ */
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

// Distinct yyyy-MM values seen in Master Log, newest first, with the
// current month always included even if it has no data yet.
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

/*************************************************************************
 * Export.gs — exports the Dashboard tab to PDF and saves it to Drive.
 * Triggered by the red "EXPORT TO PDF" checkbox at the bottom of the
 * Dashboard (see onEdit() in Dashboard.gs), or run manually any time.
 *
 * IF EXPORT FAILS WITH A PERMISSIONS ERROR (script.external_request):
 * Your project was authorized before this function needed UrlFetchApp —
 * so Google never asked for that specific permission. Fix:
 *   1. In the Apps Script editor: Project Settings (gear icon) → check
 *      "Show 'appsscript.json' manifest file in editor".
 *   2. Open appsscript.json → make sure an "oauthScopes" array is NOT
 *      restricting things (delete it if present, or use the version
 *      below) → Save.
 *   3. Go to https://myaccount.google.com/permissions → find this
 *      project → Remove Access (this forces a fresh, complete consent
 *      screen next time, instead of reusing the old partial approval).
 *   4. Function dropdown → exportDashboardToPdf → Run → approve ALL
 *      permissions this time (Sheets, Drive, AND external requests).
 *************************************************************************/

function exportDashboardToPdf() {
  var ui = SpreadsheetApp.getUi();
  var log = [];

  try {
    log.push('Step 1: Locating spreadsheet and Dashboard sheet...');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dash = ss.getSheetByName(SHEETS.DASH);
    if (!dash) throw new Error('Could not find a sheet named "' + SHEETS.DASH + '". Has it been renamed?');
    var ssId = ss.getId();
    var gid = dash.getSheetId();
    log.push('  Spreadsheet ID: ' + ssId + ' | Dashboard GID: ' + gid);

    log.push('Step 2: Building export URL...');
    var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export' +
      '?format=pdf&gid=' + gid +
      '&size=A4&portrait=false&fitw=true' +
      '&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3' +
      '&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false&horizontal_alignment=CENTER';
    log.push('  URL: ' + url);

    log.push('Step 3: Fetching OAuth token...');
    var token = ScriptApp.getOAuthToken();
    if (!token) throw new Error('Got an empty OAuth token — authorization likely was not completed.');

    log.push('Step 4: Requesting the PDF from Google (UrlFetchApp)...');
    var response;
    try {
      response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
    } catch (fetchErr) {
      throw new Error('UrlFetchApp.fetch itself threw an error — this is almost always the missing '
        + '"script.external_request" scope. See the fix steps in the comment at the top of Export.gs. Raw error: ' + fetchErr.message);
    }

    var code = response.getResponseCode();
    log.push('  HTTP response code: ' + code);
    if (code !== 200) {
      var bodySnippet = response.getContentText().substring(0, 300);
      throw new Error('Google returned HTTP ' + code + ' instead of 200. This usually means the GID is wrong, '
        + 'or the spreadsheet couldn\'t be read with this token. Response snippet: ' + bodySnippet);
    }

    log.push('Step 5: Building the PDF file blob...');
    var fileName = CONFIG.ORG_NAME + ' Dashboard - ' + Utilities.formatDate(new Date(), TZ, 'dd-MMM-yyyy HHmm') + '.pdf';
    var blob = response.getBlob().setName(fileName);
    var blobSize = blob.getBytes().length;
    log.push('  Blob size: ' + blobSize + ' bytes');
    if (blobSize < 1000) {
      throw new Error('The PDF came back suspiciously small (' + blobSize + ' bytes) — likely an error page instead of a real PDF, not a genuine export.');
    }

    log.push('Step 6: Finding or creating the Drive export folder...');
    var folder = getOrCreateExportFolder_();
    log.push('  Folder: "' + folder.getName() + '" (ID: ' + folder.getId() + ')');

    log.push('Step 7: Saving the PDF into that folder...');
    var file = folder.createFile(blob);
    log.push('  Saved as: ' + file.getName() + ' (File ID: ' + file.getId() + ')');

    Logger.log(log.join('\n'));
    ui.alert('Exported ✅', 'Saved to your Drive folder "DeeHub Dashboard Exports" as:\n' + file.getName(), ui.ButtonSet.OK);

  } catch (err) {
    Logger.log('EXPORT FAILED\n' + log.join('\n') + '\n\nError: ' + err.message);
    ui.alert('Export failed',
      'Something went wrong:\n' + err.message +
      '\n\nFull step-by-step log has been written to the Apps Script Execution log ' +
      '(Executions in the left sidebar) so you can see exactly which step failed.',
      ui.ButtonSet.OK);
  }
}

function getOrCreateExportFolder_() {
  var name = 'DeeHub Dashboard Exports';
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

// Run this directly from the Apps Script editor (not the checkbox) the
// FIRST time, after fixing permissions per the note at the top of this
// file — it forces the consent screen to appear for every scope this
// project needs, in one shot, instead of one at a time.
function authorizeAllPermissions() {
  SpreadsheetApp.getActiveSpreadsheet().getName();      // Sheets scope
  DriveApp.getRootFolder().getName();                    // Drive scope
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); // external_request scope
  Logger.log('If this ran with no error, all required permissions are granted.');
}

/*************************************************************************
 * FormBuilder.gs — builds the Google Form described in the spec:
 * Welcome → Store Information → one page per checklist section →
 * Final Remarks → Confirmation. Each section page carries a title and
 * a description/instructions line (folded into one help-text block,
 * rather than three separate fields — Forms doesn't have a distinct
 * "instructions" field type, so this is the cleanest equivalent).
 *************************************************************************/

function buildForm_() {
  var form = FormApp.create(CONFIG.FORM_TITLE);
  form.setDescription(CONFIG.FORM_DESC);
  form.setCollectEmail(false);      // per your call: free-text identity, no org restriction
  form.setProgressBar(true);
  form.setConfirmationMessage(CONFIG.CONFIRMATION_MSG);
  form.setShowLinkToRespondAgain(true);

  // TODO once you have a logo file: form.setImage(DriveApp.getFileById('LOGO_FILE_ID').getBlob())
  // Ask me to wire this in as soon as you upload the actual logo image.

  // --- Store Information page ---
  form.addPageBreakItem()
    .setTitle('Store information')
    .setHelpText('Tell us who is visiting and where, before the checklist begins.');
  form.addTextItem().setTitle(Q.AM).setRequired(true);
  form.addTextItem().setTitle(Q.CLUSTER).setRequired(true);
  form.addListItem().setTitle(Q.STORE).setChoiceValues(CONFIG.STORES).setRequired(true);
  form.addDateItem().setTitle(Q.DATE).setRequired(true);

  // --- One page per checklist section ---
  SECTIONS.forEach(function (sec) {
    form.addPageBreakItem()
      .setTitle(sec.name)
      .setHelpText('Score each item 1 (compliant) or 0 (not compliant). Leave a row blank only if it genuinely does not apply to this store.');
    form.addGridItem()
      .setTitle(sec.name + ' — checklist')
      .setRows(sec.items)
      .setColumns(CONFIG.SCORE_COLS)
      .setRequired(false);
    form.addParagraphTextItem().setTitle(sec.name + ' — Remarks (optional)');
  });

  // --- Final remarks + confirmation ---
  form.addPageBreakItem()
    .setTitle('Final remarks')
    .setHelpText('Anything else worth flagging about this visit overall?');
  form.addParagraphTextItem().setTitle(Q.FINAL_REMARKS);

  return form;
}

/*************************************************************************
 * MasterData.gs — builds the four "master" reference sheets the spec
 * calls for. These are what let Cluster/Region/RM/VM filters, checklist
 * item weighting, and non-technical settings changes exist someday
 * without rewriting code — the actual filter/weighting LOGIC in
 * Dashboard.gs doesn't use these yet (you told me to skip that for now),
 * but the data model is ready the moment you want to turn it on.
 *************************************************************************/

function buildAllMasterData_(ss) {
  buildStoreMaster_(ss);
  buildEmployeeMaster_(ss);
  buildChecklistMaster_(ss);
  buildSettings_(ss);
}

// One row per store. Cluster/Region/RM/VM start blank — fill these in
// once and every future filter/report can use them.
function buildStoreMaster_(ss) {
  var sh = ss.getSheetByName(SHEETS.STORE_MASTER) || ss.insertSheet(SHEETS.STORE_MASTER);
  var h = ['Store Code', 'Store Name', 'Cluster', 'Region', 'State',
    'Area Manager', 'Regional Manager', 'Visual Merchandiser', 'Status'];
  if (sh.getLastRow() > 0) return; // don't overwrite if you've already filled this in
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

// Blank template — add one row per Area/Regional Manager or Visual
// Merchandiser as your team grows.
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
// Weight defaults to 1 (equal weighting) — bump individual weights later
// if you decide some sections should count more toward the final score.
function buildChecklistMaster_(ss) {
  var sh = ss.getSheetByName(SHEETS.CHECKLIST_MASTER) || ss.insertSheet(SHEETS.CHECKLIST_MASTER);
  sh.clear(); // safe to rebuild — this sheet just mirrors Config.gs, no manual data lives here yet
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
  sh.getRange('A1').setNote('This mirrors SECTIONS in Config.gs. Editing Weight here has no effect yet — weighting is not wired into the dashboard until you ask for it.');
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

/*************************************************************************
 * Menu.gs — the "AM Tools" menu (spec calls for this exact set of
 * actions) plus the testing helpers behind two of its items.
 *************************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AM Tools')
    .addItem('⚙️ Setup System (first time only)', 'setup')
    .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
    .addSeparator()
    .addItem('🧪 Seed 15 Sample Visits (testing)', 'seedTestData')
    .addItem('🗑️ Reset (clear all responses + test data)', 'clearAllResponses')
    .addSeparator()
    .addItem('🖼️ Apply/Refresh Logo', 'applyLogo')
    .addItem('🔗 Get Form Link', 'getFormLink')
    .addToUi();
}

function clearAllResponses() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Clear all responses?',
    'This permanently deletes every row in Master Log, Issue Log, and Form Responses (headers kept), then refreshes the dashboard. Continue?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.MASTER));
  clearSheetKeepHeader_(ss.getSheetByName(SHEETS.ISSUES));
  ss.getSheets().forEach(function (s) {
    if (s.getName().indexOf('Form Responses') === 0) clearSheetKeepHeader_(s);
  });
  refreshDashboard();
  ui.alert('Cleared. Dashboard refreshed.');
}

// Generates realistic dummy visits (with some deliberate failures) so
// every KPI card, both charts, the ranking, and the issue list all have
// something meaningful to show while you're testing.
function seedTestData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName(SHEETS.MASTER);
  var issueSheet = ss.getSheetByName(SHEETS.ISSUES);
  var today = new Date();
  var count = 15;
  var issueRows = [];
  var totalItems = SECTIONS.reduce(function (s, sec) { return s + sec.items.length; }, 0);

  for (var i = 0; i < count; i++) {
    var d = new Date(today.getFullYear(), today.getMonth(), 1 + Math.floor(Math.random() * 27));
    var store = CONFIG.STORES[Math.floor(Math.random() * CONFIG.STORES.length)];
    var zeros = Math.floor(Math.random() * 12);
    var ones = totalItems - zeros;
    var scorePct = Math.round(ones / totalItems * 100);

    master.appendRow([new Date(), d, store, 'Test AM ' + ((i % 3) + 1), 'Cluster ' + ((i % 2) + 1),
      totalItems, ones, zeros, scorePct,
      i % 4 === 0 ? 'Sample section remark' : '', i % 5 === 0 ? 'Sample overall remark' : '']);

    for (var z = 0; z < zeros; z++) {
      var sec = SECTIONS[Math.floor(Math.random() * SECTIONS.length)];
      var item = sec.items[Math.floor(Math.random() * sec.items.length)];
      issueRows.push([new Date(), d, store, 'Test AM ' + ((i % 3) + 1), sec.name, item, 'Open', '', '', '', '']);
    }
  }
  if (issueRows.length) {
    issueSheet.getRange(issueSheet.getLastRow() + 1, 1, issueRows.length, issueRows[0].length).setValues(issueRows);
  }
  refreshDashboard();
  SpreadsheetApp.getUi().alert('Seeded ' + count + ' sample visits and refreshed the dashboard.');
}

function getFormLink() {
  var id = PropertiesService.getDocumentProperties().getProperty('FORM_ID');
  if (!id) { Logger.log('No form found yet. Run setup() first.'); return; }
  var form = FormApp.openById(id);
  var msg = 'Fill link (share with Area Managers):\n' + form.getPublishedUrl() +
    '\n\nEdit link (keep for yourself only):\n' + form.getEditUrl();
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) { /* fine if run from the script editor */ }
}

  /*************************************************************************
 * Setup.gs — the one function you actually run by hand: setup(). It
 * calls into every other file in the right order. Everything after
 * that runs itself via triggers and the AM Tools menu.
 *
 * FIRST-TIME USE
 *   1. Make sure Config.gs, Utils.gs, MasterData.gs, FormBuilder.gs,
 *      SubmitHandler.gs, Dashboard.gs, Menu.gs, Triggers.gs, and this
 *      file are ALL pasted into the same Apps Script project (each as
 *      its own file — Apps Script lets you have many files; they share
 *      one global scope, so order of files in the sidebar doesn't matter).
 *   2. Pick "setup" from the function dropdown → Run → approve
 *      permissions (first time: Advanced → Go to project (unsafe) →
 *      Allow).
 *   3. View ▸ Logs shows your FORM LINK — share it with Area Managers.
 *   4. Reopen the Sheet (or refresh the page) to see the "AM Tools" menu.
 *************************************************************************/

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('FORM_ID')) {
    Logger.log('A form already exists. Run resetEverything() first if you really want to rebuild from scratch.');
    return;
  }

  var form = buildForm_();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  buildAllMasterData_(ss);
  buildMasterLogSheet_(ss);
  buildIssueLogSheet_(ss);
  hideFormResponsesSheet_(ss);
  buildDashboardShell_(ss);
  installTriggers_(form);

  props.setProperty('FORM_ID', form.getId());
  refreshDashboard();

  Logger.log('DONE ✅\nShare this checklist link with Area Managers:\n' + form.getPublishedUrl());
  Logger.log('Edit the form here:\n' + form.getEditUrl());
}

function resetEverything() {
  PropertiesService.getDocumentProperties().deleteProperty('FORM_ID');
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  Logger.log('Reset done. Run setup() again. (Existing sheets/form are not deleted — remove them manually first if you want a truly clean rebuild.)');
}


/*************************************************************************
 * SubmitHandler.gs — runs on every Form submission. Parses the response,
 * writes ONE row to Master Log (per visit) and ZERO+ rows to Issue Log
 * (one per failed item). Also owns building those two sheets, since
 * their structure is defined entirely by what gets written here.
 *************************************************************************/

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
      return [ts, visitDate, store, am, f[0], f[1], 'Open', '', '', '', ''];
      // Status, Assigned To, Due Date, Closed Date, Severity — ready for
      // future issue-tracking workflow, not populated yet.
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

// One row per FAILED item — includes future Status/Assigned To/Due Date/
// Closed Date/Severity columns (left blank for now, ready when you want
// an actual issue-resolution workflow).
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
      try { s.hideSheet(); } catch (err) { /* can't hide the only visible sheet yet — ignore */ }
    }
  });
}


/*************************************************************************
 * Triggers.gs — installs the two triggers this system needs: one that
 * fires on every Form submission, one that refreshes the dashboard once
 * a day as a safety net (e.g. to update "days since last visit" even on
 * a day with zero submissions).
 *************************************************************************/

function installTriggers_(form) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onFormSubmitHandler' || fn === 'dailyRefresh') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmitHandler').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('dailyRefresh').timeBased().everyDays(1).atHour(9).create();
}

function dailyRefresh() { refreshDashboard(); }





