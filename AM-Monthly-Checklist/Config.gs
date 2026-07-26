
var CONFIG = {
  ORG_NAME: 'DeeHub Lifestyle',
  APP_VERSION: 'v1.0',

  FORM_TITLE: 'DeeHub Lifestyle — Area Manager Store Audit',
  FORM_DESC: 'Welcome. This checklist takes about 10–15 minutes per store visit. ' +
    'Your responses are recorded instantly and feed the live performance dashboard automatically — ' +
    'no separate reporting needed.',
  CONFIRMATION_MSG: 'Thank you — your audit has been recorded and the DeeHub dashboard has been updated.',

  
  LOGO_FILE_ID: '',

  // >>> Edit this list if your store footprint changes <<<
  STORES: [
    'DB-PWS', 'DB-BBU', 'DB-SMT', 'DB-KTR', 'DB-JMU',
    'DB-DGR', 'DB-SLP', 'DB-JRK', 'DB-PTM', 'DB-MRS'
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
