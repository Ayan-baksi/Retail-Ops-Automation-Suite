// Config.gs
// VM Store Audit — configuration for the Visual Merchandiser system.
// Same architecture as the AM/SM systems. Checklist items are taken
// directly from VM-Checklist.xlsx with wording unchanged. Each item is
// rated 1-5 (or Not Applicable), rolled up into a per-area adherence %
// and an overall store rating, matching the source workbook's Summary sheet.

var CONFIG = {
  ORG_NAME: 'DeeHub Lifestyle',
  APP_VERSION: 'v1.0',

  FORM_TITLE: 'DeeHub Lifestyle — VM Store Audit',
  FORM_DESC: 'Rate every item 1 (Poor) to 5 (Excellent), or Not Applicable. Every question is mandatory — you cannot proceed until this section is complete.',
  CONFIRMATION_MSG: 'Thank you — this VM audit has been recorded and the dashboard has been updated.',

  STORES: [
    'STORE-001', 'STORE-002', 'STORE-003', 'STORE-004', 'STORE-005',
    'STORE-006', 'STORE-007', 'STORE-008', 'STORE-009', 'STORE-010'
  ],
  LOGO_FILE_ID: '',
  OVERDUE_DAYS_DEFAULT: 30,
  LOW_SCORE_THRESHOLD: 2 // an item scored 1 or 2 is logged as an issue
};

var SHEETS = {
  MASTER: 'Master Log',
  ISSUES: 'Issue Log',
  SCORES: 'Item Score Log',
  DASH: 'Dashboard',
  STORE_MASTER: 'Store_Master',
  EMPLOYEE_MASTER: 'Employee_Master',
  CHECKLIST_MASTER: 'Checklist_Master',
  SETTINGS: 'Settings'
};

var Q = {
  VM: 'VM Name',
  STORE: 'Store',
  SM: 'Store Manager Name',
  DATE: 'Date',
  ACTIVITY: 'Activity Name',
  FINAL_REMARKS: 'Overall Remarks for This Visit'
};

var TZ = Session.getScriptTimeZone();

var FILTER_CELL = 'A5';
var MONTH_FILTER_CELL = 'A8';
var REFRESH_CHECKBOX_CELL = 'R2';

var SCORE_CHOICES = ['1', '2', '3', '4', '5', 'Not Applicable'];

// Matches the source workbook's Summary sheet exactly.
var VM_BANDS = [
  { min: 91, max: 100, label: 'Excellent', color: '#16a34a' },
  { min: 81, max: 90, label: 'Good', color: '#65a30d' },
  { min: 71, max: 80, label: 'Average', color: '#d97706' },
  { min: 61, max: 70, label: 'Poor', color: '#ea580c' },
  { min: 0, max: 60, label: 'Very Poor', color: '#dc2626' }
];

function vmBand_(pct) {
  for (var i = 0; i < VM_BANDS.length; i++) {
    if (pct >= VM_BANDS[i].min) return VM_BANDS[i];
  }
  return VM_BANDS[VM_BANDS.length - 1];
}

// Checklist definition
var SECTIONS = [
  { name: 'Façade', items: [
    'Façade ACP proprely wahsed',
    'Façade- good condition',
    'Façade - Focus lights are lit and in working  condition',
    'Façade visual are in good condition, updated as per current season',
    'Flagpole - Are lit / non lit in working condition (adress the issues to the concern mkt manager)'
  ]},
  { name: 'Overall Store upkeep', items: [
    'All lights are well lit & fused lights identified and planned to be replaced',
    'Is there an unpleasant odour or smell in the store',
    'All Trial rooms are functional and clean',
    'Trial room legal signages are placed as per guidelines',
    'All legal signages displayed on the entry gate are in good condition',
    'Store directory is placed at the designated locations and is lit',
    'Directional signages are placed for easy navigation at regular intervals',
    'All legal / mandatory signages are placed in the store as per guidelines',
    'Visual/Branding cash counter in good condition and is lit',
    'Cash Counter is clean, POS machines and the trays are cleaned. Carry Bags are neatly stacked within the cashtil',
    'floor fixtures right browsing space. Excess/short fixtures notified to RM'
  ]},
  { name: 'Store Entrance', items: [
    'Entrance  mannequins alignment is proper ( grouping)',
    'Store timing and logo is properly placed on entry gate',
    'All the mannequin styling is done as per styling',
    'Current offer  communication is placed properly and in good condition',
    'Focusing of light -Entrance cluster'
  ]},
  { name: 'Merchandise Presentation - Overall', items: [
    ' as per merchandise display guidelines Mens & Womens Denim, Chinos, Infants & Accessories',
    'Basic hygiene maintained - stacking, hanging size wise display',
    'Are the merchandise displayed as per suggestion - season wise',
    ' planogram guidelines and right signage incorporated',
    'yes',
    'Mannequin styling - executed as per the updated styling docket received  and the products needs to be displayed next to the cluster as per space availibility',
    'brick wise communication displayed as per guidelines',
    ' signages as per guidelines and the right sizes used appropriately as per fixtures',
    'Relevance of the visual ( in colum and other area) with the adjoining category',
    'Cut size merchandise to be displayed with approprite Size talkers & last of the best signage',
    'Same merchandise available in hanging display and bottom stacking?',
    'All the necessary floor fixtures  communication maintained as per planogram guidelines'
  ]},
  { name: 'Visual / VM Marketing Communication', items: [
    'Signage holder- All table top, wall signage stands and  are in good condition',
    'A5 & Gondola header Signages are placed as per the guideline & No spelling errors/ wrong font or old format used in any Signage.',
    'Correct A5 signage holders placed appropriately in respective fixtures'
  ]}
];
