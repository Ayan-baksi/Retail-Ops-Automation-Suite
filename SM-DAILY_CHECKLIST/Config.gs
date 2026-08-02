// Config.gs
// SM Daily Checklist — configuration for the Store Manager Opening/Closing
// system. Same architecture as the AM system. Checklist items are taken
// directly from UPDDATED_CHECKLIST_FOR_SM.xlsx with wording unchanged.

var CONFIG = {
  ORG_NAME: 'DeeHub Lifestyle',
  APP_VERSION: 'v1.0',

  FORM_TITLE: 'DeeHub Lifestyle — Store Manager Daily Checklist',
  FORM_DESC: 'Complete this once per Opening or Closing shift. Every checklist question is mandatory (Yes/No/Not Applicable) — you cannot proceed to the next section until all questions and any required photos on this page are complete.',
  CONFIRMATION_MSG: 'Thank you — your checklist has been recorded and the SM dashboard has been updated.',

  STORES: [
    'STORE-001', 'STORE-002', 'STORE-003', 'STORE-004', 'STORE-005',
    'STORE-006', 'STORE-007', 'STORE-008', 'STORE-009', 'STORE-010'
  ],
  LOGO_FILE_ID: '',
  OVERDUE_DAYS_DEFAULT: 2, // Opening/Closing is a daily cadence, so "overdue" is much shorter than a monthly AM visit
  TARGET_SCORE_DEFAULT: 90
};

var SHEETS = {
  MASTER: 'Master Log',
  ISSUES: 'Issue Log',
  PHOTOS: 'Photo Log',
  DASH: 'Dashboard',
  STORE_MASTER: 'Store_Master',
  EMPLOYEE_MASTER: 'Employee_Master',
  CHECKLIST_MASTER: 'Checklist_Master',
  SETTINGS: 'Settings'
};

var Q = {
  SM: 'Store Manager Name',
  STORE: 'Store',
  DATE: 'Date',
  AUDIT_TYPE: 'Audit Type (Opening / Closing)',
  FINAL_REMARKS: 'Overall Remarks for This Audit'
};

var TZ = Session.getScriptTimeZone();

var FILTER_CELL = 'A5';
var MONTH_FILTER_CELL = 'A8';
var TYPE_FILTER_CELL = 'A11';  // Opening / Closing / Both filter
var DAY_FILTER_CELL = 'A14';   // specific day, or "Today"
var REFRESH_CHECKBOX_CELL = 'R2';

// Checklist definition.
// Each item: { text, image (bool), maxFiles (optional), imageNote (optional) }
// Each section: { name, note (optional), items }
var SECTIONS_OPENING = [
  { name: 'Store Opening Activities', items: [
    { text: 'Coordinating with all the respective designated employees for smooth Opening at 9 AM', image: false },
    { text: 'Ensure the shutter, tijori and exit gate key rules are followed before opening the store', image: true },
    { text: 'Is the store façade and entrance area is cleaned and free from clutter', image: true },
    { text: 'Cash counter is cleaned, neat and free of merchandise', image: true },
    { text: 'All the security antennas, deactivation pads hard tag and soft tag detachers operational and in working condition', image: false },
    { text: 'Checking the store hygiene, cleaning and overall ambience of floor, trial rooms, washroom, straicase, parking area, baggage and façade', image: true },
    { text: 'Ensure that the first customer entered to the store is greeted by the Duty Manager or Store Manager', image: true },
    { text: 'Is the store fully ready, including functioning lights, AC, CCTV systems and music PLAYING?', image: false },
    { text: 'Floor walk along with basic Department wise display hyegine', image: true, maxFiles: 4, imageNote: "Upload photos for Men's, Ladies, Kids, and Accessories sections" },
    { text: 'Approving and reporting about Exchanges and WBC (Without barcode items)', image: false }
  ]},
  { name: 'Business Review', items: [
    { text: 'TGT vs ACHV review and revival plan with DM & CSA', image: false },
    { text: 'KPI Review with data', image: false },
    { text: 'Sales Analysis for Top 20 articles (from each division/section) & Inventory status checks for necessary actions', image: false },
    { text: 'Sales and Operations Review (Till Article Level)', image: false },
    { text: 'Analysis of Slow/Non moving products and to take feasible actions to avoid stock junk in the store', image: false },
    { text: 'Zero sale article review share the actionable in the given format', image: false }
  ]},
  { name: 'HR/Admin/Maintainance Activities',
    note: 'Pls ensure all the staff lunch breck to be completed between 1 PM to 4 PM (Including DM,SM & CSA)',
    items: [
    { text: 'Are all the staff members are having proper grooming, wearing ID card and uniform , Staff briefing to be done 11.30 Am to 12.00 AM', image: true },
    { text: 'Adjusting MPPs amongst department/division for hassle free store operations', image: false },
    { text: 'Checking the Man Power Planning and Rosters for staffs and Cashiers', image: false },
    { text: 'Recruitment of staffs (As and When needed)', image: false },
    { text: 'Maintaining all the registers of the store and undersigning by SM/ASM of all the registers', image: false },
    { text: 'Ensuring Store Safety/DG/Fire/Civil work/Parking/external hazards', image: false },
    { text: 'Responsibility for the fooding/lodging for all the staffs at the canteen are arranged and take care', image: false },
    { text: 'Check all fire and safety norms are followed such as extinguisher, CO2 and ABC are working and not expired', image: false },
    { text: 'Daily electric meter tracking & maintaining peak & non-peak hours AC & Lights', image: false }
  ]},
  { name: 'Cash Activities',
    note: 'Ensure all cash tills operational at the peak hours',
    items: [
    { text: 'Controlling and reporting of Cash Refunds, manual Billing, Bill Reprint are registered and taken prior approval taken from the management', image: false },
    { text: 'Ensure to follow the cash till management SOP', image: false },
    { text: 'Ensure to follow the exchange SOP with proper entry in exchanfe register', image: false },
    { text: 'Maintaining cash deposit, Handling discrepencies (cash short/excess), controlling expenses', image: false },
    { text: 'Observing and controlling of Inventory (Inwarding and Outwarding) and requirement of stock on daily basis', image: false }
  ]},
  { name: 'Visual', items: [
    { text: 'Are all the VM guidelines adhered and reviewing the standard', image: true },
    { text: 'Checking fixtures display quantity and taking action for replenishment', image: false },
    { text: 'Ongoing promotion/offers communication to be visble at the store and staffs are aware of the offer', image: false }
  ]},
  { name: 'Process Adherance', items: [
    { text: 'In case of stock receiving at night, proper planning and security needs to be done', image: false },
    { text: 'Focus on fast delivery of altered products to the customers', image: false },
    { text: 'Scrap process adherence', image: false },
    { text: 'Is the AC temperature in the store maintained with in 24 degrees ?', image: false },
    { text: 'Ensuring a designated cash counter is dedicated for staff billing and maintained register', image: false },
    { text: 'Is there hourly customer foot fall register maintained and checked by SM ?', image: false },
    { text: 'Ensuring Replenishment in the night, clearing the floor from cartons/ bags and scraps are done', image: false }
  ]}
];

var SECTIONS_CLOSING = [
  { name: 'Store Closing Activities', items: [
    { text: 'Half sutter down for customer closing as per SOP (9.30 PM)', image: false },
    { text: 'Politly annouce customers about store closing & final billing', image: false },
    { text: 'After final billing match all the MOP (Cash,Card,phpe)', image: false },
    { text: 'Ensure all the left over salable merchandise of cash till & trial room to be cleard & put-away at the concern department', image: true },
    { text: 'Ensure visuals, department & section display properly arranged & 100 % ready for next day', image: true },
    { text: 'After final settlement send all account related data to the concern department', image: false },
    { text: 'Ensure lock & seal cash drawer & safe', image: true },
    { text: 'Final floor walk done by closing SM/ASM/FM', image: false },
    { text: 'Switch off all the AC,Lights & POS counter', image: false },
    { text: 'Properly lock & seal all the shutter lock', image: true },
    { text: 'After the store closing click the closing selfie from outside of the store', image: true }
  ]}
];
