/**
 * ============================================================================
 * SETUP.GS
 * Configuration, sheet/column schema, and one-time initialization.
 *
 * ONE-TIME SETUP (run once from the Apps Script editor before deploying):
 *   1. Open this project in script.google.com
 *   2. Select the function `runInitialSetup` in the dropdown, click Run.
 *   3. Grant the permissions it asks for.
 *   4. Check the Execution Log — it will confirm the Spreadsheet ID, create
 *      any missing sheet tabs (with headers) on your existing spreadsheet,
 *      and create the root "Tour Visit" Drive folder, saving its ID.
 *   5. Services > Advanced Google Services > enable "Drive API" (this is
 *      required for PDF generation — see PDF_Engine.gs header comment).
 *   6. Deploy > New deployment > Web app > Execute as: Me, Who has access:
 *      Anyone. Copy the /exec URL into assets/js/config.js on the frontend.
 *
 * SECURITY NOTE: The Spreadsheet ID and Drive folder ID are intentionally
 * NOT hardcoded as literals used at runtime — they live only in this
 * project's Script Properties (Project Settings > Script Properties),
 * which are never sent to the browser and are not part of the public
 * GitHub Pages frontend. `runInitialSetup` is the only place the ID you
 * provide is written down, and only into Script Properties.
 * ============================================================================
 */

// ----------------------------------------------------------------------------
// Paste your existing Spreadsheet ID here ONLY for the initial setup run.
// After runInitialSetup() executes once, CONFIG below reads exclusively from
// PropertiesService — this constant is not referenced anywhere else.
// ----------------------------------------------------------------------------
const SETUP_SPREADSHEET_ID = '1SGqVc3OYt30WPKe5gxLETZTyO0-BGUha7oH8Zf3Bvq0';

const CONFIG = {
  get SPREADSHEET_ID() {
    const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) {
      throw new Error('SPREADSHEET_ID not set. Run runInitialSetup() once from the Apps Script editor.');
    }
    return id;
  },
  get DRIVE_ROOT_FOLDER_ID() {
    const id = PropertiesService.getScriptProperties().getProperty('DRIVE_ROOT_FOLDER_ID');
    if (!id) {
      throw new Error('DRIVE_ROOT_FOLDER_ID not set. Run runInitialSetup() once from the Apps Script editor.');
    }
    return id;
  },
  SESSION_TIMEOUT_SECONDS: 21600, // 6 hours — CacheService's hard maximum
  MAX_LOGIN_ATTEMPTS: 5,
  LOGIN_LOCKOUT_SECONDS: 900, // 15 minutes
  MAX_UPLOAD_BYTES: 10 * 1024 * 1024, // 10 MB, matches spec's file validation requirement
  ALLOWED_SIGNED_CHECKLIST_TYPES: ['application/pdf', 'image/jpeg', 'image/jpg', 'image/png'],
  API_VERSION: '2.0.0',
  QR_ENDPOINT: 'https://quickchart.io/qr' // Google's chart.googleapis.com QR endpoint was retired; this is the current reliable free alternative
};

const SHEETS = {
  BASE: 'Base',
  USERS: 'Users',
  CHECKLIST_MASTER: 'Checklist_Master',
  TOUR_VISIT: 'Tour_Visit',
  CHECKLIST_RESPONSE: 'Checklist_Response',
  AUDIT_LOG: 'Audit_Log'
};

// Named column indices (0-based) — avoids magic numbers throughout the codebase.
const COLS = {
  BASE: { DISTRICT: 0, BUS_STATION: 1 },
  USERS: { EMPLOYEE_ID: 0, NAME: 1, DESIGNATION: 2, ROLE: 3, DISTRICT: 4, PASSWORD: 5, STATUS: 6 },
  CHECKLIST_MASTER: { QUESTION_ID: 0, CATEGORY: 1, QUESTION: 2, DISPLAY_ORDER: 3, ACTIVE: 4 },
  TOUR_VISIT: {
    VISIT_ID: 0, EMPLOYEE_ID: 1, EMPLOYEE_NAME: 2, DESIGNATION: 3, DISTRICT: 4,
    BUS_STATION: 5, VISIT_DATE: 6, VISIT_TIME: 7, LATITUDE: 8, LONGITUDE: 9,
    GPS_ADDRESS: 10, OVERALL_REMARK: 11, GENERATED_PDF_URL: 12,
    SIGNED_CHECKLIST_URL: 13, CREATED_TIMESTAMP: 14
  },
  // REMARK/RATING/DIFFICULTY_OPTIONS were added after TIMESTAMP (not inserted
  // before it) so this schema change never shifts columns for rows written
  // by the old code — existing Checklist_Response data stays valid as-is.
  // See the migrateChecklistResponseSchema() one-time helper below.
  CHECKLIST_RESPONSE: {
    VISIT_ID: 0, QUESTION_ID: 1, QUESTION: 2, ANSWER: 3, TIMESTAMP: 4,
    REMARK: 5, RATING: 6, DIFFICULTY_OPTIONS: 7
  },
  AUDIT_LOG: { TIMESTAMP: 0, EMPLOYEE_ID: 1, ACTION: 2, DEVICE: 3, BROWSER: 4 }
};

const SHEET_HEADERS = {
  Base: ['District', 'Bus Station'],
  Users: ['Employee ID', 'Name', 'Designation', 'Role', 'District', 'Password', 'Status'],
  Checklist_Master: ['Question ID', 'Category', 'Question', 'Display Order', 'Active'],
  Tour_Visit: [
    'Visit ID', 'Employee ID', 'Employee Name', 'Designation', 'District', 'Bus Station',
    'Visit Date', 'Visit Time', 'Latitude', 'Longitude', 'GPS Address', 'Overall Remark',
    'Generated PDF URL', 'Signed Checklist URL', 'Created Timestamp'
  ],
  Checklist_Response: ['Visit ID', 'Question ID', 'Question', 'Answer', 'Timestamp', 'Remark', 'Rating', 'Difficulty Options'],
  Audit_Log: ['Timestamp', 'Employee ID', 'Action', 'Device', 'Browser']
};

const ROLES = {
  ADMIN: 'Admin',
  PM: 'PM',
  DM: 'DM',
  ADM: 'ADM'
};

/**
 * Run this ONCE from the Apps Script editor. Idempotent — safe to re-run;
 * it will not overwrite existing sheet data, only add missing tabs/headers.
 */
function runInitialSetup() {
  const props = PropertiesService.getScriptProperties();

  // 1. Save Spreadsheet ID
  props.setProperty('SPREADSHEET_ID', SETUP_SPREADSHEET_ID);
  const ss = SpreadsheetApp.openById(SETUP_SPREADSHEET_ID);
  Logger.log('Connected to spreadsheet: ' + ss.getName());

  // 2. Ensure every required tab exists with correct headers (non-destructive)
  Object.keys(SHEET_HEADERS).forEach(function (sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(SHEET_HEADERS[sheetName]);
      sheet.setFrozenRows(1);
      Logger.log('Created missing sheet: ' + sheetName);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(SHEET_HEADERS[sheetName]);
      sheet.setFrozenRows(1);
      Logger.log('Added headers to empty sheet: ' + sheetName);
    } else {
      Logger.log('Sheet already present: ' + sheetName + ' (' + (sheet.getLastRow() - 1) + ' data rows)');
    }
  });

  // 3. Create (or find) the root "Tour Visit" Drive folder
  let rootFolderId = props.getProperty('DRIVE_ROOT_FOLDER_ID');
  if (!rootFolderId) {
    const existing = DriveApp.getFoldersByName('Tour Visit');
    const rootFolder = existing.hasNext() ? existing.next() : DriveApp.createFolder('Tour Visit');
    rootFolderId = rootFolder.getId();
    props.setProperty('DRIVE_ROOT_FOLDER_ID', rootFolderId);
    Logger.log('Root Drive folder ready: ' + rootFolder.getUrl());
  } else {
    Logger.log('Root Drive folder already configured: ' + rootFolderId);
  }

  Logger.log('=== Setup complete ===');
  Logger.log('Next steps: enable the Drive API advanced service (see PDF_Engine.gs header), then deploy as a Web App.');
}

/**
 * Run this ONCE if Checklist_Response already existed before the Rating /
 * Difficulty-Options feature was added. It only appends the three new
 * headers (Remark, Rating, Difficulty Options) in columns F/G/H if they
 * are missing — it never touches existing rows or existing columns, so
 * historical Visit_ID/Question/Answer/Timestamp data is untouched.
 * Safe to re-run; it's a no-op if the headers are already present.
 */
function migrateChecklistResponseSchema() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CHECKLIST_RESPONSE);
  if (!sheet) {
    Logger.log('Checklist_Response sheet not found — run runInitialSetup() first.');
    return;
  }
  const headerRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 8)).getValues()[0];
  const expected = SHEET_HEADERS.Checklist_Response;
  let changed = false;
  for (let col = 0; col < expected.length; col++) {
    if (String(headerRow[col] || '').trim() !== expected[col]) {
      sheet.getRange(1, col + 1).setValue(expected[col]);
      changed = true;
    }
  }
  Logger.log(changed ? 'Checklist_Response headers updated to include Remark/Rating/Difficulty Options.' : 'Checklist_Response headers already up to date.');
}

/**
 * Run this ONCE from the Apps Script editor to replace all rows in
 * Checklist_Master with the current 16-question set. This is a full
 * REPLACE — it clears every existing data row (row 2 downward) before
 * writing the new ones, so any manual edits made directly in the sheet
 * since the last time this ran will be overwritten. Headers are left
 * untouched. Safe to re-run; re-running just rewrites the same 16 rows.
 */
function seedChecklistMaster() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.CHECKLIST_MASTER);
  if (!sheet) {
    Logger.log('Checklist_Master sheet not found — run runInitialSetup() first.');
    return;
  }

  const questions = [
    ['Q001', 'कर्मचारी व्यवस्थापन', 'सर्व कर्मचाऱ्यांना भेटून त्यांच्या अडचणी काही असल्यास  त्यांना उपाय सांगावे. अडचणी असल्यास खालील पर्याय टिक करावा.', 1, true],
    ['Q002', 'कर्मचारी व्यवस्थापन', 'सर्व कर्मचारी कंपनीने दिलेले युनिफॉर्म, शूज व आवश्यक PPE वापरत आहेत ?रेटिंग (५ पैकी): ⭐ १ / २ / ३ / ४ / ५', 2, true],
    ['Q003', 'कर्मचारी व्यवस्थापन', 'आवश्यक साफसफाई साहित्य, रसायने व उपकरणे उपलब्ध आहेत का, ते तपासावे. रेटिंग (५ पैकी): ⭐ १ / २ / ३ / ४ / ५', 3, true],
    ['Q004', 'कर्मचारी व्यवस्थापन', 'सूपरवायजरचे नियोजन कंपनीच्या धोरणाप्रमाणे आहे का ?( हो / नाही )', 4, true],
    ['Q005', 'कर्मचारी व्यवस्थापन', 'कर्मचाऱ्यानी  भरतीसाठी कुठल्याही प्रकारची देवाण घेवाण केल्याचे आढळून आले आहे का ? ( हो / नाही )', 5, true],
    ['Q006', 'साहित्य व उपकरणे', 'पुरुष व महिला स्वच्छतागृहांची स्वच्छता व स्थिति तपासावी. रेटिंग (५ पैकी): ⭐ १ / २ / ३ / ४ / ५', 6, true],
    ['Q007', 'साहित्य व उपकरणे', 'चालक–वाहक विश्रांतीगृहाची स्वच्छता व सुविधा तपासाव्यात.रेटिंग (५ पैकी): ⭐ १ / २ / ३ / ४ / ५', 7, true],
    ['Q008', 'स्वच्छता स्थिती', 'बस वॉशिंगची गुणवत्ता, व मशीनची कार्यस्थिती तपासावी.रेटिंग (५ पैकी): ⭐ १ / २ / ३ / ४ / ५', 8, true],
    ['Q009', 'स्वच्छता स्थिती', 'वॉटर प्युरिफायर, गिझर, जेट स्प्रे व इतर उपकरणे कार्यरत आहेत का, ते तपासावे.रेटिंग (५ पैकी): ⭐ १ / २ / ३ / ४ / ५', 9, true],
    ['Q010', 'स्वच्छता स्थिती', 'डस्टबिन स्वच्छ असून वेळेवर रिकामे केले जात आहेत का, ते तपासावे. रेटिंग (५ पैकी): ⭐ १ / २ / ३ / ४ / ५', 10, true],
    ['Q011', 'स्वच्छता स्थिती', 'बस स्थानकातील ओपन एरिया, प्लॅटफॉर्म, वेटिंग एरिया, तिकीट काऊंटर परिसर व इतर सर्व भाग स्वच्छ आहेत का, ते तपासावे.', 11, true],
    ['Q012', 'स्वच्छता स्थिती', 'ड्रेनेज, पाणीपुरवठा, विद्युत किंवा इतर कोणत्याही तांत्रिक समस्या असल्यास त्याची नोंद घ्यावी.', 12, true],
    ['Q013', 'पायाभूत सुविधा', 'कोणतीही तक्रार, त्रुटी किंवा विशेष निरीक्षण असल्यास त्याची नोंद करून त्वरित वरिष्ठांना कळवावे.', 13, true],
    ['Q014', 'तक्रार व अहवाल', 'डेपो मॅनेजर यांची भेट घेऊन बस स्थानकातील स्वच्छतेबाबत अभिप्राय (Feedback) घ्यावा.', 14, true],
    ['Q015', 'डेपो व्यवस्थापन समन्वय', 'डेपो मॅनेजर यांनी दिलेल्या सूचना व निरीक्षणांची नोंद करून त्यावर सकारात्मक व वेळेत कार्यवाही करावी. तसेच, केलेल्या कार्यवाहीची माहिती वेळोवेळी डेपो मॅनेजर यांना देण्यात यावी.', 15, true],
    ['Q016', 'डेपो व्यवस्थापन समन्वय', 'सुपरवायझर दररोजची चेकलिस्ट वर नियमितपणे डेपो मॅनेजर किंवा इतर अधिकाऱ्याची सही-शिक्का घेत आहेत का?', 17, true]
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }
    sheet.getRange(2, 1, questions.length, 5).setValues(questions);
  } finally {
    lock.releaseLock();
  }

  // Master data is cached (see handleGetChecklist) — clear it so the app
  // picks up these rows immediately instead of waiting out the old cache.
  CacheService.getScriptCache().remove('data_checklist');

  Logger.log('Checklist_Master seeded with ' + questions.length + ' questions.');
}
