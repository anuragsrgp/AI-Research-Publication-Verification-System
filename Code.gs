/***************************************************************
 * UNIVERSITY RESEARCH PUBLICATION VERIFICATION SYSTEM
 * VERSION: 4.0
 *
 * ARCHITECTURE
 * -------------------------------------------------------------
 * FORM_REGISTRY
 *      ↓
 * Google Forms / Response Sheets
 *      ↓
 * MASTER_DATA
 *      ↓
 * VERIFICATION
 *      ↓
 * Crossref API
 *
 * IMPORTANT:
 * MASTER_DATA contains USER SUBMITTED DATA.
 * API data NEVER replaces MASTER_DATA user data.
 ***************************************************************/


/* ============================================================
   1. GLOBAL CONFIGURATION
   ============================================================ */

const CONFIG = {

  SHEETS: {
    REGISTRY: 'FORM_REGISTRY',
    MASTER: 'MASTER_DATA',
    VERIFICATION: 'VERIFICATION',
    LOG: 'SYSTEM_LOG'
  },

  STATUS: {
    VERIFIED: 'VERIFIED',
    NOT_VERIFIED: 'NOT VERIFIED',
    PENDING: 'PENDING',
    MANUAL_REVIEW: 'MANUAL REVIEW',
    API_ERROR: 'API ERROR'
  },

  CROSSREF: {
    BASE_URL: 'https://api.crossref.org/works',
    MAILTO: '',
    ROWS: 3,
    CACHE_SECONDS: 21600
  },

  AI: {
    ENABLED: true,
    PROVIDER: 'GROQ',
    MODEL: 'openai/gpt-oss-20b',
    FREE_ONLY: true,
    MAX_RETRIES: 1,
    RETRY_DELAYS_MS: [0],
    API_KEY_PROPERTY: 'GROQ_API_KEY',
    BASE_URL: 'https://api.groq.com/openai/v1/chat/completions',
    CACHE_SECONDS: 21600,
    MAX_COMPLETION_TOKENS: 600,
    // AI is used only for ambiguous Crossref results.
    RUN_FOR_STATUSES: ['MANUAL REVIEW', 'NOT VERIFIED', 'PENDING']
  },

  MATCH: {
    VERIFIED: 90,
    MANUAL_REVIEW: 70
  },

  SYNC: {
    BATCH_SIZE: 200
  },

  PERFORMANCE: {
    VERIFY_CACHE_HOURS: 24,
    SKIP_UNCHANGED_AUTOMATICALLY: true
  },

  MASTER_METADATA_COLUMNS: [
    'System ID',
    'Form Name',
    'Form ID',
    'Response ID',
    'Response Timestamp',
    'Last Synced'
  ],

  VERIFICATION_COLUMNS: [
    'System ID',
    'Verification Status',
    'API Status',
    'API DOI',
    'API Title',
    'API Authors',
    'API Journal / Conference / Book',
    'API Publisher',
    'Match Score %',
    'Verification Message',
    'Last Verified',
    'Verification Source',
    'Input DOI',
    'Input Title',
    'Input Authors',
    'Input Category',
    'Source Form',
    'Source DOI',
    'API URL',
    'Source URL',
    'AI Status',
    'AI Match Score %',
    'AI Confidence',
    'AI Decision',
    'AI Explanation',
    'AI Recommendation',
    'AI Title Similarity %',
    'AI Author Similarity %',
    'AI Checked At'
  ],

  STATUS_LIST: [
    'VERIFIED',
    'NOT VERIFIED',
    'PENDING',
    'MANUAL REVIEW',
    'API ERROR'
  ]
};


/* ============================================================
   2. MENU
   ============================================================ */

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('Research System')

    .addItem('Setup / Repair System', 'setupSystem')

    .addSeparator()

    .addItem('Add Form', 'addForm')

    .addItem('Sync All Forms', 'syncAllForms')

    .addItem('Sync Selected Form', 'syncSelectedForm')

    .addSeparator()

    .addItem('Verify All Publications', 'verifyAllPublications')
    .addItem('Force Verify All Publications', 'forceVerifyAllPublications')

    .addItem('Verify Selected Row', 'verifySelectedRow')

    .addItem('Apply Verification Dropdown', 'applyVerificationDropdown')

    .addItem('Configure Groq API Key', 'configureAIKey')
    .addItem('Test Groq AI Integration', 'testAIIntegration')

    .addSeparator()

    .addItem('Run Full System Repair', 'runFullSystemRepair')

    .addItem('Run Diagnostics', 'runDiagnostics')

    .addItem('Clear Old Logs', 'clearOldLogs')

    .addSeparator()

    .addItem('Install Automatic Triggers', 'installAutomaticTriggers')

    .addToUi();
}


/* ============================================================
   3. BASIC UTILITIES
   ============================================================ */

function getSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}


function getSheet_(name) {

  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  return sheet;
}


function removeSheet1_() {
  try {
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName('Sheet1');
    if (!sheet) return;

    const required = Object.keys(CONFIG.SHEETS).map(function(k) {
      return CONFIG.SHEETS[k];
    });

    if (required.indexOf(sheet.getName()) !== -1) return;
    if (ss.getSheets().length <= 1) return;

    ss.deleteSheet(sheet);
    info_('removeSheet1_', 'Removed default Sheet1', 'Sheet1 deleted');
  } catch (e) {
    warning_('removeSheet1_', 'Could not remove Sheet1', e.message);
  }
}


function now_() {
  return new Date();
}


function safeString_(value) {

  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}


function normalize_(value) {

  return safeString_(value)
    .toLowerCase()
    .replace(/[\r\n]+/g, ' ')
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function normalizeDOI_(value) {

  let doi = safeString_(value);

  if (!doi) {
    return '';
  }

  doi = doi
    .replace(/^https?:\/\/(dx\.)?doi\.org\//i, '')
    .replace(/^doi:\s*/i, '')
    .trim();

  return doi;
}


function generateSystemId_() {

  return Utilities.getUuid();
}


function escapeRegex_(text) {

  return safeString_(text)
    .replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


/* ============================================================
   4. LOGGING
   ============================================================ */

function log_(level, functionName, message, details) {

  try {

    const sheet = getSheet_(CONFIG.SHEETS.LOG);

    if (sheet.getLastRow() === 0) {

      sheet.getRange(1, 1, 1, 7).setValues([[
        'Time',
        'Level',
        'Function',
        'Message',
        'Details',
        'Timestamp',
        'Context'
      ]]);

      formatHeader_(sheet, 7);
    }

    sheet.appendRow([
      new Date(),
      level,
      functionName,
      message,
      details || '',
      new Date(),
      'Research System'
    ]);

  } catch (error) {

    // Logging must NEVER break main application.
    console.log(
      '[LOG ERROR] ' +
      functionName +
      ' : ' +
      error.message
    );
  }
}


function info_(fn, msg, details) {
  log_('INFO', fn, msg, details);
}


function warning_(fn, msg, details) {
  log_('WARNING', fn, msg, details);
}


function error_(fn, msg, details) {
  log_('ERROR', fn, msg, details);
}


/* ============================================================
   5. HEADER / SHEET MANAGEMENT
   ============================================================ */

function formatHeader_(sheet, columnCount) {

  if (columnCount <= 0) {
    return;
  }

  const range = sheet.getRange(1, 1, 1, columnCount);

  range
    .setFontWeight('bold')
    .setBackground('#17365D')
    .setFontColor('#FFFFFF')
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(1);
}


function ensureHeaders_(sheet, requiredHeaders) {

  const lastColumn = Math.max(sheet.getLastColumn(), 1);

  let currentHeaders = [];

  if (sheet.getLastRow() >= 1) {

    currentHeaders = sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0]
      .map(safeString_);
  }

  if (
    currentHeaders.length === 1 &&
    currentHeaders[0] === ''
  ) {
    currentHeaders = [];
  }

  let changed = false;

  requiredHeaders.forEach(function(header) {

    if (
      currentHeaders
        .map(normalize_)
        .indexOf(normalize_(header)) === -1
    ) {

      currentHeaders.push(header);
      changed = true;
    }

  });

  if (changed || sheet.getLastRow() === 0) {

    sheet
      .getRange(
        1,
        1,
        1,
        currentHeaders.length
      )
      .setValues([currentHeaders]);

    formatHeader_(sheet, currentHeaders.length);
  }

  return currentHeaders;
}


function getHeaders_(sheet) {

  if (sheet.getLastColumn() === 0) {
    return [];
  }

  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(safeString_);
}


function getColumnIndex_(headers, possibleNames) {

  const normalizedHeaders =
    headers.map(normalize_);

  for (let i = 0; i < possibleNames.length; i++) {

    const index =
      normalizedHeaders.indexOf(
        normalize_(possibleNames[i])
      );

    if (index !== -1) {
      return index + 1;
    }
  }

  return 0;
}


function ensureCoreSheets_() {

  const registry = getSheet_(CONFIG.SHEETS.REGISTRY);
  const master = getSheet_(CONFIG.SHEETS.MASTER);
  const verification = getSheet_(CONFIG.SHEETS.VERIFICATION);
  const log = getSheet_(CONFIG.SHEETS.LOG);

  ensureHeaders_(registry, [
    'Form Name',
    'Form ID',
    'Form URL',
    'Response Spreadsheet ID',
    'Response Sheet',
    'Status',
    'Last Sync',
    'Records Synced',
    'Last Error',
    'Created At',
    'Notes'
  ]);

  ensureHeaders_(
    master,
    CONFIG.MASTER_METADATA_COLUMNS
  );

  ensureHeaders_(
    verification,
    CONFIG.VERIFICATION_COLUMNS
  );

  if (log.getLastRow() === 0) {

    ensureHeaders_(log, [
      'Time',
      'Level',
      'Function',
      'Message',
      'Details',
      'Timestamp',
      'Context'
    ]);
  }

  return true;
}


/* ============================================================
   6. SYSTEM SETUP
   ============================================================ */

function setupSystem() {

  try {

    ensureCoreSheets_();
    removeSheet1_();

    autoFormatSheets_();

    installAutomaticTriggers();

    info_(
      'setupSystem',
      'System setup / repair completed',
      'Version 4.0'
    );

    SpreadsheetApp.getUi().alert(
      'Research System',
      'System setup / repair completed successfully.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {

    error_(
      'setupSystem',
      'System setup failed',
      error.stack || error.message
    );

    SpreadsheetApp.getUi().alert(
      'Setup Error',
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/* ============================================================
   7. ADD FORM
   ============================================================ */

/*
 * IMPORTANT:
 *
 * Form title is NOT mandatory.
 *
 * If FormApp.getTitle() fails:
 * 1. Try Drive file name.
 * 2. If still unavailable, use "Form <ID>".
 *
 * Therefore:
 * TITLE FAILURE MUST NOT BLOCK FORM REGISTRATION.
 */

function addForm() {

  const ui = SpreadsheetApp.getUi();

  try {

    ensureCoreSheets_();

    const response = ui.prompt(
      'Add Google Form',
      'Paste Google Form URL or Form ID:',
      ui.ButtonSet.OK_CANCEL
    );

    if (
      response.getSelectedButton() !==
      ui.Button.OK
    ) {
      return;
    }

    const input = safeString_(
      response.getResponseText()
    );

    if (!input) {

      ui.alert(
        'Add Form',
        'Form URL / ID cannot be blank.',
        ui.ButtonSet.OK
      );

      return;
    }

    const formId = extractFormId_(input);

    if (!formId) {

      ui.alert(
        'Add Form',
        'Invalid Google Form URL / ID.',
        ui.ButtonSet.OK
      );

      return;
    }

    const result = registerFormById_(formId);

    if (result.success) {

      ui.alert(
        'Form Added',
        'Form registered successfully.\n\n' +
        'Form Name: ' +
        result.formName +
        '\n\nForm ID: ' +
        formId,
        ui.ButtonSet.OK
      );

    } else {

      ui.alert(
        'Form Registration',
        result.message,
        ui.ButtonSet.OK
      );
    }

  } catch (error) {

    error_(
      'addForm',
      'Unexpected add form error',
      error.stack || error.message
    );

    ui.alert(
      'Add Form Error',
      'The application handled the error safely.\n\n' +
      error.message,
      ui.ButtonSet.OK
    );
  }
}


function extractFormId_(input) {

  const text = safeString_(input);

  if (!text) {
    return '';
  }

  /*
   * Standard:
   * https://docs.google.com/forms/d/FORM_ID/edit
   */

  let match =
    text.match(
      /\/forms\/d\/([a-zA-Z0-9_-]+)/i
    );

  if (match && match[1]) {
    return match[1];
  }

  /*
   * Alternate forms URL
   */

  match =
    text.match(
      /\/forms\/([a-zA-Z0-9_-]+)/i
    );

  if (match && match[1]) {
    return match[1];
  }

  /*
   * If user pasted only ID
   */

  if (
    /^[a-zA-Z0-9_-]{20,}$/.test(text)
  ) {
    return text;
  }

  return '';
}


function registerFormById_(formId) {

  const functionName = 'registerFormById_';

  try {

    const registry =
      getSheet_(CONFIG.SHEETS.REGISTRY);

    const headers = getHeaders_(registry);

    /*
     * Check existing form.
     */

    const formIdColumn =
      getColumnIndex_(
        headers,
        ['Form ID']
      );

    if (formIdColumn > 0) {

      const lastRow =
        registry.getLastRow();

      if (lastRow >= 2) {

        const ids =
          registry
            .getRange(
              2,
              formIdColumn,
              lastRow - 1,
              1
            )
            .getValues()
            .flat()
            .map(safeString_);

        if (ids.indexOf(formId) !== -1) {

          warning_(
            functionName,
            'Form already registered',
            formId
          );

          return {
            success: true,
            alreadyExists: true,
            formName: getFormDisplayName_(formId)
          };
        }
      }
    }

    /*
     * Get title safely.
     */

    const formName =
      getFormDisplayName_(formId);

    /*
     * Get response destination safely.
     */

    let responseSpreadsheetId = '';
    let responseSheetName = '';

    try {

      const form =
        FormApp.openById(formId);

      try {
        responseSpreadsheetId =
          form.getDestinationId() || '';
      } catch (e) {
        warning_(
          functionName,
          'Could not read response destination',
          e.message
        );
      }

    } catch (formError) {

      /*
       * IMPORTANT:
       * Registration does NOT fail here.
       */

      warning_(
        functionName,
        'FormApp access unavailable; continuing',
        formError.message
      );
    }

    /*
     * If destination not available,
     * user can provide response sheet manually later.
     */

    const formUrl =
      'https://docs.google.com/forms/d/' +
      formId +
      '/edit';

    registry.appendRow([
      formName,
      formId,
      formUrl,
      responseSpreadsheetId,
      responseSheetName,
      'ACTIVE',
      '',
      0,
      '',
      new Date(),
      'Registered by Research System'
    ]);

    info_(
      functionName,
      'Form registered',
      formId + ' | ' + formName
    );

    return {
      success: true,
      formName: formName,
      formId: formId
    };

  } catch (error) {

    error_(
      functionName,
      'Form registration failed',
      error.stack || error.message
    );

    return {
      success: false,
      message:
        'Form registration could not be completed.\n\n' +
        error.message
    };
  }
}


/* ============================================================
   8. SAFE FORM TITLE
   ============================================================ */

function getFormDisplayName_(formId) {

  /*
   * Attempt 1: FormApp
   */

  try {

    const form =
      FormApp.openById(formId);

    const title =
      safeString_(form.getTitle());

    if (title) {
      return title;
    }

  } catch (error) {

    warning_(
      'getFormDisplayName_',
      'FormApp title unavailable',
      error.message
    );
  }


  /*
   * Attempt 2: DriveApp
   */

  try {

    const file =
      DriveApp.getFileById(formId);

    const fileName =
      safeString_(file.getName());

    if (fileName) {
      return fileName;
    }

  } catch (error) {

    warning_(
      'getFormDisplayName_',
      'Drive title unavailable',
      error.message
    );
  }


  /*
   * Attempt 3: Safe fallback
   */

  return 'Form ' + formId;
}


/* ============================================================
   9. SYNC ALL FORMS
   ============================================================ */

function syncAllForms() {

  ensureCoreSheets_();

  const registry =
    getSheet_(CONFIG.SHEETS.REGISTRY);

  const data =
    registry.getDataRange().getValues();

  if (data.length < 2) {

    SpreadsheetApp.getUi().alert(
      'Sync',
      'No forms found in FORM_REGISTRY.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return;
  }

  const headers = data[0];

  const formIdCol =
    getColumnIndex_(headers, ['Form ID']);

  const statusCol =
    getColumnIndex_(headers, ['Status']);

  let total = 0;

  for (let r = 1; r < data.length; r++) {

    const row = data[r];

    const formId =
      formIdCol
        ? safeString_(row[formIdCol - 1])
        : '';

    if (!formId) {
      continue;
    }

    const status =
      statusCol
        ? safeString_(row[statusCol - 1])
        : '';

    if (
      status &&
      status.toUpperCase() === 'INACTIVE'
    ) {
      continue;
    }

    try {

      const result =
        syncOneForm_(r + 1);

      total += result.records || 0;

    } catch (error) {

      error_(
        'syncAllForms',
        'One form failed; continuing with next form',
        formId + ' | ' + error.message
      );
    }
  }

  SpreadsheetApp.getUi().alert(
    'Sync Complete',
    'Forms processed successfully.\n\n' +
    'Records synchronized: ' +
    total,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/* ============================================================
   10. SYNC SELECTED FORM
   ============================================================ */

function syncSelectedForm() {

  try {

    const sheet =
      getSheet_(CONFIG.SHEETS.REGISTRY);

    const row =
      sheet.getActiveRange().getRow();

    if (row < 2) {

      SpreadsheetApp.getUi().alert(
        'Select a form row first.'
      );

      return;
    }

    const result =
      syncOneForm_(row);

    SpreadsheetApp.getUi().alert(
      'Sync Complete',
      'Records synchronized: ' +
      result.records,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {

    error_(
      'syncSelectedForm',
      'Selected form sync failed',
      error.stack || error.message
    );

    SpreadsheetApp.getUi().alert(
      'Sync Error',
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/* ============================================================
   11. SYNC ONE FORM
   ============================================================ */

function syncOneForm_(registryRow) {

  const functionName = 'syncOneForm_';

  const registry =
    getSheet_(CONFIG.SHEETS.REGISTRY);

  const registryHeaders =
    getHeaders_(registry);

  const rowValues =
    registry
      .getRange(
        registryRow,
        1,
        1,
        registry.getLastColumn()
      )
      .getValues()[0];


  const formId =
    getRegistryValue_(
      registryHeaders,
      rowValues,
      ['Form ID']
    );

  if (!formId) {

    throw new Error(
      'Form ID is missing in FORM_REGISTRY row ' +
      registryRow
    );
  }


  let formName =
    getRegistryValue_(
      registryHeaders,
      rowValues,
      ['Form Name']
    );

  if (!formName) {
    formName =
      getFormDisplayName_(formId);
  }


  let responseSpreadsheetId =
    getRegistryValue_(
      registryHeaders,
      rowValues,
      ['Response Spreadsheet ID']
    );


  /*
   * If destination is missing,
   * attempt FormApp.
   */

  if (!responseSpreadsheetId) {

    try {

      const form =
        FormApp.openById(formId);

      responseSpreadsheetId =
        safeString_(
          form.getDestinationId()
        );

    } catch (error) {

      warning_(
        functionName,
        'Could not determine response spreadsheet',
        formId + ' | ' + error.message
      );
    }
  }


  /*
   * If still unavailable, try opening form responses
   * through FormApp directly.
   */

  let responseSheet = null;

  if (responseSpreadsheetId) {

    responseSheet =
      findResponseSheet_(
        responseSpreadsheetId
      );
  }


  /*
   * If response spreadsheet is unavailable,
   * use FormApp response data.
   */

  if (!responseSheet) {

    return syncFormUsingFormApp_(
      registryRow,
      formId,
      formName
    );
  }


  /*
   * Update response spreadsheet ID.
   */

  setRegistryValue_(
    registry,
    registryRow,
    registryHeaders,
    'Response Spreadsheet ID',
    responseSpreadsheetId
  );


  /*
   * Read response sheet.
   */

  const values =
    responseSheet.getDataRange().getValues();

  if (values.length < 2) {

    updateRegistrySyncStatus_(
      registry,
      registryRow,
      registryHeaders,
      0,
      ''
    );

    return {
      success: true,
      records: 0
    };
  }


  const responseHeaders =
    values[0].map(safeString_);


  /*
   * Add every response field to MASTER_DATA.
   *
   * This is the key architecture:
   *
   * MASTER_DATA = user data.
   * API data is NOT inserted here.
   */

  const master =
    getSheet_(CONFIG.SHEETS.MASTER);

  const masterHeaders =
    ensureHeaders_(
      master,
      CONFIG.MASTER_METADATA_COLUMNS
        .concat(responseHeaders)
    );


  let recordsAdded = 0;
  const pendingRows = [];


  /*
   * Existing Response IDs
   */

  const existingResponseIds =
    getExistingResponseIds_(master);


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    try {

      const responseRow =
        values[i];

      const responseId =
        getResponseIdFromRow_(
          responseHeaders,
          responseRow
        );

      /*
       * Avoid duplicate rows.
       */

      if (
        responseId &&
        existingResponseIds.has(responseId)
      ) {
        continue;
      }


      const timestamp =
        getTimestampFromResponse_(
          responseHeaders,
          responseRow
        );


      const systemId =
        generateSystemId_();


      const output =
        new Array(masterHeaders.length)
          .fill('');


      setArrayValue_(
        output,
        masterHeaders,
        'System ID',
        systemId
      );

      setArrayValue_(
        output,
        masterHeaders,
        'Form Name',
        formName
      );

      setArrayValue_(
        output,
        masterHeaders,
        'Form ID',
        formId
      );

      setArrayValue_(
        output,
        masterHeaders,
        'Response ID',
        responseId
      );

      setArrayValue_(
        output,
        masterHeaders,
        'Response Timestamp',
        timestamp
      );

      setArrayValue_(
        output,
        masterHeaders,
        'Last Synced',
        new Date()
      );


      /*
       * Copy EXACT user response values.
       */

      for (
        let c = 0;
        c < responseHeaders.length;
        c++
      ) {

        const header =
          responseHeaders[c];

        const index =
          findHeaderIndex_(
            masterHeaders,
            header
          );

        if (index !== -1) {

          output[index] =
            responseRow[c];
        }
      }


      master.appendRow(output);

      recordsAdded++;

      if (responseId) {
        existingResponseIds.add(responseId);
      }


    } catch (rowError) {

      /*
       * One bad response must NOT stop
       * the entire form.
       */

      error_(
        functionName,
        'Response row skipped',
        'Form=' +
        formId +
        ' Row=' +
        (i + 1) +
        ' Error=' +
        rowError.message
      );
    }
  }


  if (pendingRows.length) {
    master.getRange(master.getLastRow() + 1, 1, pendingRows.length, masterHeaders.length).setValues(pendingRows);
  }

  updateRegistrySyncStatus_(
    registry,
    registryRow,
    registryHeaders,
    recordsAdded,
    ''
  );


  info_(
    functionName,
    'Form synchronization completed',
    formId +
    ' | Records=' +
    recordsAdded
  );


  return {
    success: true,
    records: recordsAdded
  };
}


/* ============================================================
   12. FORMAPP FALLBACK SYNC
   ============================================================ */

function syncFormUsingFormApp_(
  registryRow,
  formId,
  formName
) {

  const functionName =
    'syncFormUsingFormApp_';

  try {

    const form =
      FormApp.openById(formId);

    const responses =
      form.getResponses();

    if (!responses || responses.length === 0) {

      return {
        success: true,
        records: 0
      };
    }


    const master =
      getSheet_(CONFIG.SHEETS.MASTER);


    /*
     * Collect all questions across responses.
     */

    const questionSet = [];

    responses.forEach(function(response) {

      response
        .getItemResponses()
        .forEach(function(itemResponse) {

          const title =
            safeString_(
              itemResponse
                .getItem()
                .getTitle()
            );

          if (
            title &&
            questionSet.indexOf(title) === -1
          ) {

            questionSet.push(title);
          }

        });
    });


    const masterHeaders =
      ensureHeaders_(
        master,
        CONFIG.MASTER_METADATA_COLUMNS
          .concat(questionSet)
      );


    const existingResponseIds =
      getExistingResponseIds_(master);


    let added = 0;
    const pendingRows = [];


    responses.forEach(function(response) {

      try {

        const responseId =
          safeString_(
            response.getId()
          );

        if (
          responseId &&
          existingResponseIds.has(responseId)
        ) {
          return;
        }


        const output =
          new Array(masterHeaders.length)
            .fill('');


        const systemId =
          generateSystemId_();


        setArrayValue_(
          output,
          masterHeaders,
          'System ID',
          systemId
        );

        setArrayValue_(
          output,
          masterHeaders,
          'Form Name',
          formName
        );

        setArrayValue_(
          output,
          masterHeaders,
          'Form ID',
          formId
        );

        setArrayValue_(
          output,
          masterHeaders,
          'Response ID',
          responseId
        );

        setArrayValue_(
          output,
          masterHeaders,
          'Response Timestamp',
          response.getTimestamp()
        );

        setArrayValue_(
          output,
          masterHeaders,
          'Last Synced',
          new Date()
        );


        response
          .getItemResponses()
          .forEach(function(itemResponse) {

            const title =
              safeString_(
                itemResponse
                  .getItem()
                  .getTitle()
              );

            const answer =
              itemResponse.getResponse();

            setArrayValue_(
              output,
              masterHeaders,
              title,
              Array.isArray(answer)
                ? answer.join(', ')
                : answer
            );
          });


        pendingRows.push(output);

        existingResponseIds.add(responseId);

        added++;

      } catch (responseError) {

        error_(
          functionName,
          'One FormApp response skipped',
          responseError.message
        );
      }

    });


    if (pendingRows.length) {
      master.getRange(master.getLastRow() + 1, 1, pendingRows.length, masterHeaders.length).setValues(pendingRows);
    }

    info_(
      functionName,
      'FormApp fallback sync completed',
      formId +
      ' | Added=' +
      added
    );


    return {
      success: true,
      records: added
    };


  } catch (error) {

    /*
     * Important:
     * Store error in registry but do not crash app.
     */

    const registry =
      getSheet_(CONFIG.SHEETS.REGISTRY);

    const headers =
      getHeaders_(registry);

    setRegistryValue_(
      registry,
      registryRow,
      headers,
      'Status',
      'ERROR'
    );

    setRegistryValue_(
      registry,
      registryRow,
      headers,
      'Last Error',
      error.message
    );


    error_(
      functionName,
      'FormApp fallback failed',
      formId +
      ' | ' +
      error.stack
    );


    return {
      success: false,
      records: 0,
      error: error.message
    };
  }
}


/* ============================================================
   13. RESPONSE SHEET FINDER
   ============================================================ */

function findResponseSheet_(
  spreadsheetId
) {

  try {

    const ss =
      SpreadsheetApp.openById(
        spreadsheetId
      );

    const sheets =
      ss.getSheets();

    /*
     * Prefer sheet with response timestamp.
     */

    for (let i = 0; i < sheets.length; i++) {

      const sheet = sheets[i];

      const lastColumn =
        sheet.getLastColumn();

      if (lastColumn === 0) {
        continue;
      }

      const headers =
        sheet
          .getRange(
            1,
            1,
            1,
            lastColumn
          )
          .getValues()[0]
          .map(safeString_);

      const hasTimestamp =
        headers.some(function(header) {

          const n =
            normalize_(header);

          return (
            n === 'timestamp' ||
            n === 'response timestamp'
          );
        });

      if (hasTimestamp) {
        return sheet;
      }
    }


    /*
     * Fallback to first sheet.
     */

    if (sheets.length > 0) {
      return sheets[0];
    }


  } catch (error) {

    warning_(
      'findResponseSheet_',
      'Could not open response spreadsheet',
      spreadsheetId +
      ' | ' +
      error.message
    );
  }

  return null;
}


/* ============================================================
   14. MASTER DATA HELPERS
   ============================================================ */

function getExistingResponseIds_(master) {

  const result = new Set();

  if (master.getLastRow() < 2) {
    return result;
  }

  const headers =
    getHeaders_(master);

  const col =
    getColumnIndex_(
      headers,
      ['Response ID']
    );

  if (!col) {
    return result;
  }

  const values =
    master
      .getRange(
        2,
        col,
        master.getLastRow() - 1,
        1
      )
      .getValues();

  values.forEach(function(row) {

    const id =
      safeString_(row[0]);

    if (id) {
      result.add(id);
    }
  });

  return result;
}


function getResponseIdFromRow_(
  headers,
  row
) {

  const possible =
    [
      'Response ID',
      'Response Id',
      'ResponseID'
    ];

  const col =
    getColumnIndex_(
      headers,
      possible
    );

  if (col) {
    return safeString_(row[col - 1]);
  }

  /*
   * Google Forms response sheet often
   * does not contain Response ID.
   *
   * Generate deterministic ID from timestamp + row.
   */

  const timestampCol =
    getColumnIndex_(
      headers,
      [
        'Timestamp',
        'Response Timestamp'
      ]
    );

  const timestamp =
    timestampCol
      ? safeString_(row[timestampCol - 1])
      : '';

  return (
    timestamp +
    '|' +
    row.join('|')
  );
}


function getTimestampFromResponse_(
  headers,
  row
) {

  const col =
    getColumnIndex_(
      headers,
      [
        'Timestamp',
        'Response Timestamp'
      ]
    );

  if (!col) {
    return new Date();
  }

  return row[col - 1] || new Date();
}


/* ============================================================
   15. REGISTRY HELPERS
   ============================================================ */

function getRegistryValue_(
  headers,
  row,
  names
) {

  const col =
    getColumnIndex_(
      headers,
      names
    );

  if (!col) {
    return '';
  }

  return safeString_(
    row[col - 1]
  );
}


function setRegistryValue_(
  sheet,
  row,
  headers,
  name,
  value
) {

  const col =
    getColumnIndex_(
      headers,
      [name]
    );

  if (!col) {
    return;
  }

  sheet
    .getRange(
      row,
      col
    )
    .setValue(value);
}


function updateRegistrySyncStatus_(
  registry,
  row,
  headers,
  records,
  errorMessage
) {

  setRegistryValue_(
    registry,
    row,
    headers,
    'Status',
    errorMessage
      ? 'ERROR'
      : 'ACTIVE'
  );

  setRegistryValue_(
    registry,
    row,
    headers,
    'Last Sync',
    new Date()
  );

  setRegistryValue_(
    registry,
    row,
    headers,
    'Records Synced',
    records
  );

  setRegistryValue_(
    registry,
    row,
    headers,
    'Last Error',
    errorMessage || ''
  );
}


/* ============================================================
   16. ARRAY HELPERS
   ============================================================ */

function findHeaderIndex_(
  headers,
  header
) {

  const target =
    normalize_(header);

  for (
    let i = 0;
    i < headers.length;
    i++
  ) {

    if (
      normalize_(headers[i]) ===
      target
    ) {
      return i;
    }
  }

  return -1;
}


function setArrayValue_(
  array,
  headers,
  header,
  value
) {

  const index =
    findHeaderIndex_(
      headers,
      header
    );

  if (index !== -1) {
    array[index] = value;
  }
}


/* ============================================================
   17. VERIFICATION ENTRY
   ============================================================ */

function getVerificationInputSignature_(headers, row) {
  return [
    normalizeDOI_(getFirstRowValue_(headers, row, ['DOI', 'DOI / eISSN', 'Paper DOI', 'Publication DOI'])),
    normalize_(getFirstRowValue_(headers, row, ['Title of the Paper', 'Paper Title', 'Chapter Title', 'Book Title', 'Patent Title', 'Title of the Funded Project', 'Title'])),
    normalize_(getFirstRowValue_(headers, row, ['Author(s)', 'Authors', 'Inventor(s)', 'Principal Investigator'])),
    normalize_(getFirstRowValue_(headers, row, ['Publication / Research Output Category', 'Paper Type', 'Publication Category']))
  ].join('|');
}

function isVerificationFresh_(verification, verificationHeaders, existingRow, signature) {
  if (!CONFIG.PERFORMANCE.SKIP_UNCHANGED_AUTOMATICALLY || !existingRow) return false;
  const lastCol = getColumnIndex_(verificationHeaders, ['Last Verified']);
  const doiCol = getColumnIndex_(verificationHeaders, ['Input DOI']);
  const titleCol = getColumnIndex_(verificationHeaders, ['Input Title']);
  const authorsCol = getColumnIndex_(verificationHeaders, ['Input Authors']);
  const categoryCol = getColumnIndex_(verificationHeaders, ['Input Category']);
  if (!lastCol) return false;
  const row = verification.getRange(existingRow, 1, 1, verification.getLastColumn()).getValues()[0];
  const last = row[lastCol - 1];
  if (!(last instanceof Date) || !last.getTime()) return false;
  const ageHours = (Date.now() - last.getTime()) / 3600000;
  if (ageHours > Number(CONFIG.PERFORMANCE.VERIFY_CACHE_HOURS || 24)) return false;
  const oldSignature = [
    doiCol ? normalizeDOI_(row[doiCol - 1]) : '',
    titleCol ? normalize_(row[titleCol - 1]) : '',
    authorsCol ? normalize_(row[authorsCol - 1]) : '',
    categoryCol ? normalize_(row[categoryCol - 1]) : ''
  ].join('|');
  return oldSignature === signature;
}

function verifyAllPublications() {
  ensureCoreSheets_();
  const master = getSheet_(CONFIG.SHEETS.MASTER);
  const verification = getSheet_(CONFIG.SHEETS.VERIFICATION);

  if (master.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Verification', 'MASTER_DATA has no publication records.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const masterHeaders = getHeaders_(master);
  const data = master.getRange(2, 1, master.getLastRow() - 1, master.getLastColumn()).getValues();
  const verificationHeaders = ensureHeaders_(verification, CONFIG.VERIFICATION_COLUMNS);
  const existing = getVerificationMap_(verification);

  let verified = 0, notVerified = 0, pending = 0, manual = 0, apiErrors = 0, skipped = 0;

  data.forEach(function(row, index) {
    try {
      const systemId = getRowValue_(masterHeaders, row, ['System ID']);
      const existingRow = existing.get(systemId);
      const signature = getVerificationInputSignature_(masterHeaders, row);

      // Incremental verification: unchanged records are not sent to Crossref/Gemini again.
      if (isVerificationFresh_(verification, verificationHeaders, existingRow, signature)) {
        skipped++;
        const statusCol = getColumnIndex_(verificationHeaders, ['Verification Status']);
        const oldStatus = statusCol ? safeString_(verification.getRange(existingRow, statusCol).getValue()) : '';
        if (oldStatus === CONFIG.STATUS.VERIFIED) verified++;
        else if (oldStatus === CONFIG.STATUS.NOT_VERIFIED) notVerified++;
        else if (oldStatus === CONFIG.STATUS.PENDING) pending++;
        else if (oldStatus === CONFIG.STATUS.MANUAL_REVIEW) manual++;
        else if (oldStatus === CONFIG.STATUS.API_ERROR) apiErrors++;
        return;
      }

      const result = verifyMasterRow_(masterHeaders, row);
      upsertVerification_(verification, verificationHeaders, existing, result);

      if (result.status === CONFIG.STATUS.VERIFIED) verified++;
      else if (result.status === CONFIG.STATUS.NOT_VERIFIED) notVerified++;
      else if (result.status === CONFIG.STATUS.PENDING) pending++;
      else if (result.status === CONFIG.STATUS.MANUAL_REVIEW) manual++;
      else if (result.status === CONFIG.STATUS.API_ERROR) apiErrors++;
    } catch (rowError) {
      error_('verifyAllPublications', 'Publication verification row failed', 'Master row=' + (index + 2) + ' | ' + rowError.message);
    }
  });

  SpreadsheetApp.getUi().alert(
    'Verification Complete',
    'VERIFIED: ' + verified +
    '\nNOT VERIFIED: ' + notVerified +
    '\nPENDING: ' + pending +
    '\nMANUAL REVIEW: ' + manual +
    '\nAPI ERROR: ' + apiErrors +
    '\nSKIPPED (unchanged): ' + skipped,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function forceVerifyAllPublications() {
  const old = CONFIG.PERFORMANCE.SKIP_UNCHANGED_AUTOMATICALLY;
  CONFIG.PERFORMANCE.SKIP_UNCHANGED_AUTOMATICALLY = false;
  try {
    verifyAllPublications();
  } finally {
    CONFIG.PERFORMANCE.SKIP_UNCHANGED_AUTOMATICALLY = old;
  }
}


/* ============================================================
   18. VERIFY SELECTED ROW
   ============================================================ */

function verifySelectedRow() {

  try {

    const master =
      getSheet_(CONFIG.SHEETS.MASTER);

    const rowNumber =
      master.getActiveRange().getRow();

    if (rowNumber < 2) {

      SpreadsheetApp.getUi().alert(
        'Select a MASTER_DATA publication row.'
      );

      return;
    }


    const headers =
      getHeaders_(master);

    const row =
      master
        .getRange(
          rowNumber,
          1,
          1,
          master.getLastColumn()
        )
        .getValues()[0];


    const result =
      verifyMasterRow_(
        headers,
        row
      );


    const verification =
      getSheet_(
        CONFIG.SHEETS.VERIFICATION
      );

    const verificationHeaders =
      ensureHeaders_(
        verification,
        CONFIG.VERIFICATION_COLUMNS
      );


    const existing =
      getVerificationMap_(
        verification
      );


    upsertVerification_(
      verification,
      verificationHeaders,
      existing,
      result
    );


    applyVerificationDropdown();


    SpreadsheetApp.getUi().alert(
      'Verification Result',
      'Status: ' +
      result.status +

      '\nMatch Score: ' +
      result.matchScore +
      '%' +

      '\nMessage: ' +
      result.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );


  } catch (error) {

    error_(
      'verifySelectedRow',
      'Selected row verification failed',
      error.stack || error.message
    );

    SpreadsheetApp.getUi().alert(
      'Verification Error',
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/* ============================================================
   19. VERIFY MASTER ROW
   ============================================================ */

function verifyMasterRow_(
  masterHeaders,
  row
) {

  const systemId =
    getRowValue_(
      masterHeaders,
      row,
      ['System ID']
    );


  const inputDOI =
    normalizeDOI_(
      getFirstRowValue_(
        masterHeaders,
        row,
        [
          'DOI',
          'DOI / eISSN',
          'Paper DOI',
          'Publication DOI'
        ]
      )
    );


  const inputTitle =
    getFirstRowValue_(
      masterHeaders,
      row,
      [
        'Title of the Paper',
        'Paper Title',
        'Chapter Title',
        'Book Title',
        'Patent Title',
        'Title of the Funded Project',
        'Title'
      ]
    );


  const inputAuthors =
    getFirstRowValue_(
      masterHeaders,
      row,
      [
        'Author(s)',
        'Authors',
        'Inventor(s)',
        'Principal Investigator'
      ]
    );


  const inputCategory =
    getFirstRowValue_(
      masterHeaders,
      row,
      [
        'Publication / Research Output Category',
        'Paper Type',
        'Publication Category'
      ]
    );


  const formName =
    getFirstRowValue_(
      masterHeaders,
      row,
      [
        'Form Name'
      ]
    );


  /*
   * CASE 1
   * No DOI and no title
   */

  if (!inputDOI && !inputTitle) {

    return buildVerificationResult_({
      systemId: systemId,
      status: CONFIG.STATUS.PENDING,
      apiStatus: 'NOT CHECKED',
      apiDOI: '',
      apiTitle: '',
      apiAuthors: '',
      apiVenue: '',
      apiPublisher: '',
      matchScore: 0,
      message:
        'DOI and publication title are both blank.',
      source:
        'NO INPUT',
      inputDOI: '',
      inputTitle: '',
      inputAuthors: inputAuthors,
      inputCategory: inputCategory,
      sourceForm: formName,
      sourceDOI: ''
    });
  }


  /*
   * CASE 2
   * DOI available
   */

  if (inputDOI) {

    const doiResult =
      verifyByDOI_(
        inputDOI,
        inputTitle,
        inputAuthors
      );


    if (
      doiResult.found &&
      doiResult.score >= CONFIG.MATCH.VERIFIED
    ) {

      return buildVerificationResult_({
        systemId: systemId,
        status: CONFIG.STATUS.VERIFIED,
        apiStatus: 'SUCCESS',
        apiDOI: doiResult.doi,
        apiTitle: doiResult.title,
        apiAuthors: doiResult.authors,
        apiVenue: doiResult.venue,
        apiPublisher: doiResult.publisher,
        matchScore: doiResult.score,
        message:
          'DOI and publication metadata matched.',
        source:
          'Crossref DOI',
        inputDOI: inputDOI,
        inputTitle: inputTitle,
        inputAuthors: inputAuthors,
        inputCategory: inputCategory,
        sourceForm: formName,
        sourceDOI: inputDOI
      });
    }


    if (
      doiResult.found &&
      doiResult.score >= CONFIG.MATCH.MANUAL_REVIEW
    ) {

      return buildVerificationResult_({
        systemId: systemId,
        status: CONFIG.STATUS.MANUAL_REVIEW,
        apiStatus: 'SUCCESS',
        apiDOI: doiResult.doi,
        apiTitle: doiResult.title,
        apiAuthors: doiResult.authors,
        apiVenue: doiResult.venue,
        apiPublisher: doiResult.publisher,
        matchScore: doiResult.score,
        message:
          'DOI found but metadata needs manual review.',
        source:
          'Crossref DOI',
        inputDOI: inputDOI,
        inputTitle: inputTitle,
        inputAuthors: inputAuthors,
        inputCategory: inputCategory,
        sourceForm: formName,
        sourceDOI: inputDOI
      });
    }


    if (
      doiResult.found &&
      doiResult.score < CONFIG.MATCH.MANUAL_REVIEW
    ) {

      return buildVerificationResult_({
        systemId: systemId,
        status: CONFIG.STATUS.NOT_VERIFIED,
        apiStatus: 'SUCCESS',
        apiDOI: doiResult.doi,
        apiTitle: doiResult.title,
        apiAuthors: doiResult.authors,
        apiVenue: doiResult.venue,
        apiPublisher: doiResult.publisher,
        matchScore: doiResult.score,
        message:
          'DOI found, but submitted metadata does not match sufficiently.',
        source:
          'Crossref DOI',
        inputDOI: inputDOI,
        inputTitle: inputTitle,
        inputAuthors: inputAuthors,
        inputCategory: inputCategory,
        sourceForm: formName,
        sourceDOI: inputDOI
      });
    }


    /*
     * DOI not found.
     *
     * If title exists, automatically try title search.
     */

    if (inputTitle) {

      const titleResult =
        verifyByTitle_(
          inputTitle,
          inputAuthors
        );


      if (
        titleResult.found &&
        titleResult.score >= CONFIG.MATCH.VERIFIED
      ) {

        return buildVerificationResult_({
          systemId: systemId,
          status: CONFIG.STATUS.VERIFIED,
          apiStatus: 'SUCCESS',
          apiDOI: titleResult.doi,
          apiTitle: titleResult.title,
          apiAuthors: titleResult.authors,
          apiVenue: titleResult.venue,
          apiPublisher: titleResult.publisher,
          matchScore: titleResult.score,
          message:
            'DOI was not directly found, but title metadata matched.',
          source:
            'Crossref Title Search',
          inputDOI: inputDOI,
          inputTitle: inputTitle,
          inputAuthors: inputAuthors,
          inputCategory: inputCategory,
          sourceForm: formName,
          sourceDOI: inputDOI
        });
      }


      if (
        titleResult.found &&
        titleResult.score >= CONFIG.MATCH.MANUAL_REVIEW
      ) {

        return buildVerificationResult_({
          systemId: systemId,
          status: CONFIG.STATUS.MANUAL_REVIEW,
          apiStatus: 'SUCCESS',
          apiDOI: titleResult.doi,
          apiTitle: titleResult.title,
          apiAuthors: titleResult.authors,
          apiVenue: titleResult.venue,
          apiPublisher: titleResult.publisher,
          matchScore: titleResult.score,
          message:
            'Title candidate found but requires manual review.',
          source:
            'Crossref Title Search',
          inputDOI: inputDOI,
          inputTitle: inputTitle,
          inputAuthors: inputAuthors,
          inputCategory: inputCategory,
          sourceForm: formName,
          sourceDOI: inputDOI
        });
      }


      return buildVerificationResult_({
        systemId: systemId,
        status: CONFIG.STATUS.NOT_VERIFIED,
        apiStatus: 'NOT FOUND',
        apiDOI: '',
        apiTitle: titleResult.title || '',
        apiAuthors: titleResult.authors || '',
        apiVenue: titleResult.venue || '',
        apiPublisher: titleResult.publisher || '',
        matchScore: titleResult.score || 0,
        message:
          'No sufficiently matching Crossref record was found.',
        source:
          'Crossref',
        inputDOI: inputDOI,
        inputTitle: inputTitle,
        inputAuthors: inputAuthors,
        inputCategory: inputCategory,
        sourceForm: formName,
        sourceDOI: inputDOI
      });
    }


    return buildVerificationResult_({
      systemId: systemId,
      status: CONFIG.STATUS.NOT_VERIFIED,
      apiStatus: 'NOT FOUND',
      apiDOI: '',
      apiTitle: '',
      apiAuthors: '',
      apiVenue: '',
      apiPublisher: '',
      matchScore: 0,
      message:
        'DOI could not be found in Crossref.',
      source:
        'Crossref DOI',
      inputDOI: inputDOI,
      inputTitle: inputTitle,
      inputAuthors: inputAuthors,
      inputCategory: inputCategory,
      sourceForm: formName,
      sourceDOI: inputDOI
    });
  }


  /*
   * CASE 3
   * No DOI → title verification
   */

  const titleResult =
    verifyByTitle_(
      inputTitle,
      inputAuthors
    );


  if (
    titleResult.apiError
  ) {

    return buildVerificationResult_({
      systemId: systemId,
      status: CONFIG.STATUS.API_ERROR,
      apiStatus: 'ERROR',
      apiDOI: '',
      apiTitle: '',
      apiAuthors: '',
      apiVenue: '',
      apiPublisher: '',
      matchScore: 0,
      message:
        titleResult.message,
      source:
        'Crossref Title Search',
      inputDOI: '',
      inputTitle: inputTitle,
      inputAuthors: inputAuthors,
      inputCategory: inputCategory,
      sourceForm: formName,
      sourceDOI: ''
    });
  }


  if (
    titleResult.found &&
    titleResult.score >= CONFIG.MATCH.VERIFIED
  ) {

    return buildVerificationResult_({
      systemId: systemId,
      status: CONFIG.STATUS.VERIFIED,
      apiStatus: 'SUCCESS',
      apiDOI: titleResult.doi,
      apiTitle: titleResult.title,
      apiAuthors: titleResult.authors,
      apiVenue: titleResult.venue,
      apiPublisher: titleResult.publisher,
      matchScore: titleResult.score,
      message:
        'Publication title matched Crossref metadata.',
      source:
        'Crossref Title Search',
      inputDOI: '',
      inputTitle: inputTitle,
      inputAuthors: inputAuthors,
      inputCategory: inputCategory,
      sourceForm: formName,
      sourceDOI: ''
    });
  }


  if (
    titleResult.found &&
    titleResult.score >= CONFIG.MATCH.MANUAL_REVIEW
  ) {

    return buildVerificationResult_({
      systemId: systemId,
      status: CONFIG.STATUS.MANUAL_REVIEW,
      apiStatus: 'SUCCESS',
      apiDOI: titleResult.doi,
      apiTitle: titleResult.title,
      apiAuthors: titleResult.authors,
      apiVenue: titleResult.venue,
      apiPublisher: titleResult.publisher,
      matchScore: titleResult.score,
      message:
        'Publication title candidate found; manual review required.',
      source:
        'Crossref Title Search',
      inputDOI: '',
      inputTitle: inputTitle,
      inputAuthors: inputAuthors,
      inputCategory: inputCategory,
      sourceForm: formName,
      sourceDOI: ''
    });
  }


  return buildVerificationResult_({
    systemId: systemId,
    status: CONFIG.STATUS.NOT_VERIFIED,
    apiStatus: 'NOT FOUND',
    apiDOI: '',
    apiTitle: titleResult.title || '',
    apiAuthors: titleResult.authors || '',
    apiVenue: titleResult.venue || '',
    apiPublisher: titleResult.publisher || '',
    matchScore: titleResult.score || 0,
    message:
      'No reliable matching publication found.',
    source:
      'Crossref Title Search',
    inputDOI: '',
    inputTitle: inputTitle,
    inputAuthors: inputAuthors,
    inputCategory: inputCategory,
    sourceForm: formName,
    sourceDOI: ''
  });
}


/* ============================================================
   20. DOI VERIFICATION
   ============================================================ */

function verifyByDOI_(
  doi,
  inputTitle,
  inputAuthors
) {

  try {

    const cleanDOI =
      normalizeDOI_(doi);

    if (!cleanDOI) {

      return {
        found: false,
        score: 0
      };
    }


    const cache = CacheService.getScriptCache();
    const cacheKey = 'cr_doi_' + Utilities.base64EncodeWebSafe(cleanDOI).substring(0, 200);
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (ignore) {}
    }

    const url =
      CONFIG.CROSSREF.BASE_URL +
      '/' +
      encodeURIComponent(cleanDOI);


    const response =
      UrlFetchApp.fetch(
        addMailto_(url),
        {
          method: 'get',
          muteHttpExceptions: true,
          headers: {
            'Accept': 'application/json'
          }
        }
      );


    const code =
      response.getResponseCode();


    if (code < 200 || code >= 300) {

      return {
        found: false,
        score: 0,
        message:
          'Crossref HTTP ' + code
      };
    }


    const json =
      JSON.parse(
        response.getContentText()
      );


    const item =
      json &&
      json.message
        ? json.message
        : null;


    if (!item) {

      return {
        found: false,
        score: 0
      };
    }


    const metadata =
      crossrefItemToMetadata_(item);


    const score =
      calculateMatchScore_(
        inputTitle,
        inputAuthors,
        metadata.title,
        metadata.authors
      );


    const result = {
      found: true,
      score: score,
      doi: metadata.doi,
      title: metadata.title,
      authors: metadata.authors,
      venue: metadata.venue,
      publisher: metadata.publisher
    };
    try { cache.put(cacheKey, JSON.stringify(result), CONFIG.CROSSREF.CACHE_SECONDS); } catch (ignore) {}
    return result;


  } catch (error) {

    error_(
      'verifyByDOI_',
      'Crossref DOI verification error',
      error.stack || error.message
    );


    return {
      found: false,
      score: 0,
      apiError: true,
      message:
        'Crossref DOI API error: ' +
        error.message
    };
  }
}


/* ============================================================
   21. TITLE VERIFICATION
   ============================================================ */

function verifyByTitle_(
  inputTitle,
  inputAuthors
) {

  try {

    const title =
      safeString_(inputTitle);

    if (!title) {

      return {
        found: false,
        score: 0
      };
    }


    const cache = CacheService.getScriptCache();
    const cacheKey = 'cr_title_' + Utilities.base64EncodeWebSafe(normalize_(title)).substring(0, 200);
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (ignore) {}
    }

    const url =
      CONFIG.CROSSREF.BASE_URL +
      '?query.title=' +
      encodeURIComponent(title) +
      '&rows=' +
      CONFIG.CROSSREF.ROWS;


    const response =
      UrlFetchApp.fetch(
        addMailto_(url),
        {
          method: 'get',
          muteHttpExceptions: true,
          headers: {
            'Accept': 'application/json'
          }
        }
      );


    const code =
      response.getResponseCode();


    if (code < 200 || code >= 300) {

      return {
        found: false,
        score: 0,
        apiError: true,
        message:
          'Crossref HTTP ' + code
      };
    }


    const json =
      JSON.parse(
        response.getContentText()
      );


    const items =
      json &&
      json.message &&
      Array.isArray(json.message.items)
        ? json.message.items
        : [];


    if (items.length === 0) {

      return {
        found: false,
        score: 0
      };
    }


    let best = null;


    items.forEach(function(item) {

      const metadata =
        crossrefItemToMetadata_(item);


      const score =
        calculateMatchScore_(
          title,
          inputAuthors,
          metadata.title,
          metadata.authors
        );


      if (
        !best ||
        score > best.score
      ) {

        best = {
          found: true,
          score: score,
          doi: metadata.doi,
          title: metadata.title,
          authors: metadata.authors,
          venue: metadata.venue,
          publisher: metadata.publisher
        };
      }

    });


    const result = best || { found: false, score: 0 };
    try { cache.put(cacheKey, JSON.stringify(result), CONFIG.CROSSREF.CACHE_SECONDS); } catch (ignore) {}
    return result;


  } catch (error) {

    error_(
      'verifyByTitle_',
      'Crossref title verification error',
      error.stack || error.message
    );


    return {
      found: false,
      score: 0,
      apiError: true,
      message:
        'Crossref title API error: ' +
        error.message
    };
  }
}


/* ============================================================
   22. CROSSREF METADATA
   ============================================================ */

function crossrefItemToMetadata_(
  item
) {

  const title =
    Array.isArray(item.title)
      ? safeString_(item.title[0])
      : '';


  const authors =
    Array.isArray(item.author)
      ? item.author
          .map(function(author) {

            return [
              safeString_(author.given),
              safeString_(author.family)
            ]
              .filter(Boolean)
              .join(' ');

          })
          .filter(Boolean)
          .join('; ')
      : '';


  const venue =
    Array.isArray(item.container-title)
      ? safeString_(item.container-title[0])
      : '';


  return {

    doi:
      normalizeDOI_(
        item.DOI || ''
      ),

    title:
      title,

    authors:
      authors,

    venue:
      venue,

    publisher:
      safeString_(
        item.publisher || ''
      )
  };
}


/* ============================================================
   23. MATCH SCORE
   ============================================================ */

function calculateMatchScore_(
  inputTitle,
  inputAuthors,
  apiTitle,
  apiAuthors
) {

  const titleScore =
    similarityScore_(
      normalize_(inputTitle),
      normalize_(apiTitle)
    );


  let authorScore = 100;

  if (
    safeString_(inputAuthors) &&
    safeString_(apiAuthors)
  ) {

    authorScore =
      similarityScore_(
        normalize_(inputAuthors),
        normalize_(apiAuthors)
      );
  }


  /*
   * Title is more important than authors.
   */

  const score =
    Math.round(
      (
        titleScore * 0.80 +
        authorScore * 0.20
      )
    );


  return Math.max(
    0,
    Math.min(
      100,
      score
    )
  );
}


/* ============================================================
   24. STRING SIMILARITY
   ============================================================ */

function similarityScore_(
  a,
  b
) {

  if (!a || !b) {
    return 0;
  }

  if (a === b) {
    return 100;
  }


  const tokensA =
    new Set(
      a.split(' ')
        .filter(Boolean)
    );


  const tokensB =
    new Set(
      b.split(' ')
        .filter(Boolean)
    );


  let intersection = 0;


  tokensA.forEach(function(token) {

    if (tokensB.has(token)) {
      intersection++;
    }

  });


  const union =
    new Set(
      Array.from(tokensA)
        .concat(
          Array.from(tokensB)
        )
    ).size;


  const jaccard =
    union
      ? intersection / union
      : 0;


  /*
   * Also check containment.
   */

  let containment = 0;

  if (
    a.indexOf(b) !== -1 ||
    b.indexOf(a) !== -1
  ) {

    containment = 1;
  }


  const score =
    Math.round(
      (
        jaccard * 0.70 +
        containment * 0.30
      ) * 100
    );


  return score;
}


/* ============================================================
   25. VERIFICATION RESULT
   ============================================================ */

function buildDOIUrl_(doi) {
  const clean = normalizeDOI_(doi);
  return clean ? 'https://doi.org/' + encodeURIComponent(clean) : '';
}


function buildVerificationResult_(data) {
  const result = {
    systemId: data.systemId || '',
    status: data.status || CONFIG.STATUS.PENDING,
    apiStatus: data.apiStatus || '',
    apiDOI: data.apiDOI || '',
    apiTitle: data.apiTitle || '',
    apiAuthors: data.apiAuthors || '',
    apiVenue: data.apiVenue || '',
    apiPublisher: data.apiPublisher || '',
    matchScore: Number(data.matchScore || 0),
    message: data.message || '',
    lastVerified: new Date(),
    source: data.source || '',
    inputDOI: data.inputDOI || '',
    inputTitle: data.inputTitle || '',
    inputAuthors: data.inputAuthors || '',
    inputCategory: data.inputCategory || '',
    sourceForm: data.sourceForm || '',
    sourceDOI: data.sourceDOI || '',
    apiURL: data.apiURL || buildDOIUrl_(data.apiDOI),
    sourceURL: data.sourceURL || buildDOIUrl_(data.sourceDOI),
    aiStatus: 'NOT CONFIGURED',
    aiMatchScore: 0,
    aiConfidence: '',
    aiDecision: '',
    aiExplanation: '',
    aiRecommendation: '',
    aiTitleSimilarity: 0,
    aiAuthorSimilarity: 0,
    aiCheckedAt: ''
  };

  const aiStatuses = CONFIG.AI && Array.isArray(CONFIG.AI.RUN_FOR_STATUSES)
    ? CONFIG.AI.RUN_FOR_STATUSES
    : [];

  if (
    CONFIG.AI &&
    CONFIG.AI.ENABLED &&
    getAIKey_() &&
    aiStatuses.indexOf(result.status) !== -1
  ) {
    try {
      const ai = analyzePublicationWithAI_(result);
      result.aiStatus = ai.status || 'ERROR';
      result.aiMatchScore = Number(ai.matchScore || 0);
      result.aiConfidence = ai.confidence || '';
      result.aiDecision = ai.decision || '';
      result.aiExplanation = ai.explanation || '';
      result.aiRecommendation = ai.recommendation || '';
      result.aiTitleSimilarity = Number(ai.titleSimilarity || 0);
      result.aiAuthorSimilarity = Number(ai.authorSimilarity || 0);
      result.aiCheckedAt = new Date();
    } catch (e) {
      result.aiStatus = 'ERROR';
      result.aiExplanation = 'AI analysis failed: ' + e.message;
      result.aiCheckedAt = new Date();
      warning_('buildVerificationResult_', 'AI analysis skipped', e.message);
    }
  }

  return result;
}


/* ============================================================
   26. VERIFICATION UPSERT
   ============================================================ */

function getVerificationMap_(
  sheet
) {

  const map = new Map();

  if (sheet.getLastRow() < 2) {
    return map;
  }


  const headers =
    getHeaders_(sheet);


  const systemIdColumn =
    getColumnIndex_(
      headers,
      ['System ID']
    );


  if (!systemIdColumn) {
    return map;
  }


  const data =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        sheet.getLastColumn()
      )
      .getValues();


  data.forEach(function(row, index) {

    const id =
      safeString_(
        row[systemIdColumn - 1]
      );

    if (id) {

      map.set(
        id,
        index + 2
      );
    }

  });


  return map;
}


function upsertVerification_(
  sheet,
  headers,
  map,
  result
) {

  const output =
    new Array(headers.length)
      .fill('');


  setArrayValue_(
    output,
    headers,
    'System ID',
    result.systemId
  );

  setArrayValue_(
    output,
    headers,
    'Verification Status',
    result.status
  );

  setArrayValue_(
    output,
    headers,
    'API Status',
    result.apiStatus
  );

  setArrayValue_(
    output,
    headers,
    'API DOI',
    result.apiDOI
  );

  setArrayValue_(
    output,
    headers,
    'API Title',
    result.apiTitle
  );

  setArrayValue_(
    output,
    headers,
    'API Authors',
    result.apiAuthors
  );

  setArrayValue_(
    output,
    headers,
    'API Journal / Conference / Book',
    result.apiVenue
  );

  setArrayValue_(
    output,
    headers,
    'API Publisher',
    result.apiPublisher
  );

  setArrayValue_(
    output,
    headers,
    'Match Score %',
    result.matchScore
  );

  setArrayValue_(
    output,
    headers,
    'Verification Message',
    result.message
  );

  setArrayValue_(
    output,
    headers,
    'Last Verified',
    result.lastVerified
  );

  setArrayValue_(
    output,
    headers,
    'Verification Source',
    result.source
  );

  setArrayValue_(
    output,
    headers,
    'Input DOI',
    result.inputDOI
  );

  setArrayValue_(
    output,
    headers,
    'Input Title',
    result.inputTitle
  );

  setArrayValue_(
    output,
    headers,
    'Input Authors',
    result.inputAuthors
  );

  setArrayValue_(
    output,
    headers,
    'Input Category',
    result.inputCategory
  );

  setArrayValue_(
    output,
    headers,
    'Source Form',
    result.sourceForm
  );

  setArrayValue_(
    output,
    headers,
    'Source DOI',
    result.sourceDOI
  );

  setArrayValue_(output, headers, 'API URL', result.apiURL);
  setArrayValue_(output, headers, 'Source URL', result.sourceURL);
  setArrayValue_(output, headers, 'AI Status', result.aiStatus);
  setArrayValue_(output, headers, 'AI Match Score %', result.aiMatchScore);
  setArrayValue_(output, headers, 'AI Confidence', result.aiConfidence);
  setArrayValue_(output, headers, 'AI Decision', result.aiDecision);
  setArrayValue_(output, headers, 'AI Explanation', result.aiExplanation);
  setArrayValue_(output, headers, 'AI Recommendation', result.aiRecommendation);
  setArrayValue_(output, headers, 'AI Title Similarity %', result.aiTitleSimilarity);
  setArrayValue_(output, headers, 'AI Author Similarity %', result.aiAuthorSimilarity);
  setArrayValue_(output, headers, 'AI Checked At', result.aiCheckedAt);


  const existingRow =
    map.get(
      result.systemId
    );


  if (existingRow) {

    sheet
      .getRange(
        existingRow,
        1,
        1,
        headers.length
      )
      .setValues([output]);

  } else {

    sheet.appendRow(output);

    map.set(
      result.systemId,
      sheet.getLastRow()
    );
  }


  /*
   * Format status immediately.
   */

  colorVerificationStatus_(
    sheet,
    existingRow || sheet.getLastRow(),
    headers
  );
}


/* ============================================================
   27. STATUS DROPDOWN
   ============================================================ */

function applyVerificationDropdown() {

  try {

    const sheet =
      getSheet_(
        CONFIG.SHEETS.VERIFICATION
      );


    const headers =
      ensureHeaders_(
        sheet,
        CONFIG.VERIFICATION_COLUMNS
      );


    const statusColumn =
      getColumnIndex_(
        headers,
        ['Verification Status']
      );


    if (!statusColumn) {
      return;
    }


    const maxRows =
      Math.max(
        sheet.getMaxRows(),
        1000
      );


    const rule =
      SpreadsheetApp
        .newDataValidation()
        .requireValueInList(
          CONFIG.STATUS_LIST,
          true
        )
        .setAllowInvalid(false)
        .build();


    sheet
      .getRange(
        2,
        statusColumn,
        maxRows - 1,
        1
      )
      .setDataValidation(rule);


    /*
     * Conditional formatting.
     */

    const range =
      sheet.getRange(
        2,
        statusColumn,
        maxRows - 1,
        1
      );


    const rules =
      sheet
        .getConditionalFormatRules()
        .filter(function(rule) {

          const ranges =
            rule.getRanges();

          return !ranges.some(
            function(r) {

              return (
                r.getSheet().getSheetId() ===
                sheet.getSheetId() &&
                r.getColumn() ===
                statusColumn
              );

            }
          );

        });


    rules.push(

      SpreadsheetApp
        .newConditionalFormatRule()
        .whenTextEqualTo('VERIFIED')
        .setBackground('#B7E1CD')
        .setRanges([range])
        .build()

    );


    rules.push(

      SpreadsheetApp
        .newConditionalFormatRule()
        .whenTextEqualTo('NOT VERIFIED')
        .setBackground('#F4CCCC')
        .setRanges([range])
        .build()

    );


    rules.push(

      SpreadsheetApp
        .newConditionalFormatRule()
        .whenTextEqualTo('PENDING')
        .setBackground('#FFF2CC')
        .setRanges([range])
        .build()

    );


    rules.push(

      SpreadsheetApp
        .newConditionalFormatRule()
        .whenTextEqualTo('MANUAL REVIEW')
        .setBackground('#D9EAD3')
        .setRanges([range])
        .build()

    );


    rules.push(

      SpreadsheetApp
        .newConditionalFormatRule()
        .whenTextEqualTo('API ERROR')
        .setBackground('#EAD1DC')
        .setRanges([range])
        .build()

    );


    sheet.setConditionalFormatRules(
      rules
    );


  } catch (error) {

    /*
     * Dropdown must never crash application.
     */

    error_(
      'applyVerificationDropdown',
      'Could not apply dropdown',
      error.message
    );
  }
}


/* ============================================================
   27A. AI INTEGRATION - GROQ FREE TIER
   ============================================================ */

function configureAIKey() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Configure Groq AI',
    'Paste your Groq API key. The key is stored in Script Properties and is never written to any sheet.',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const key = safeString_(response.getResponseText());
  if (!key) {
    ui.alert('Groq API key cannot be blank.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(
    CONFIG.AI.API_KEY_PROPERTY,
    key
  );
  ui.alert('Groq API key saved securely.');
  info_('configureAIKey', 'Groq API key configured');
}

function getAIKey_() {
  return safeString_(
    PropertiesService.getScriptProperties().getProperty(
      CONFIG.AI.API_KEY_PROPERTY
    )
  );
}

function testAIIntegration() {
  try {
    if (!getAIKey_()) {
      SpreadsheetApp.getUi().alert('Configure the Groq API key first.');
      return;
    }

    const result = analyzePublicationWithAI_({
      inputTitle: 'Artificial intelligence in higher education',
      inputAuthors: 'Test Author',
      inputDOI: '',
      apiTitle: 'Artificial intelligence in higher education',
      apiAuthors: 'Test Author',
      apiDOI: '',
      apiVenue: 'Test Journal',
      apiPublisher: 'Test Publisher',
      matchScore: 100,
      status: CONFIG.STATUS.PENDING
    });

    SpreadsheetApp.getUi().alert(
      'Groq AI Test Result\n\n' +
      JSON.stringify(result, null, 2)
    );
  } catch (e) {
    error_('testAIIntegration', 'Groq AI test failed', e.stack || e.message);
    SpreadsheetApp.getUi().alert('Groq AI test failed: ' + e.message);
  }
}

function analyzePublicationWithAI_(publication) {
  const key = getAIKey_();
  if (!key) return { status: 'NOT CONFIGURED' };

  const data = {
    inputTitle: safeString_(publication.inputTitle),
    inputAuthors: safeString_(publication.inputAuthors),
    inputDOI: normalizeDOI_(publication.inputDOI),
    apiTitle: safeString_(publication.apiTitle),
    apiAuthors: safeString_(publication.apiAuthors),
    apiDOI: normalizeDOI_(publication.apiDOI),
    apiVenue: safeString_(publication.apiVenue),
    apiPublisher: safeString_(publication.apiPublisher),
    crossrefMatchScore: Number(publication.matchScore || 0),
    crossrefStatus: safeString_(publication.status)
  };

  if (!data.inputTitle && !data.inputDOI && !data.apiTitle && !data.apiDOI) {
    return {
      status: 'INSUFFICIENT_DATA',
      model: CONFIG.AI.MODEL,
      matchScore: 0,
      confidence: 'LOW',
      decision: 'PENDING',
      explanation: 'Insufficient publication metadata for AI comparison.',
      recommendation: 'Provide a title or DOI for verification.',
      titleSimilarity: 0,
      authorSimilarity: 0
    };
  }

  // Avoid repeated AI calls for the same metadata.
  const cache = CacheService.getScriptCache();
  const signature = JSON.stringify(data);
  const cacheKey = 'groq_ai_' + Utilities.base64EncodeWebSafe(signature).substring(0, 200);
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (ignore) {}
  }

  const prompt = [
    'You are a research-publication metadata verification assistant.',
    'Compare the submitted metadata with the trusted Crossref metadata supplied below.',
    'Do not invent facts. Do not use outside knowledge. Use only the supplied fields.',
    'Return ONLY one valid JSON object.',
    'Required keys: status, matchScore, confidence, decision, explanation, recommendation, titleSimilarity, authorSimilarity.',
    'status: OK or INSUFFICIENT_DATA.',
    'decision: VERIFIED, NOT VERIFIED, MANUAL REVIEW, or PENDING.',
    'confidence: HIGH, MEDIUM, or LOW.',
    'All score values must be integers from 0 to 100.',
    'Keep explanation and recommendation concise.',
    JSON.stringify(data)
  ].join('\n');

  const payload = {
    model: CONFIG.AI.MODEL,
    messages: [
      { role: 'user', content: prompt }
    ],
    temperature: 0,
    max_completion_tokens: Number(CONFIG.AI.MAX_COMPLETION_TOKENS || 600),
    reasoning_effort: 'low',
    reasoning_format: 'hidden',
    response_format: { type: 'json_object' },
    stream: false
  };

  let response;
  let code = 0;
  let body = '';

  try {
    response = UrlFetchApp.fetch(CONFIG.AI.BASE_URL, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: {
        Authorization: 'Bearer ' + key,
        Accept: 'application/json'
      },
      payload: JSON.stringify(payload)
    });
    code = response.getResponseCode();
    body = response.getContentText();
  } catch (e) {
    return aiUnavailableResult_(0, e.message);
  }

  if (code < 200 || code >= 300) {
    // Do not retry automatically: retries make free-tier rate limits worse.
    return aiUnavailableResult_(code, body);
  }

  try {
    const envelope = JSON.parse(body);
    const text = safeString_(
      envelope &&
      envelope.choices &&
      envelope.choices[0] &&
      envelope.choices[0].message &&
      envelope.choices[0].message.content
    );

    if (!text) {
      return aiUnavailableResult_(code, 'Groq returned an empty response.');
    }

    const clean = text
      .replace(/^```json\s*/i, '')
      .replace(/^```\s*/i, '')
      .replace(/\s*```$/i, '')
      .trim();

    const result = JSON.parse(clean);
    const output = {
      status: safeString_(result.status) || 'OK',
      model: CONFIG.AI.MODEL,
      matchScore: clampAI_(result.matchScore),
      confidence: safeString_(result.confidence) || 'LOW',
      decision: safeString_(result.decision) || 'PENDING',
      explanation: safeString_(result.explanation),
      recommendation: safeString_(result.recommendation),
      titleSimilarity: clampAI_(result.titleSimilarity),
      authorSimilarity: clampAI_(result.authorSimilarity)
    };

    try {
      cache.put(
        cacheKey,
        JSON.stringify(output),
        Number(CONFIG.AI.CACHE_SECONDS || 21600)
      );
    } catch (ignore) {}

    return output;
  } catch (e) {
    return aiUnavailableResult_(code, 'Invalid Groq JSON response: ' + e.message);
  }
}

function aiUnavailableResult_(code, body) {
  const detail = safeString_(body).substring(0, 900);
  warning_(
    'analyzePublicationWithAI_',
    'Groq AI unavailable; Crossref result retained',
    'HTTP=' + code + ' | ' + detail
  );

  return {
    status: 'TEMPORARILY_UNAVAILABLE',
    model: CONFIG.AI.MODEL,
    matchScore: 0,
    confidence: 'LOW',
    decision: 'PENDING',
    explanation: 'Groq AI is temporarily unavailable. Crossref verification is retained.',
    recommendation: code === 429
      ? 'Retry after the Groq free-tier rate limit resets.'
      : 'Retry AI verification later.',
    titleSimilarity: 0,
    authorSimilarity: 0,
    errorCode: code,
    errorDetails: detail
  };
}

function clampAI_(value) {
  const n = Number(value);
  return isFinite(n)
    ? Math.max(0, Math.min(100, Math.round(n)))
    : 0;
}

/* ============================================================
   28. STATUS COLOR
   ============================================================ */

function colorVerificationStatus_(
  sheet,
  row,
  headers
) {

  try {

    const column =
      getColumnIndex_(
        headers,
        ['Verification Status']
      );

    if (!column || row < 2) {
      return;
    }

    /*
     * Conditional formatting handles
     * visual status automatically.
     */

  } catch (error) {

    warning_(
      'colorVerificationStatus_',
      'Status color skipped',
      error.message
    );
  }
}


/* ============================================================
   29. ROW VALUE HELPERS
   ============================================================ */

function getRowValue_(
  headers,
  row,
  names
) {

  const col =
    getColumnIndex_(
      headers,
      names
    );

  if (!col) {
    return '';
  }

  return safeString_(
    row[col - 1]
  );
}


function getFirstRowValue_(
  headers,
  row,
  names
) {

  for (
    let i = 0;
    i < names.length;
    i++
  ) {

    const value =
      getRowValue_(
        headers,
        row,
        [names[i]]
      );

    if (value) {
      return value;
    }
  }

  return '';
}


/* ============================================================
   30. MAILTO
   ============================================================ */

function addMailto_(url) {

  if (!CONFIG.CROSSREF.MAILTO) {
    return url;
  }

  const separator =
    url.indexOf('?') === -1
      ? '?'
      : '&';

  return (
    url +
    separator +
    'mailto=' +
    encodeURIComponent(
      CONFIG.CROSSREF.MAILTO
    )
  );
}


/* ============================================================
   31. AUTOMATIC TRIGGERS
   ============================================================ */

function installAutomaticTriggers() {

  try {

    const existing =
      ScriptApp.getProjectTriggers();


    existing.forEach(function(trigger) {

      const handler =
        trigger.getHandlerFunction();

      if (
        handler === 'automaticSync_' ||
        handler === 'automaticVerification_'
      ) {

        ScriptApp.deleteTrigger(
          trigger
        );
      }

    });


    /*
     * Sync every 10 minutes.
     */

    ScriptApp.newTrigger(
      'automaticSync_'
    )
      .timeBased()
      .everyMinutes(10)
      .create();


    /*
     * Verification every 15 minutes.
     */

    ScriptApp.newTrigger(
      'automaticVerification_'
    )
      .timeBased()
      .everyMinutes(15)
      .create();


    info_(
      'installAutomaticTriggers',
      'Automatic triggers installed',
      'Sync=10m, Verification=15m'
    );


  } catch (error) {

    error_(
      'installAutomaticTriggers',
      'Trigger installation failed',
      error.message
    );
  }
}


function automaticSync_() {

  try {

    syncAllFormsSilent_();

  } catch (error) {

    error_(
      'automaticSync_',
      'Automatic sync failed',
      error.stack || error.message
    );
  }
}


function automaticVerification_() {

  try {

    verifyAllSilent_();

  } catch (error) {

    error_(
      'automaticVerification_',
      'Automatic verification failed',
      error.stack || error.message
    );
  }
}


function syncAllFormsSilent_() {
  const registry = getSheet_(CONFIG.SHEETS.REGISTRY);
  if (registry.getLastRow() < 2) return;

  const data = registry.getDataRange().getValues();
  const headers = data[0];
  const formIdColumn = getColumnIndex_(headers, ['Form ID']);
  const statusColumn = getColumnIndex_(headers, ['Status']);
  if (!formIdColumn) return;

  for (let i = 1; i < data.length; i++) {
    const formId = safeString_(data[i][formIdColumn - 1]);
    const status = statusColumn ? safeString_(data[i][statusColumn - 1]).toUpperCase() : '';
    if (!formId || status === 'INACTIVE') continue;
    try {
      syncOneForm_(i + 1);
    } catch (error) {
      error_('syncAllFormsSilent_', 'Form skipped during automatic sync', 'Row=' + (i + 1) + ' | ' + error.message);
    }
  }
}


function verifyAllSilent_() {
  const master = getSheet_(CONFIG.SHEETS.MASTER);
  if (master.getLastRow() < 2) return;

  const headers = getHeaders_(master);
  const data = master.getRange(2, 1, master.getLastRow() - 1, master.getLastColumn()).getValues();
  const verification = getSheet_(CONFIG.SHEETS.VERIFICATION);
  const verificationHeaders = ensureHeaders_(verification, CONFIG.VERIFICATION_COLUMNS);
  const existing = getVerificationMap_(verification);

  data.forEach(function(row) {
    try {
      const systemId = getRowValue_(headers, row, ['System ID']);
      const existingRow = existing.get(systemId);
      const signature = getVerificationInputSignature_(headers, row);
      if (isVerificationFresh_(verification, verificationHeaders, existingRow, signature)) return;

      const result = verifyMasterRow_(headers, row);
      upsertVerification_(verification, verificationHeaders, existing, result);
    } catch (error) {
      error_('verifyAllSilent_', 'One publication skipped', error.message);
    }
  });
}


/* ============================================================
   32. FULL REPAIR
   ============================================================ */

function runFullSystemRepair() {

  try {

    ensureCoreSheets_();
    removeSheet1_();

    autoFormatSheets_();

    applyVerificationDropdown();

    installAutomaticTriggers();

    /*
     * Do not automatically delete user data.
     */

    syncAllFormsSilent_();

    verifyAllSilent_();


    info_(
      'runFullSystemRepair',
      'Full system repair completed',
      'No user data was deleted.'
    );


    SpreadsheetApp.getUi().alert(
      'Full System Repair',
      'Completed successfully.\n\n' +
      'Existing user data was preserved.\n' +
      'Forms were synchronized.\n' +
      'Verification was refreshed.\n' +
      'Dropdown was applied.\n' +
      'Automatic triggers were repaired.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );


  } catch (error) {

    error_(
      'runFullSystemRepair',
      'Full repair failed',
      error.stack || error.message
    );


    SpreadsheetApp.getUi().alert(
      'Repair Error',
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/* ============================================================
   33. FORMATTING
   ============================================================ */

function autoFormatSheets_() {

  try {

    const names = [
      CONFIG.SHEETS.REGISTRY,
      CONFIG.SHEETS.MASTER,
      CONFIG.SHEETS.VERIFICATION,
      CONFIG.SHEETS.LOG
    ];


    names.forEach(function(name) {

      const sheet =
        getSheet_(name);

      const lastColumn =
        sheet.getLastColumn();

      if (lastColumn > 0) {

        formatHeader_(
          sheet,
          lastColumn
        );

        sheet.autoResizeColumns(
          1,
          Math.min(
            lastColumn,
            30
          )
        );
      }

    });


  } catch (error) {

    warning_(
      'autoFormatSheets_',
      'Formatting partially failed',
      error.message
    );
  }
}


/* ============================================================
   34. DIAGNOSTICS
   ============================================================ */

function runDiagnostics() {

  const output = [];


  try {

    output.push(
      'Spreadsheet: ' +
      getSpreadsheet_().getName()
    );


    output.push(
      'User: ' +
      Session.getEffectiveUser().getEmail()
    );


    output.push(
      'Registry sheet: ' +
      (
        getSpreadsheet_()
          .getSheetByName(
            CONFIG.SHEETS.REGISTRY
          )
          ? 'OK'
          : 'MISSING'
      )
    );


    output.push(
      'Master sheet: ' +
      (
        getSpreadsheet_()
          .getSheetByName(
            CONFIG.SHEETS.MASTER
          )
          ? 'OK'
          : 'MISSING'
      )
    );


    output.push(
      'Verification sheet: ' +
      (
        getSpreadsheet_()
          .getSheetByName(
            CONFIG.SHEETS.VERIFICATION
          )
          ? 'OK'
          : 'MISSING'
      )
    );


    output.push(
      'System log: ' +
      (
        getSpreadsheet_()
          .getSheetByName(
            CONFIG.SHEETS.LOG
          )
          ? 'OK'
          : 'MISSING'
      )
    );


    const registry =
      getSheet_(
        CONFIG.SHEETS.REGISTRY
      );


    if (registry.getLastRow() >= 2) {

      const headers =
        getHeaders_(registry);

      const formIdCol =
        getColumnIndex_(
          headers,
          ['Form ID']
        );


      for (
        let r = 2;
        r <= registry.getLastRow();
        r++
      ) {

        const formId =
          formIdCol
            ? safeString_(
                registry
                  .getRange(
                    r,
                    formIdCol
                  )
                  .getValue()
              )
            : '';


        if (!formId) {
          continue;
        }


        /*
         * Form access test.
         */

        let formAccess =
          'FAILED';

        try {

          const form =
            FormApp.openById(
              formId
            );

          formAccess =
            'SUCCESS';

          output.push(
            'Form ID: ' +
            formId +
            ' | FormApp Access: ' +
            formAccess +
            ' | Title: ' +
            safeString_(
              form.getTitle()
            )
          );

        } catch (error) {

          output.push(
            'Form ID: ' +
            formId +
            ' | FormApp Access: FAILED | ' +
            error.message
          );
        }


        /*
         * Drive access test.
         */

        try {

          const file =
            DriveApp.getFileById(
              formId
            );

          output.push(
            'Drive Access: SUCCESS | ' +
            file.getName()
          );

        } catch (error) {

          output.push(
            'Drive Access: FAILED | ' +
            error.message
          );
        }

      }

    }


    info_(
      'runDiagnostics',
      'Diagnostics completed',
      output.join('\n')
    );


    SpreadsheetApp.getUi().alert(
      'System Diagnostics',
      output.join('\n\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );


  } catch (error) {

    error_(
      'runDiagnostics',
      'Diagnostics failed',
      error.stack || error.message
    );


    SpreadsheetApp.getUi().alert(
      'Diagnostics Error',
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


/* ============================================================
   35. CLEAR OLD LOGS
   ============================================================ */

function clearOldLogs() {

  const ui =
    SpreadsheetApp.getUi();


  const answer =
    ui.alert(
      'Clear Logs',
      'Delete existing SYSTEM_LOG records?',
      ui.ButtonSet.YES_NO
    );


  if (
    answer !== ui.Button.YES
  ) {
    return;
  }


  const sheet =
    getSheet_(
      CONFIG.SHEETS.LOG
    );


  if (sheet.getLastRow() > 1) {

    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        sheet.getLastColumn()
      )
      .clearContent();
  }


  info_(
    'clearOldLogs',
    'Old logs cleared'
  );


  ui.alert(
    'Logs',
    'Old logs cleared.',
    ui.ButtonSet.OK
  );
}


/* ============================================================
   36. TRIGGER-SAFE ERROR HANDLER
   ============================================================ */

function onEdit(e) {

  /*
   * Deliberately lightweight.
   *
   * No dependency on undefined functions.
   */

  try {

    if (!e || !e.range) {
      return;
    }


    const sheet =
      e.range.getSheet();


    if (
      sheet.getName() ===
      CONFIG.SHEETS.VERIFICATION
    ) {

      const headers =
        getHeaders_(sheet);


      const statusColumn =
        getColumnIndex_(
          headers,
          ['Verification Status']
        );


      if (
        statusColumn &&
        e.range.getColumn() ===
        statusColumn
      ) {

        /*
         * User manually changed status.
         * We intentionally do not overwrite it.
         */

        info_(
          'onEdit',
          'Manual verification status change',
          e.range.getA1Notation() +
          ' = ' +
          e.value
        );
      }

    }

  } catch (error) {

    /*
     * NEVER break spreadsheet edit.
     */

    console.log(
      'onEdit error: ' +
      error.message
    );
  }
}
