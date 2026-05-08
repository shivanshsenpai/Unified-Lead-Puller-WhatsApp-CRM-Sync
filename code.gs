/**
 * UNIFIED LEAD PULLER + WHATSAPP CRM + CONTACT SYNC
 * Flow:
 * 1) Pull leads into Customer sheet
 * 2) Reuse the same processing logic as onEdit (categorize, naming, group-200 logic)
 * 3) Append to Processed_data
 * 4) Sync to Google Contacts (background/manual)
 * 5) Export WhatsApp CSV by category/group
 */

// ==========================================
// 1. GLOBAL SETTINGS
// ==========================================

const S_CUSTOMER = 'customer';
const S_PROCESSED = 'Processed_data';
const S_CONFIG = 'Config';
const S_LOGS = 'Logs';

const TRIGGER_HANDLE_ON_EDIT = 'handleOnEdit';
const TRIGGER_BACKGROUND_SYNC = 'processBackgroundSync';
const TRIGGER_IMPORT = 'fetchLeadBatch';

const IMPORT_PROP_START_DATE = 'START_DATE';
const IMPORT_PROP_END_DATE = 'END_DATE';
const IMPORT_PROP_PAGE = 'PAGE';
const IMPORT_PROP_TOTAL_IMPORTED = 'TOTAL_IMPORTED';
const IMPORT_PROP_STOP_REQUESTED = 'IMPORT_STOP_REQUESTED';
const IMPORT_PROP_LAST_STATUS = 'IMPORT_LAST_STATUS';
const IMPORT_PROP_EMPTY_RETRY_COUNT = 'EMPTY_RETRY_COUNT';
const IMPORT_PROP_LAST_API_HIT_MS = 'LAST_API_HIT_MS';
const IMPORT_PROP_INTERVAL_DAYS = 'IMPORT_INTERVAL_DAYS';
const IMPORT_PROP_LAST_RUN_DATE = 'IMPORT_LAST_RUN_DATE';
const IMPORT_PROP_RANGE_END_DATE = 'IMPORT_RANGE_END_DATE';
const IMPORT_PROP_IS_RECURRING = 'IMPORT_IS_RECURRING';

// --- SENSITIVE CREDENTIALS & CONSTANTS ---
const API_KEY = 'YOUR_API_KEY_HERE'; 
const EXPORT_FOLDER_NAME = 'WhatsApp_Exports';
// -----------------------------------------

const IMPORT_EMPTY_MAX_RETRIES = 10;
const IMPORT_RATE_LIMIT_WINDOW_MS = (5 * 60 * 1000) + 15000; // 5m + safety buffer

// ==========================================
// 2. MENU + TRIGGERS
// ==========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Lead Puller + CRM')
    .addItem('Lead API: Open Puller', 'showApiSidebar')
    .addItem('Lead API: Start Import (Saved Dates)', 'startLeadImport')
    .addItem('Lead API: Setup Weekly Recurring Import', 'showRecurringImportSetup')
    .addItem('Lead API: View Recurring Status', 'showRecurringStatus')
    .addItem('Lead API: Stop Import', 'stopApiImport')
    .addSeparator()
    .addItem('Export Selected to WhatsApp CSVs', 'showCsvPopup')
    .addItem('Manual Sync to Google Contacts', 'showSyncPopup')
    .addSeparator()
    .addItem('Initialize CRM Automation', 'setupTriggers')
    .addToUi();
}

function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  clearTriggersByHandlers_([TRIGGER_HANDLE_ON_EDIT, TRIGGER_BACKGROUND_SYNC]);

  ScriptApp.newTrigger(TRIGGER_HANDLE_ON_EDIT)
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger(TRIGGER_BACKGROUND_SYNC)
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getUi().alert(
    'Automation Active!\n\n1. Customer edits/import are auto-processed.\n2. Contact sync runs every 5 minutes.'
  );
}

function clearTriggersByHandlers_(handlerNames) {
  const allowed = {};
  for (let i = 0; i < handlerNames.length; i++) {
    allowed[handlerNames[i]] = true;
  }

  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (allowed[fn]) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ==========================================
// 3. CUSTOMER -> PROCESSED ENGINE (onEdit + Import)
// ==========================================

function handleOnEdit(e) {
  if (!e || !e.range || !e.source) return;

  const range = e.range;
  const sheet = range.getSheet();
  if (!isCustomerSheet_(sheet)) return;
  if (range.getRow() === 1) return;

  const editedStartCol = range.getColumn();
  const editedEndCol = range.getLastColumn();
  if (editedStartCol > 5 || editedEndCol < 1) return;

  processCustomerRows_(sheet, range.getRow(), range.getNumRows());
}

function processCustomerRows_(customerSheet, startRow, numRows) {
  if (!customerSheet || numRows <= 0) return 0;

  const ss = customerSheet.getParent();
  const procSheet = getOrCreateSheetCaseInsensitive_(ss, S_PROCESSED);
  const configSheet = getSheetCaseInsensitive_(ss, S_CONFIG);

  ensureProcessedHeader_(procSheet);

  const rules = getCategoryRules_(configSheet);
  const existingPhones = getExistingProcessedPhoneSet_(procSheet);
  const fallbackCounters = {};

  const values = customerSheet.getRange(startRow, 1, numRows, 5).getValues();
  const pendingRows = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const custName = safeCellString_(row[0]);
    const city = safeCellString_(row[1]);
    const orderTime = row[2];
    const prodName = safeCellString_(row[3]);
    const rawPhone = safeCellString_(row[4]);

    if (!custName || !prodName || !rawPhone) continue;
    if (custName === '-' || prodName === '-' || rawPhone === '-') continue;

    const cleanPhone = normalizePhone_(rawPhone);
    if (!cleanPhone) continue;
    if (existingPhones[cleanPhone]) continue;

    const categoryMatch = matchCategory_(prodName, rules);
    if (categoryMatch.category === 'EXCLUDED') continue;

    const counter = getNextGroupCounter_(configSheet, categoryMatch, fallbackCounters);

    const categoryInitial = categoryMatch.initial || 'OTH';
    const cleanName = sanitizeNameForBroadcast_(custName);
    const broadcastName = categoryInitial + '_G' + counter.groupNum + '_C' + counter.recCount + '_' + cleanName;

    pendingRows.push([
      custName,
      city,
      orderTime,
      prodName,
      toDisplayPhone_(cleanPhone),
      categoryMatch.category,
      broadcastName,
      cleanPhone,
      'Pending'
    ]);

    existingPhones[cleanPhone] = true;
  }

  if (pendingRows.length > 0) {
    const appendStartRow = procSheet.getLastRow() + 1;
    procSheet.getRange(appendStartRow, 1, pendingRows.length, 9).setValues(pendingRows);

    try {
      syncProcessedRowsImmediate_(procSheet, appendStartRow, pendingRows.length);
    } catch (e) {
      writeLog_('ERROR', 'Immediate sync failed: ' + e.message);
    }

    try {
      autoExportCsvForProcessedRows_(procSheet, appendStartRow, pendingRows.length);
    } catch (e) {
      writeLog_('ERROR', 'Auto CSV export failed: ' + e.message);
    }
  }

  return pendingRows.length;
}

function isCustomerSheet_(sheet) {
  if (!sheet) return false;
  return safeCellString_(sheet.getName()).toLowerCase() === S_CUSTOMER;
}

function ensureProcessedHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;

  sheet.appendRow([
    'Customer_Name',
    'City',
    'Order_Time',
    'Product_Name',
    'Phone_Number',
    'category',
    'Broadcast_Name',
    'WhatsApp_Number',
    'Synced_to_Contacts'
  ]);
}

function getCategoryRules_(configSheet) {
  if (!configSheet || configSheet.getLastRow() < 2) return [];

  const data = configSheet.getDataRange().getValues();
  const rules = [];

  for (let i = 1; i < data.length; i++) {
    const catName = safeCellString_(data[i][0]);
    if (!catName) continue;

    const initialRaw = safeCellString_(data[i][1]);
    const initial = (initialRaw || catName.substring(0, 3)).toUpperCase().replace(/\s+/g, '');
    const keywordsRaw = safeCellString_(data[i][2]).toLowerCase();

    if (!keywordsRaw) continue;

    const keywords = keywordsRaw
      .split('|')
      .map(function(k) { return k.trim(); })
      .filter(function(k) { return !!k; });

    if (!keywords.length) continue;

    const escaped = keywords.map(function(k) {
      return k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    });

    const pattern = '\\b(' + escaped.join('|') + ')\\b';

    rules.push({
      category: catName,
      initial: initial,
      regex: new RegExp(pattern, 'i'),
      configRow: i + 1
    });
  }

  return rules;
}

function matchCategory_(productName, rules) {
  const defaultCategory = {
    category: 'Others',
    initial: 'OTH',
    configRow: -1
  };

  if (!productName || !rules || !rules.length) return defaultCategory;

  for (let i = 0; i < rules.length; i++) {
    if (rules[i].regex.test(productName)) {
      if (safeCellString_(rules[i].category).toLowerCase() === 'exclude') {
        return { category: 'EXCLUDED', initial: 'EXC', configRow: rules[i].configRow };
      }
      return {
        category: rules[i].category,
        initial: rules[i].initial,
        configRow: rules[i].configRow
      };
    }
  }

  return defaultCategory;
}

function getNextGroupCounter_(configSheet, categoryMatch, fallbackCounters) {
  if (configSheet && categoryMatch.configRow > 0) {
    let groupNum = parseInt(configSheet.getRange(categoryMatch.configRow, 4).getValue(), 10) || 1;
    let recCount = parseInt(configSheet.getRange(categoryMatch.configRow, 5).getValue(), 10) || 0;

    if (recCount >= 200) {
      groupNum++;
      recCount = 0;
    }

    recCount++;

    configSheet.getRange(categoryMatch.configRow, 4, 1, 2).setValues([[groupNum, recCount]]);

    return { groupNum: groupNum, recCount: recCount };
  }

  const key = categoryMatch.initial || 'OTH';
  if (!fallbackCounters[key]) {
    fallbackCounters[key] = { groupNum: 1, recCount: 0 };
  }

  if (fallbackCounters[key].recCount >= 200) {
    fallbackCounters[key].groupNum++;
    fallbackCounters[key].recCount = 0;
  }

  fallbackCounters[key].recCount++;

  return {
    groupNum: fallbackCounters[key].groupNum,
    recCount: fallbackCounters[key].recCount
  };
}

function getExistingProcessedPhoneSet_(procSheet) {
  const result = {};
  if (!procSheet || procSheet.getLastRow() < 2) return result;

  const rowCount = procSheet.getLastRow() - 1;
  const rawPhoneCol = procSheet.getRange(2, 5, rowCount, 1).getValues();
  const waPhoneCol = procSheet.getRange(2, 8, rowCount, 1).getValues();

  for (let i = 0; i < rowCount; i++) {
    const p1 = normalizePhone_(rawPhoneCol[i][0]);
    const p2 = normalizePhone_(waPhoneCol[i][0]);
    if (p1) result[p1] = true;
    if (p2) result[p2] = true;
  }

  return result;
}

function normalizePhone_(raw) {
  const digitsOnly = safeCellString_(raw).replace(/\D/g, '');
  if (!digitsOnly) return '';

  let digits = digitsOnly;

  if (digits.length === 12 && digits.indexOf('91') === 0) {
    digits = digits.substring(2);
  }

  if (digits.length > 10) {
    digits = digits.slice(-10);
  }

  if (digits.length !== 10) return '';

  return digits;
}

function toDisplayPhone_(tenDigitPhone) {
  if (!tenDigitPhone) return '';
  return '+91' + tenDigitPhone;
}

function toE164Format_(rawPhone) {
  const ten = normalizePhone_(rawPhone);
  if (!ten) return '';
  return '+91' + ten;
}

function sanitizeNameForBroadcast_(name) {
  const clean = safeCellString_(name).replace(/\s+/g, '_').replace(/[^A-Za-z0-9_]/g, '');
  return clean || 'NoName';
}

function safeCellString_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

// ==========================================
// 4. API IMPORT (INTEGRATED)
// ==========================================

function showSidebar() {
  showApiSidebar();
}

function showApiSidebar() {
  const html = HtmlService
    .createHtmlOutput(`
      <div style="font-family: Arial; padding: 20px;">
        <h2 style="margin-top:0;">API Lead Puller</h2>

        <label>From Date</label><br><br>
        <input type="date" id="startDate" style="width:100%; padding:10px; font-size:14px; box-sizing:border-box;">

        <br><br><br>

        <label>To Date</label><br><br>
        <input type="date" id="endDate" style="width:100%; padding:10px; font-size:14px; box-sizing:border-box;">

        <br><br><br>

        <button onclick="startImport()" style="width:100%; padding:12px; background:#4285f4; color:white; border:none; font-size:15px; cursor:pointer; border-radius:5px;">
          Pull Leads
        </button>
        <br><br>
        <button onclick="stopImport()" style="width:100%; padding:12px; background:#d93025; color:white; border:none; font-size:15px; cursor:pointer; border-radius:5px;">
          Stop Import
        </button>

        <div id="status" style="margin-top:20px; color:green; font-size:14px;"></div>

        <script>
          google.script.run.withSuccessHandler(function(msg) {
            document.getElementById("status").innerHTML = msg;
          }).getImportStatusMessage();

          function formatDate(dateString) {
            const parts = dateString.split("-");
            const date = new Date(
              Number(parts[0]),
              Number(parts[1]) - 1,
              Number(parts[2])
            );
            const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
            const day = String(date.getDate()).padStart(2, '0');
            const month = months[date.getMonth()];
            const year = date.getFullYear();
            return day + "-" + month + "-" + year;
          }

          function startImport() {
            const rawStart = document.getElementById("startDate").value;
            const rawEnd = document.getElementById("endDate").value;

            if (!rawStart || !rawEnd) {
              document.getElementById("status").innerHTML = "<span style='color:red;'>Please select both dates.</span>";
              return;
            }
            if (rawStart > rawEnd) {
              document.getElementById("status").innerHTML = "<span style='color:red;'>From Date cannot be after To Date.</span>";
              return;
            }
            const startObj = new Date(rawStart + "T00:00:00");
            const endObj = new Date(rawEnd + "T00:00:00");
            const diffDays = Math.floor((endObj - startObj) / (24 * 60 * 60 * 1000));
            if (diffDays > 6) {
              document.getElementById("status").innerHTML = "<span style='color:red;'>API allows max 7 days per request. Please select a smaller range.</span>";
              return;
            }

            const startDate = formatDate(rawStart);
            const endDate = formatDate(rawEnd);

            document.getElementById("status").innerHTML = "Starting import...";

            google.script.run
              .withSuccessHandler(function(msg) {
                document.getElementById("status").innerHTML = msg;
              })
              .withFailureHandler(function(err) {
                document.getElementById("status").innerHTML = "<span style='color:red;'>" + err.message + "</span>";
              })
              .initializeLeadImport(startDate, endDate);
          }

          function stopImport() {
            document.getElementById("status").innerHTML = "Stopping import...";
            google.script.run
              .withSuccessHandler(function(msg) {
                document.getElementById("status").innerHTML = msg;
              })
              .withFailureHandler(function(err) {
                document.getElementById("status").innerHTML = "<span style='color:red;'>" + err.message + "</span>";
              })
              .stopApiImportFromUi();
          }
        </script>
      </div>
    `)
    .setTitle('API Lead Manager')
    .setWidth(320);

  SpreadsheetApp.getUi().showSidebar(html);
}

function initializeLeadImport(startDate, endDate) {
  stopApiImport(false);

  const props = PropertiesService.getScriptProperties();
  props.setProperty(IMPORT_PROP_START_DATE, startDate);
  props.setProperty(IMPORT_PROP_END_DATE, endDate);
  props.setProperty(IMPORT_PROP_PAGE, '1');
  props.setProperty(IMPORT_PROP_TOTAL_IMPORTED, '0');
  props.setProperty(IMPORT_PROP_STOP_REQUESTED, '0');
  props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
  props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, '0');
  props.setProperty(IMPORT_PROP_LAST_STATUS, 'Import queued from page 1.');

  createApiImportTrigger_(1000);

  return 'Import started in background. Leads will be pushed to customer sheet and processed automatically.';
}

function startLeadImport() {
  const props = PropertiesService.getScriptProperties();
  const startDate = props.getProperty(IMPORT_PROP_START_DATE);
  const endDate = props.getProperty(IMPORT_PROP_END_DATE);

  if (!startDate || !endDate) {
    SpreadsheetApp.getUi().alert('Please open "Lead API: Open Puller" and choose date range first.');
    return;
  }

  stopApiImport(false);
  props.setProperty(IMPORT_PROP_STOP_REQUESTED, '0');
  props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
  props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, '0');
  props.setProperty(IMPORT_PROP_LAST_STATUS, 'Import restarted from saved dates.');
  createApiImportTrigger_(1000);
  SpreadsheetApp.getUi().alert('API import restarted with saved date range.');
}

function showRecurringImportSetup() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: Arial; padding: 20px; }
          .container { max-width: 400px; margin: 0 auto; }
          h2 { color: #1f73e6; margin-top: 0; }
          .form-group { margin-bottom: 15px; }
          label { display: block; margin-bottom: 5px; font-weight: bold; }
          input { width: 100%; padding: 8px; box-sizing: border-box; font-size: 14px; }
          button { width: 100%; padding: 10px; background: #25D366; color: white; border: none; cursor: pointer; font-weight: bold; border-radius: 4px; }
          button:hover { background: #20a756; }
          .info { background: #f0f0f0; padding: 10px; border-left: 4px solid #1f73e6; margin: 10px 0; font-size: 12px; }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>Weekly Recurring Import</h2>
          
          <div class="info">
            <strong>How it works:</strong><br>
            • Sets up automatic weekly data pulls<br>
            • Loads COMPLETE week of data before moving to next<br>
            • Ensures no partial data is left behind
          </div>

          <div class="form-group">
            <label for="interval">Import Interval (Days):</label>
            <input type="number" id="interval" min="1" max="30" value="7" placeholder="e.g., 7 for weekly">
          </div>

          <button onclick="startRecurring()">Start Recurring Import</button>
          
          <script>
            function startRecurring() {
              const intervalInput = document.getElementById('interval').value;
              const interval = parseFloat(intervalInput);
              
              if (!intervalInput || isNaN(interval)) {
                alert('❌ Enter a valid number');
                return;
              }
              if (interval < 1 || interval > 365) {
                alert('❌ Invalid interval range');
                return;
              }
              
              google.script.run
                .withSuccessHandler(function() {
                  document.body.innerHTML = '<p style="text-align:center; padding:20px;"><strong>✓ Recurring import started!</strong><br>Trigger created. Check logs for progress.</p>';
                })
                .withFailureHandler(function(err) {
                  alert('❌ Error: ' + err.message);
                })
                .startRecurringImportWithInterval(Math.round(interval));
            }
          </script>
        </div>
      </body>
    </html>
  `);

  SpreadsheetApp.getUi().showModelessDialog(html, 'Setup Weekly Import');
}

function startRecurringImportWithInterval(intervalDays) {
  const validatedDays = parseInt(intervalDays, 10);
  if (!validatedDays || isNaN(validatedDays) || validatedDays < 1 || validatedDays > 365) {
    throw new Error('Invalid interval');
  }

  stopApiImport(false);

  const props = PropertiesService.getScriptProperties();
  const now = new Date();

  props.setProperty(IMPORT_PROP_INTERVAL_DAYS, String(validatedDays));
  props.setProperty(IMPORT_PROP_IS_RECURRING, '1');
  props.setProperty(IMPORT_PROP_LAST_RUN_DATE, formatDateForImport_(now));

  const endDate = formatDateForImport_(now);
  const startDate = formatDateForImport_(new Date(now.getTime() - (validatedDays * 24 * 60 * 60 * 1000)));

  props.setProperty(IMPORT_PROP_START_DATE, startDate);
  props.setProperty(IMPORT_PROP_END_DATE, endDate);
  props.setProperty(IMPORT_PROP_RANGE_END_DATE, endDate);
  props.setProperty(IMPORT_PROP_PAGE, '1');
  props.setProperty(IMPORT_PROP_TOTAL_IMPORTED, '0');
  props.setProperty(IMPORT_PROP_STOP_REQUESTED, '0');
  props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
  props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, '0');
  props.setProperty(IMPORT_PROP_LAST_STATUS, 'Recurring import started. Loading ' + validatedDays + ' days of data.');

  writeLog_('INFO', 'RECURRING IMPORT STARTED: ' + validatedDays + ' day interval. Range: ' + startDate + ' to ' + endDate);

  createApiImportTrigger_(1000);

  SpreadsheetApp.getUi().alert('Recurring import activated!\nInterval: ' + validatedDays + ' days.');
}

function formatDateForImport_(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return yyyy + '-' + mm + '-' + dd;
}

function showRecurringStatus() {
  const props = PropertiesService.getScriptProperties();
  const isRecurring = props.getProperty(IMPORT_PROP_IS_RECURRING) === '1';
  
  if (!isRecurring) {
    SpreadsheetApp.getUi().alert('No recurring import is currently active.');
    return;
  }

  const intervalDays = props.getProperty(IMPORT_PROP_INTERVAL_DAYS);
  const startDate = props.getProperty(IMPORT_PROP_START_DATE);
  const endDate = props.getProperty(IMPORT_PROP_END_DATE);
  const totalImported = props.getProperty(IMPORT_PROP_TOTAL_IMPORTED);
  const lastStatus = props.getProperty(IMPORT_PROP_LAST_STATUS);

  const message = 
    'RECURRING IMPORT STATUS\n\n' +
    'Interval: ' + intervalDays + ' days\n' +
    'Current Range: ' + startDate + ' to ' + endDate + '\n' +
    'Total Imported: ' + (totalImported || '0') + ' leads\n\n' +
    'Status: ' + (lastStatus || 'Running...');

  SpreadsheetApp.getUi().alert(message);
}

function fetchLeadBatch() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1)) {
    writeLog_('INFO', 'Skipped fetch because another import execution is already running.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  const startDate = props.getProperty(IMPORT_PROP_START_DATE);
  const endDate = props.getProperty(IMPORT_PROP_END_DATE);
  let page = parseInt(props.getProperty(IMPORT_PROP_PAGE), 10) || 1;
  let totalImported = parseInt(props.getProperty(IMPORT_PROP_TOTAL_IMPORTED), 10) || 0;
  let emptyRetryCount = parseInt(props.getProperty(IMPORT_PROP_EMPTY_RETRY_COUNT), 10) || 0;
  const lastApiHitMs = parseInt(props.getProperty(IMPORT_PROP_LAST_API_HIT_MS), 10) || 0;
  const nowMs = new Date().getTime();

  if (isImportStopRequested_()) {
    writeLog_('INFO', 'Import stop was requested. No further pages scheduled.');
    return;
  }

  if (!startDate || !endDate) {
    writeLog_('ERROR', 'Import aborted. Missing start/end date.');
    props.setProperty(IMPORT_PROP_LAST_STATUS, 'Import aborted: missing dates.');
    stopApiImport(false);
    return;
  }

  if (lastApiHitMs > 0) {
    const elapsedMs = nowMs - lastApiHitMs;
    if (elapsedMs < IMPORT_RATE_LIMIT_WINDOW_MS) {
      const waitMs = IMPORT_RATE_LIMIT_WINDOW_MS - elapsedMs;
      const waitSec = Math.ceil(waitMs / 1000);
      const waitMsg = 'Waiting ' + waitSec + 's before next API hit to respect limit.';
      writeLog_('INFO', waitMsg);
      props.setProperty(IMPORT_PROP_LAST_STATUS, waitMsg);
      createApiImportTrigger_(waitMs);
      return;
    }
  }

  const customerSheet = getOrCreateSheetCaseInsensitive_(ss, S_CUSTOMER);
  ensureCustomerHeader_(customerSheet);

  writeLog_('INFO', 'Starting API page ' + page + ' | Range: ' + startDate + ' to ' + endDate);

  // EXAMPLE API ENDPOINT STRUCTURE - Adjust according to your provider
  const url =
    'https://mapi.indiamart.com/wservce/crm/crmListing/v2/' +
    '?glusr_crm_key=' + encodeURIComponent(API_KEY) +
    '&start_time=' + encodeURIComponent(startDate) +
    '&end_time=' + encodeURIComponent(endDate) +
    '&page=' + encodeURIComponent(page);

  const options = {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0',
      'Accept': 'application/json'
    }
  };

  try {
    props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, String(new Date().getTime()));
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const text = response.getContentText();

    writeLog_('API', 'HTTP ' + responseCode + ' | Page ' + page);

    if (responseCode !== 200) {
      writeLog_('ERROR', text.substring(0, 500));
      if (responseCode === 429) {
        const extMsg = 'Rate-limit hit (HTTP 429). Waiting before retry.';
        writeLog_('INFO', extMsg);
        props.setProperty(IMPORT_PROP_LAST_STATUS, extMsg);
      }
      const delayMs = responseCode === 429 ? IMPORT_RATE_LIMIT_WINDOW_MS : 30000;
      createApiImportTrigger_(delayMs);
      return;
    }

    const json = JSON.parse(text);

    if (!json.RESPONSE || json.RESPONSE.length === 0) {
      const apiMessage = safeCellString_(json.MESSAGE || json.message || json.STATUS || '');
      const rateLimitDetected = isRateLimitMessage_(apiMessage);
      const rangeLimitDetected = isDateRangeLimitMessage_(apiMessage);

      if (rangeLimitDetected) {
        const limitMsg = 'Invalid date range. Import stopped.';
        writeLog_('ERROR', limitMsg + ' API message: ' + apiMessage);
        props.setProperty(IMPORT_PROP_LAST_STATUS, limitMsg);
        stopApiImport(false);
        return;
      }

      if (emptyRetryCount < IMPORT_EMPTY_MAX_RETRIES) {
        emptyRetryCount++;
        props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, String(emptyRetryCount));

        const delayMs = rateLimitDetected
          ? getRateLimitRetryDelayMs_(apiMessage)
          : Math.min(120000, 15000 * emptyRetryCount);
        const delaySec = Math.round(delayMs / 1000);
        const retryMsg = 'API returned no RESPONSE rows. Retrying ' + emptyRetryCount + '/' + IMPORT_EMPTY_MAX_RETRIES + ' in ' + delaySec + 's.';

        writeLog_('INFO', retryMsg);
        props.setProperty(IMPORT_PROP_LAST_STATUS, retryMsg);
        createApiImportTrigger_(delayMs);
        return;
      }

      writeLog_('DONE', 'Import completed after retries. Total leads imported: ' + totalImported);
      
      const isRecurring = props.getProperty(IMPORT_PROP_IS_RECURRING) === '1';
      const intervalDays = parseInt(props.getProperty(IMPORT_PROP_INTERVAL_DAYS), 10) || 0;

      if (isRecurring && intervalDays > 0) {
        const nextStart = new Date(new Date(endDate).getTime() + (24 * 60 * 60 * 1000));
        const nextEnd = new Date(nextStart.getTime() + (intervalDays * 24 * 60 * 60 * 1000));
        
        props.setProperty(IMPORT_PROP_START_DATE, formatDateForImport_(nextStart));
        props.setProperty(IMPORT_PROP_END_DATE, formatDateForImport_(nextEnd));
        props.setProperty(IMPORT_PROP_RANGE_END_DATE, formatDateForImport_(nextEnd));
        props.setProperty(IMPORT_PROP_PAGE, '1');
        props.setProperty(IMPORT_PROP_TOTAL_IMPORTED, '0');
        props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
        props.setProperty(IMPORT_PROP_LAST_API_HIT_MS, '0');
        
        const delayMs = (intervalDays * 24 * 60 * 60 * 1000);
        const nextRangeMsg = 'Current range complete (' + totalImported + ' leads imported). Next batch scheduled for ' + intervalDays + ' days.';
        
        writeLog_('INFO', nextRangeMsg);
        props.setProperty(IMPORT_PROP_LAST_STATUS, nextRangeMsg);
        createApiImportTrigger_(delayMs);
      } else {
        const completeMsg = 'Import completed. Total imported: ' + totalImported + '.';
        props.setProperty(IMPORT_PROP_LAST_STATUS, completeMsg);
        stopApiImport(false);
      }
      return;
    }

    if (emptyRetryCount !== 0) {
      props.setProperty(IMPORT_PROP_EMPTY_RETRY_COUNT, '0');
    }

    let processedCount = 0;
    let insertedCount = 0;

    for (let i = 0; i < json.RESPONSE.length; i++) {
      if (isImportStopRequested_()) {
        writeLog_('INFO', 'Import stopped by user during page ' + page);
        props.setProperty(IMPORT_PROP_LAST_STATUS, 'Import stopped. Total imported: ' + totalImported);
        stopApiImport(false);
        return;
      }

      const lead = json.RESPONSE[i];
      const row = [
        safeCellString_(lead.SENDER_NAME),
        safeCellString_(lead.SENDER_CITY),
        safeCellString_(lead.QUERY_TIME),
        safeCellString_(lead.QUERY_PRODUCT_NAME),
        safeCellString_(lead.SENDER_MOBILE)
      ];

      const insertRow = customerSheet.getLastRow() + 1;
      customerSheet.getRange(insertRow, 1, 1, 5).setValues([row]);

      processedCount += processCustomerRows_(customerSheet, insertRow, 1);
      insertedCount++;
      totalImported++;
      props.setProperty(IMPORT_PROP_TOTAL_IMPORTED, String(totalImported));
    }

    writeLog_('SUCCESS', 'Imported ' + insertedCount + ' leads (page ' + page + ').');
    props.setProperty(IMPORT_PROP_LAST_STATUS, 'Running page ' + (page + 1) + '. Total imported: ' + totalImported + '.');
    props.setProperty(IMPORT_PROP_PAGE, String(page + 1));

    createApiImportTrigger_(IMPORT_RATE_LIMIT_WINDOW_MS);
  } catch (e) {
    writeLog_('ERROR', e.toString());
    props.setProperty(IMPORT_PROP_LAST_STATUS, 'Error: ' + e.toString());
    createApiImportTrigger_(30000);
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
}

function ensureCustomerHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow(['Customer_Name', 'City', 'Order_Time', 'Product_Name', 'Phone_Number']);
}

function createApiImportTrigger_(delayMs) {
  if (isImportStopRequested_()) return;
  const safeDelay = Math.max(1000, parseInt(delayMs, 10) || 1000);
  clearTriggersByHandlers_([TRIGGER_IMPORT]);
  ScriptApp.newTrigger(TRIGGER_IMPORT)
    .timeBased()
    .after(safeDelay)
    .create();
}

function stopApiImport(markStopped) {
  const props = PropertiesService.getScriptProperties();
  const shouldMarkStopped = markStopped !== false;

  if (shouldMarkStopped) {
    props.setProperty(IMPORT_PROP_STOP_REQUESTED, '1');
    props.setProperty(IMPORT_PROP_LAST_STATUS, 'Import stopped by user.');
    props.setProperty(IMPORT_PROP_IS_RECURRING, '0');
  }

  clearTriggersByHandlers_([TRIGGER_IMPORT]);
}

function stopApiImportFromUi() {
  stopApiImport(true);
  return 'Import stopped. No new batches will run.';
}

function isImportStopRequested_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(IMPORT_PROP_STOP_REQUESTED) === '1';
}

function getImportStatusMessage() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(IMPORT_PROP_LAST_STATUS) || 'Idle.';
}

function isRateLimitMessage_(message) {
  const m = safeCellString_(message).toLowerCase();
  if (!m) return false;
  return m.indexOf('once in every 5 minutes') !== -1 || m.indexOf('try again after 5 minutes') !== -1;
}

function isDateRangeLimitMessage_(message) {
  const m = safeCellString_(message).toLowerCase();
  if (!m) return false;
  return m.indexOf('maximum allowed difference') !== -1;
}

function getRateLimitRetryDelayMs_(message) {
  const m = safeCellString_(message).toLowerCase();
  let minutes = 5;
  const match = m.match(/after\s+(\d+)\s+minute/);
  if (match && match[1]) minutes = parseInt(match[1], 10) || 5;
  return (minutes * 60 * 1000) + 15000;
}

function getLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(S_LOGS);

  if (!logSheet) {
    logSheet = ss.insertSheet(S_LOGS);
    logSheet.appendRow(['Timestamp', 'Type', 'Message']);
  }

  return logSheet;
}

function writeLog_(type, message) {
  const logSheet = getLogSheet_();
  logSheet.appendRow([new Date(), type, message]);
}

// ==========================================
// 5. BACKGROUND SYNC ENGINE
// ==========================================

function processBackgroundSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const procSheet = getSheetCaseInsensitive_(ss, S_PROCESSED);
  if (!procSheet) return;

  const data = procSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const startTime = new Date().getTime();
  const maxExecutionTime = 240000;

  const headers = data[0];
  const bNameIdx = headers.indexOf('Broadcast_Name');
  const phoneIdx = headers.indexOf('WhatsApp_Number');
  const syncIdx = headers.indexOf('Synced_to_Contacts');

  if (bNameIdx === -1 || phoneIdx === -1 || syncIdx === -1) return;

  let groupCache = {};
  try {
    const groupsResponse = People.ContactGroups.list({ pageSize: 1000 });
    const existingGroups = groupsResponse.contactGroups || [];

    for (let i = 0; i < existingGroups.length; i++) {
      groupCache[existingGroups[i].name] = existingGroups[i].resourceName;
    }
  } catch (e) {
    console.error('People API Error: ' + e.message);
    return;
  }

  for (let i = 1; i < data.length; i++) {
    if (new Date().getTime() - startTime > maxExecutionTime) {
      console.log('Time limit reached. Pausing until next cycle.');
      break;
    }

    const isSynced = safeCellString_(data[i][syncIdx]);
    if (isSynced !== 'Pending' && isSynced !== 'Syncing...') continue;

    const broadcastName = safeCellString_(data[i][bNameIdx]);
    const phone = toE164Format_(data[i][phoneIdx]);

    if (!broadcastName || !phone) continue;

    const labelName = getContactGroupNameFromBroadcast_(broadcastName);

    let groupId = groupCache[labelName];
    if (!groupId) {
      try {
        const newGroup = People.ContactGroups.create({ contactGroup: { name: labelName } });
        groupId = newGroup.resourceName;
        groupCache[labelName] = groupId;
      } catch (e) {
        procSheet.getRange(i + 1, syncIdx + 1).setValue('Error: ' + e.message);
        continue;
      }
    }

    const contact = {
      names: [{ givenName: broadcastName }],
      phoneNumbers: [{ value: phone, type: 'mobile' }],
      memberships: [{ contactGroupMembership: { contactGroupResourceName: groupId } }]
    };

    let success = false;
    let lastError = '';

    for (let retry = 0; retry < 3; retry++) {
      try {
        People.People.createContact(contact);
        success = true;
        break;
      } catch (e) {
        lastError = e.message;
        const lowerErr = safeCellString_(lastError).toLowerCase();

        if (lowerErr.indexOf('invalid') !== -1 || lowerErr.indexOf('bad request') !== -1 || lowerErr.indexOf('not found') !== -1) {
          procSheet.getRange(i + 1, syncIdx + 1).setValue('API Error: ' + lastError);
          break;
        }

        Utilities.sleep(2000 * (retry + 1));
      }
    }

    if (success) {
      procSheet.getRange(i + 1, syncIdx + 1).setValue('Yes');
      Utilities.sleep(300);
    } else if (safeCellString_(procSheet.getRange(i + 1, syncIdx + 1).getValue()).indexOf('API Error') === -1) {
      procSheet.getRange(i + 1, syncIdx + 1).setValue('Error: Blocked/Rate Limit');
      break;
    }
  }
}

function getContactGroupNameFromBroadcast_(broadcastName) {
  const parts = safeCellString_(broadcastName).split('_');
  if (parts.length >= 2) return parts[0] + '_' + parts[1];
  return 'Unknown_Group';
}

function syncProcessedRowsImmediate_(procSheet, startRow, numRows) {
  if (!procSheet || numRows <= 0) return;

  const headers = procSheet.getRange(1, 1, 1, procSheet.getLastColumn()).getValues()[0];
  const bNameIdx = headers.indexOf('Broadcast_Name');
  const phoneIdx = headers.indexOf('WhatsApp_Number');
  const syncIdx = headers.indexOf('Synced_to_Contacts');

  if (bNameIdx === -1 || phoneIdx === -1 || syncIdx === -1) return;

  const width = Math.max(bNameIdx, Math.max(phoneIdx, syncIdx)) + 1;
  const rows = procSheet.getRange(startRow, 1, numRows, width).getValues();

  const groupCache = getContactGroupCache_();
  const statusUpdates = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const currentStatus = safeCellString_(row[syncIdx]);
    if (currentStatus !== 'Pending' && currentStatus !== 'Syncing...') continue;

    const broadcastName = safeCellString_(row[bNameIdx]);
    const phone = toE164Format_(row[phoneIdx]);

    if (!broadcastName || !phone) continue;

    const labelName = getContactGroupNameFromBroadcast_(broadcastName);

    let groupId = groupCache[labelName];
    if (!groupId) {
      try {
        const newGroup = People.ContactGroups.create({ contactGroup: { name: labelName } });
        groupId = newGroup.resourceName;
        groupCache[labelName] = groupId;
      } catch (e) {
        statusUpdates.push({ row: startRow + i, value: 'Error: ' + e.message });
        continue;
      }
    }

    const contact = {
      names: [{ givenName: broadcastName }],
      phoneNumbers: [{ value: phone, type: 'mobile' }],
      memberships: [{ contactGroupMembership: { contactGroupResourceName: groupId } }]
    };

    let success = false;
    let lastError = '';

    for (let retry = 0; retry < 3; retry++) {
      try {
        People.People.createContact(contact);
        success = true;
        break;
      } catch (e) {
        lastError = e.message;
        const lowerErr = safeCellString_(lastError).toLowerCase();
        if (lowerErr.indexOf('invalid') !== -1 || lowerErr.indexOf('bad request') !== -1 || lowerErr.indexOf('not found') !== -1) {
          statusUpdates.push({ row: startRow + i, value: 'API Error: ' + lastError });
          break;
        }
        Utilities.sleep(1200 * (retry + 1));
      }
    }

    if (success) {
      statusUpdates.push({ row: startRow + i, value: 'Yes' });
      Utilities.sleep(120);
    } else if (!statusUpdates.some(function(u) { return u.row === startRow + i; })) {
      statusUpdates.push({ row: startRow + i, value: 'Error: Blocked/Rate Limit' });
    }
  }

  for (let j = 0; j < statusUpdates.length; j++) {
    procSheet.getRange(statusUpdates[j].row, syncIdx + 1).setValue(statusUpdates[j].value);
  }
}

function getContactGroupCache_() {
  const cache = {};
  const groupsResponse = People.ContactGroups.list({ pageSize: 1000 });
  const existingGroups = groupsResponse.contactGroups || [];

  for (let i = 0; i < existingGroups.length; i++) {
    cache[existingGroups[i].name] = existingGroups[i].resourceName;
  }

  return cache;
}

function autoExportCsvForProcessedRows_(procSheet, startRow, numRows) {
  if (!procSheet || numRows <= 0) return;

  const headers = procSheet.getRange(1, 1, 1, procSheet.getLastColumn()).getValues()[0];
  const bNameIdx = headers.indexOf('Broadcast_Name');
  const phoneIdx = headers.indexOf('WhatsApp_Number');
  if (bNameIdx === -1 || phoneIdx === -1) return;

  const width = Math.max(bNameIdx, phoneIdx) + 1;
  const newRows = procSheet.getRange(startRow, 1, numRows, width).getValues();
  const targetPrefixes = {};

  for (let i = 0; i < newRows.length; i++) {
    const bName = safeCellString_(newRows[i][bNameIdx]);
    if (!bName) continue;
    targetPrefixes[getContactGroupNameFromBroadcast_(bName)] = true;
  }

  const prefixes = Object.keys(targetPrefixes);
  if (!prefixes.length) return;

  const allData = procSheet.getDataRange().getValues();
  const csvByPrefix = {};

  for (let p = 0; p < prefixes.length; p++) {
    csvByPrefix[prefixes[p]] = ['Name,Phone'];
  }

  for (let r = 1; r < allData.length; r++) {
    const bName = safeCellString_(allData[r][bNameIdx]);
    if (!bName) continue;
    const prefix = getContactGroupNameFromBroadcast_(bName);
    if (!targetPrefixes[prefix]) continue;

    const phone = normalizePhone_(allData[r][phoneIdx]);
    if (!phone) continue;
    csvByPrefix[prefix].push('"' + bName + '","' + phone + '"');
  }

  const folder = getOrCreateFolder_(EXPORT_FOLDER_NAME);

  for (let k = 0; k < prefixes.length; k++) {
    const prefix = prefixes[k];
    const fileName = prefix + '_WhatsApp.csv';
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }
    folder.createFile(Utilities.newBlob(csvByPrefix[prefix].join('\n'), MimeType.CSV, fileName));
  }
}

// ==========================================
// 6. CSV EXPORT LOGIC
// ==========================================

function showCsvPopup() {
  const categories = getAvailableCategories_();
  if (!categories.length) {
    SpreadsheetApp.getUi().alert('No categories found. Add categories in Config sheet or Processed_data first.');
    return;
  }

  const checkboxesHtml = categories.map(function(cat) {
    return '<label class="cb-container"><input type="checkbox" value="' + cat + '" class="cat-checkbox"> ' + cat + '</label>';
  }).join('');

  const html = HtmlService
    .createHtmlOutput(getPopupHtml_('WhatsApp CSV Export', checkboxesHtml, 'runExport()', 'Generate CSV Files'))
    .setWidth(450)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, 'WhatsApp CSV Exporter');
}

function executeWhatsAppExport(selectedCategories) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetCaseInsensitive_(ss, S_PROCESSED);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const bNameIdx = headers.indexOf('Broadcast_Name');
  const phoneIdx = headers.indexOf('WhatsApp_Number');
  const catIdx = headers.indexOf('category');

  if (bNameIdx === -1 || phoneIdx === -1 || catIdx === -1) return;

  const selectedMap = toLookupMap_(selectedCategories);
  const csvGroups = {};

  for (let i = 1; i < data.length; i++) {
    const category = safeCellString_(data[i][catIdx]);
    if (!selectedMap[category]) continue;

    const bName = safeCellString_(data[i][bNameIdx]);
    const phone = normalizePhone_(data[i][phoneIdx]);
    if (!bName || !phone) continue;

    const filePrefix = getContactGroupNameFromBroadcast_(bName);
    if (!csvGroups[filePrefix]) {
      csvGroups[filePrefix] = ['Name,Phone'];
    }

    csvGroups[filePrefix].push('"' + bName + '","' + phone + '"');
  }

  const folder = getOrCreateFolder_(EXPORT_FOLDER_NAME);
  let fileCount = 0;

  for (const prefix in csvGroups) {
    if (!Object.prototype.hasOwnProperty.call(csvGroups, prefix)) continue;

    const fileName = prefix + '_WhatsApp.csv';
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }

    folder.createFile(Utilities.newBlob(csvGroups[prefix].join('\n'), MimeType.CSV, fileName));
    fileCount++;
  }

  SpreadsheetApp.getUi().alert('Export complete! Generated ' + fileCount + ' CSV file(s).');
}

// ==========================================
// 7. MANUAL CONTACT SYNC
// ==========================================

function showSyncPopup() {
  const categories = getAvailableCategories_();
  if (!categories.length) {
    SpreadsheetApp.getUi().alert('No categories found. Add categories in Config sheet or Processed_data first.');
    return;
  }

  const checkboxesHtml = categories.map(function(cat) {
    return '<label class="cb-container"><input type="checkbox" value="' + cat + '" class="cat-checkbox"> ' + cat + '</label>';
  }).join('');

  const html = HtmlService
    .createHtmlOutput(getPopupHtml_('Manual Contacts Sync', checkboxesHtml, 'startSync()', 'Sync Checked Categories'))
    .setWidth(450)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, 'Google Contacts Sync');
}

function executeContactSync(selectedCategories, startIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetCaseInsensitive_(ss, S_PROCESSED);
  if (!sheet) throw new Error("Sheet 'Processed_data' not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { status: 'complete', syncedThisBatch: 0 };

  const headers = data[0];
  const bNameIdx = headers.indexOf('Broadcast_Name');
  const phoneIdx = headers.indexOf('WhatsApp_Number');
  const catIdx = headers.indexOf('category');
  const syncIdx = headers.indexOf('Synced_to_Contacts');

  if (bNameIdx === -1 || phoneIdx === -1 || catIdx === -1 || syncIdx === -1) {
    throw new Error('Processed_data sheet headers are missing required columns.');
  }

  const selectedMap = toLookupMap_(selectedCategories);

  let groupCache = {};
  try {
    const groupsResponse = People.ContactGroups.list({ pageSize: 1000 });
    const existingGroups = groupsResponse.contactGroups || [];
    for (let g = 0; g < existingGroups.length; g++) {
      groupCache[existingGroups[g].name] = existingGroups[g].resourceName;
    }
  } catch (e) {
    throw new Error('Could not load Contact Groups. Error: ' + e.message);
  }

  let syncedThisBatch = 0;
  const startTime = new Date().getTime();
  const start = Math.max(1, parseInt(startIndex, 10) || 1);

  for (let i = start; i < data.length; i++) {
    if (new Date().getTime() - startTime > 240000) {
      return { status: 'partial', nextRow: i, syncedThisBatch: syncedThisBatch };
    }

    const row = data[i];
    const category = safeCellString_(row[catIdx]);
    if (!selectedMap[category]) continue;

    const broadcastName = safeCellString_(row[bNameIdx]);
    const phone = toE164Format_(row[phoneIdx]);
    const isSynced = safeCellString_(row[syncIdx]);

    if (isSynced === 'Yes' || !broadcastName || !phone) continue;

    const labelName = getContactGroupNameFromBroadcast_(broadcastName);

    let groupId = groupCache[labelName];
    if (!groupId) {
      const newGroup = People.ContactGroups.create({ contactGroup: { name: labelName } });
      groupId = newGroup.resourceName;
      groupCache[labelName] = groupId;
    }

    const newContact = {
      names: [{ givenName: broadcastName }],
      phoneNumbers: [{ value: phone, type: 'mobile' }],
      memberships: [{ contactGroupMembership: { contactGroupResourceName: groupId } }]
    };

    let success = false;
    let lastError = '';

    for (let retry = 0; retry < 3; retry++) {
      try {
        People.People.createContact(newContact);
        success = true;
        break;
      } catch (e) {
        lastError = e.message;
        const lowerErr = safeCellString_(lastError).toLowerCase();

        if (lowerErr.indexOf('invalid') !== -1 || lowerErr.indexOf('bad request') !== -1 || lowerErr.indexOf('not found') !== -1) {
          sheet.getRange(i + 1, syncIdx + 1).setValue('API Error: ' + lastError);
          break;
        }

        Utilities.sleep(2000 * (retry + 1));
      }
    }

    if (success) {
      sheet.getRange(i + 1, syncIdx + 1).setValue('Yes');
      Utilities.sleep(500);
      syncedThisBatch++;
    } else if (safeCellString_(sheet.getRange(i + 1, syncIdx + 1).getValue()).indexOf('API Error') === -1) {
      sheet.getRange(i + 1, syncIdx + 1).setValue('Error: Blocked/Rate Limit');
      throw new Error('Google blocked the connection. Error: ' + lastError);
    }
  }

  return { status: 'complete', syncedThisBatch: syncedThisBatch };
}

// ==========================================
// 8. SHARED HELPERS (CSV/UI/CATEGORIES)
// ==========================================

function getAvailableCategories_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = {};

  const configSheet = getSheetCaseInsensitive_(ss, S_CONFIG);
  if (configSheet && configSheet.getLastRow() >= 2) {
    const cfg = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < cfg.length; i++) {
      const cat = safeCellString_(cfg[i][0]);
      if (!cat) continue;
      if (cat.toLowerCase() === 'exclude') continue;
      map[cat] = true;
    }
  }

  const processed = getSheetCaseInsensitive_(ss, S_PROCESSED);
  if (processed && processed.getLastRow() >= 2) {
    const headers = processed.getRange(1, 1, 1, processed.getLastColumn()).getValues()[0];
    const catIdx = headers.indexOf('category');

    if (catIdx !== -1) {
      const cats = processed.getRange(2, catIdx + 1, processed.getLastRow() - 1, 1).getValues();
      for (let j = 0; j < cats.length; j++) {
        const c = safeCellString_(cats[j][0]);
        if (!c) continue;
        if (c.toLowerCase() === 'excluded') continue;
        map[c] = true;
      }
    }
  }

  const out = Object.keys(map);
  out.sort();
  return out;
}

function toLookupMap_(arr) {
  const map = {};
  if (!arr || !arr.length) return map;

  for (let i = 0; i < arr.length; i++) {
    const value = safeCellString_(arr[i]);
    if (value) map[value] = true;
  }

  return map;
}

function getOrCreateFolder_(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getSheetCaseInsensitive_(ss, name) {
  if (!ss || !name) return null;

  const target = safeCellString_(name).toLowerCase();
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    if (safeCellString_(sheets[i].getName()).toLowerCase() === target) {
      return sheets[i];
    }
  }

  return null;
}

function getOrCreateSheetCaseInsensitive_(ss, name) {
  let sheet = getSheetCaseInsensitive_(ss, name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);
  return sheet;
}

// ==========================================
// 9. POPUP HTML GENERATOR
// ==========================================

function getPopupHtml_(title, checkboxes, fnName, btnText) {
  return `
  <!DOCTYPE html>
  <html>
    <head>
      <style>
        body { font-family: sans-serif; padding: 15px; }
        .scroll-box { max-height: 300px; overflow-y: auto; border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; }
        .cb-container { display: block; margin-bottom: 5px; cursor: pointer; }
        button { width: 100%; padding: 10px; background: #25D366; color: white; border: none; cursor: pointer; font-weight: bold; }
        .btn-select { background: #f0f0f0; color: #333; margin-bottom: 10px; font-size: 12px; }
      </style>
    </head>
    <body>
      <h3>${title}</h3>
      <button class="btn-select" onclick="selectAll()">Check/Uncheck All</button>
      <div class="scroll-box">${checkboxes}</div>
      <button id="runBtn" onclick="${fnName}">${btnText}</button>

      <script>
        let allChecked = false;
        let totalSynced = 0;

        function selectAll() {
          allChecked = !allChecked;
          document.querySelectorAll('.cat-checkbox').forEach(function(cb) { cb.checked = allChecked; });
        }

        function runExport() {
          const cats = Array.from(document.querySelectorAll('.cat-checkbox:checked')).map(function(cb) { return cb.value; });
          if (cats.length === 0) return alert('Select at least one category.');
          document.getElementById('runBtn').innerText = 'Processing...';
          google.script.run.withSuccessHandler(function() {
            google.script.host.close();
          }).executeWhatsAppExport(cats);
        }

        function startSync() {
          const cats = Array.from(document.querySelectorAll('.cat-checkbox:checked')).map(function(cb) { return cb.value; });
          if (cats.length === 0) { alert('Select at least one category.'); return; }

          const btn = document.getElementById('runBtn');
          btn.disabled = true;
          btn.style.backgroundColor = '#999';
          totalSynced = 0;

          runSyncBatch(1, cats);
        }

        function runSyncBatch(startRowIndex, cats) {
          const btn = document.getElementById('runBtn');
          if (startRowIndex === 1) btn.innerText = 'Syncing... DO NOT close window';

          google.script.run
            .withSuccessHandler(function(response) {
              totalSynced += response.syncedThisBatch;

              if (response.status === 'partial') {
                btn.innerText = 'Synced ' + totalSynced + ' so far... Continuing';
                runSyncBatch(response.nextRow, cats);
              } else {
                alert('Done! Successfully synced ' + totalSynced + ' total contacts.');
                google.script.host.close();
              }
            })
            .withFailureHandler(function(err) {
              alert('Error: ' + err.message);
              btn.innerText = 'Retry Sync';
              btn.disabled = false;
              btn.style.backgroundColor = '#25D366';
            })
            .executeContactSync(cats, startRowIndex);
        }
      </script>
    </body>
  </html>`;
}
