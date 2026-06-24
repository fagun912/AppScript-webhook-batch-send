// ── CONFIGURATION ─────────────────────────────────────────────────────────────
const CONFIG = {
  LOG_SHEET_NAME: 'WebhookLogs',
  TEMP_SHEET_NAME: '_PabblyTemp',
  START_ROW_KEY: 'lastProcessedRow',
  IS_RUNNING_KEY: 'isRunning',
  SHEET_NAME_KEY: 'sheetName',
  WEBHOOK_URL_KEY: 'webhookUrl',
  INTERVAL_KEY: 'triggerInterval',
  BATCH_SIZE_KEY: 'batchSize'
};

// ── MENU SETUP ────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📤 Webhook Sync')
    .addItem('▶️ Start Processing', 'startProcessing')
    .addItem('⛔ Stop Processing', 'stopProcessing')
    .addItem('📊 Check Status', 'checkStatus')
    .addSeparator()
    .addItem('🧪 Send Test Batch', 'sendTestBatch')
    .addSeparator()
    .addItem('⚙️ Update Settings', 'promptSettings')
    .addToUi();
}

// ── TEMP SHEET ────────────────────────────────────────────────────────────────
function getTempSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.TEMP_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.TEMP_SHEET_NAME);
    sheet.hideSheet();
  }
  return sheet;
}

function writeTempSettings(sheetName, webhookUrl, batchSize, interval) {
  const sheet = getTempSheet();
  sheet.clearContents();
  sheet.getRange('A1:D1').setValues([[sheetName, webhookUrl, batchSize, interval]]);
}

function readAndApplyTempSettings() {
  try {
    const sheet = getTempSheet();
    const values = sheet.getRange('A1:D1').getValues()[0];
    if (!values[0]) return false;
    const props = PropertiesService.getScriptProperties();
    props.setProperty(CONFIG.SHEET_NAME_KEY, values[0].toString());
    props.setProperty(CONFIG.WEBHOOK_URL_KEY, values[1].toString());
    props.setProperty(CONFIG.BATCH_SIZE_KEY, values[2].toString());
    props.setProperty(CONFIG.INTERVAL_KEY, values[3].toString());
    sheet.clearContents();
    return true;
  } catch(e) {
    return false;
  }
}

// ── OPEN SETTINGS DIALOG ──────────────────────────────────────────────────────
function promptSettings() {
  const props = PropertiesService.getScriptProperties();
  const current = {
    sheetName: props.getProperty(CONFIG.SHEET_NAME_KEY) || '',
    webhookUrl: props.getProperty(CONFIG.WEBHOOK_URL_KEY) || '',
    batchSize: props.getProperty(CONFIG.BATCH_SIZE_KEY) || '100',
    interval: props.getProperty(CONFIG.INTERVAL_KEY) || '5'
  };
  const template = HtmlService.createTemplateFromFile('Settings');
  template.current = JSON.stringify(current);
  const html = template.evaluate().setWidth(420).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, '⚙️ Pabbly Sync Settings');
}

function saveSettingsFromDialog(sheetName, webhookUrl, batchSize, interval) {
  writeTempSettings(sheetName, webhookUrl, batchSize, interval);
}

function applyPendingSettings() {
  if (readAndApplyTempSettings()) {
    SpreadsheetApp.getUi().alert('✅ Settings saved successfully!');
  }
}

// ── START PROCESSING ──────────────────────────────────────────────────────────
function startProcessing() {
  const props = PropertiesService.getScriptProperties();
  const ui = SpreadsheetApp.getUi();

  const sheetName = props.getProperty(CONFIG.SHEET_NAME_KEY);
  const webhookUrl = props.getProperty(CONFIG.WEBHOOK_URL_KEY);
  const interval = props.getProperty(CONFIG.INTERVAL_KEY);
  const batchSize = props.getProperty(CONFIG.BATCH_SIZE_KEY);

  if (!sheetName || !webhookUrl || !interval || !batchSize) {
    ui.alert('⚙️ First time setup required!', 'Please configure your settings first.', ui.ButtonSet.OK);
    promptSettings();
    return;
  }

  if (props.getProperty(CONFIG.IS_RUNNING_KEY) === 'true') {
    ui.alert('⚠️ Already running!', 'Use "Stop Processing" first if you want to restart.', ui.ButtonSet.OK);
    return;
  }

  props.setProperty(CONFIG.START_ROW_KEY, '2');
  props.setProperty(CONFIG.IS_RUNNING_KEY, 'true');

  processSingleBatch();
  setupTrigger();

  const totalRows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getLastRow() - 1;
  const totalBatches = Math.ceil(totalRows / parseInt(batchSize));
  const actualInterval = getValidInterval(parseInt(interval));
  const totalMinutes = totalBatches * actualInterval;
  const totalHours = (totalMinutes / 60).toFixed(1);

  ui.alert(
    '✅ Started!',
    `Batch size: ${batchSize} rows\nInterval: every ${actualInterval} min\n\nTotal rows: ${totalRows}\nTotal batches: ${totalBatches}\nEst. total time: ${totalMinutes} min (~${totalHours} hrs)`,
    ui.ButtonSet.OK
  );
}

// ── PROCESS SINGLE BATCH (fired by trigger) ───────────────────────────────────
function processSingleBatch() {
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty(CONFIG.IS_RUNNING_KEY) !== 'true') {
    deleteTrigger();
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = props.getProperty(CONFIG.SHEET_NAME_KEY);
  const batchSize = parseInt(props.getProperty(CONFIG.BATCH_SIZE_KEY) || '100');
  const dataSheet = ss.getSheetByName(sheetName);
  const logSheet = getOrCreateLogSheet(ss);
  const totalRows = dataSheet.getLastRow();
  const totalCols = dataSheet.getLastColumn();
  const headers = dataSheet.getRange(1, 1, 1, totalCols).getValues()[0];
  const currentRow = parseInt(props.getProperty(CONFIG.START_ROW_KEY) || '2');

  if (currentRow > totalRows) {
    logResponse(logSheet, -1, -1, 0, { status: 'COMPLETE', message: 'All rows processed successfully' });
    deleteTrigger();
    props.setProperty(CONFIG.IS_RUNNING_KEY, 'false');
    props.deleteProperty(CONFIG.START_ROW_KEY);
    return;
  }

  const rowsToRead = Math.min(batchSize, totalRows - currentRow + 1);
  const rawValues = dataSheet.getRange(currentRow, 1, rowsToRead, totalCols).getValues();

  const data = rawValues.map((row, i) => {
    const obj = { row_number: currentRow + i };
    headers.forEach((header, j) => {
      obj[header || `column_${j + 1}`] = row[j];
    });
    return obj;
  });

  const result = sendToPabbly(data, currentRow, false);
  logResponse(logSheet, currentRow, currentRow + rowsToRead - 1, rowsToRead, result);
  props.setProperty(CONFIG.START_ROW_KEY, (currentRow + rowsToRead).toString());

  Logger.log(`Batch sent: rows ${currentRow}–${currentRow + rowsToRead - 1} | Status: ${result.status}`);
}

// ── SEND TEST BATCH ───────────────────────────────────────────────────────────
function sendTestBatch() {
  const props = PropertiesService.getScriptProperties();
  const ui = SpreadsheetApp.getUi();

  if (props.getProperty(CONFIG.IS_RUNNING_KEY) === 'true') {
    ui.alert('⚠️ Cannot run test while processing is active.\nPlease stop it first.');
    return;
  }

  const sheetName = props.getProperty(CONFIG.SHEET_NAME_KEY);
  const webhookUrl = props.getProperty(CONFIG.WEBHOOK_URL_KEY);
  const batchSize = parseInt(props.getProperty(CONFIG.BATCH_SIZE_KEY) || '100');

  if (!sheetName || !webhookUrl) {
    ui.alert('⚙️ Settings not configured!', 'Please run Update Settings first.', ui.ButtonSet.OK);
    promptSettings();
    return;
  }

  const confirm = ui.alert(
    '🧪 Send Test Batch?',
    `This will send the first ${batchSize} rows from "${sheetName}" to your webhook.\n\nThis does NOT affect main processing progress.\n\nProceed?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(sheetName);
  const totalRows = dataSheet.getLastRow();
  const totalCols = dataSheet.getLastColumn();
  const headers = dataSheet.getRange(1, 1, 1, totalCols).getValues()[0];

  if (totalRows < 2) { ui.alert('❌ No data found in the sheet.'); return; }

  const rowsToRead = Math.min(batchSize, totalRows - 1);
  const rawValues = dataSheet.getRange(2, 1, rowsToRead, totalCols).getValues();

  const data = rawValues.map((row, i) => {
    const obj = { row_number: 2 + i };
    headers.forEach((header, j) => {
      obj[header || `column_${j + 1}`] = row[j];
    });
    return obj;
  });

  if (data.length === 0) { ui.alert('❌ No valid data found.'); return; }

  const result = sendToPabbly(data, 2, true);

  getOrCreateLogSheet(ss).appendRow([
    new Date().toISOString(),
    '🧪 TEST BATCH',
    data.length,
    result.status,
    result.message
  ]);

  const success = result.status === 200;
  ui.alert(
    success ? '✅ Test Batch Sent!' : '❌ Test Batch Failed',
    `Rows sent: ${data.length}\nStatus: ${result.status}\nResponse: ${result.message.substring(0, 300)}`,
    ui.ButtonSet.OK
  );
}

// ── SEND TO PABBLY ────────────────────────────────────────────────────────────
function sendToPabbly(data, startRow, isTest) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty(CONFIG.WEBHOOK_URL_KEY);
  try {
    const response = UrlFetchApp.fetch(webhookUrl, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({
        is_test: isTest || false,
        batch_start_row: startRow,
        batch_size: data.length,
        data: data,
        timestamp: new Date().toISOString()
      }),
      muteHttpExceptions: true
    });
    return { status: response.getResponseCode(), message: response.getContentText().substring(0, 500) };
  } catch (e) {
    return { status: 'ERROR', message: e.toString() };
  }
}

// ── CHECK STATUS ──────────────────────────────────────────────────────────────
function checkStatus() {
  const props = PropertiesService.getScriptProperties();
  const isRunning = props.getProperty(CONFIG.IS_RUNNING_KEY) === 'true';
  const currentRow = parseInt(props.getProperty(CONFIG.START_ROW_KEY) || '2');
  const sheetName = props.getProperty(CONFIG.SHEET_NAME_KEY);
  const webhookUrl = props.getProperty(CONFIG.WEBHOOK_URL_KEY);
  const interval = parseInt(props.getProperty(CONFIG.INTERVAL_KEY) || '5');
  const batchSize = parseInt(props.getProperty(CONFIG.BATCH_SIZE_KEY) || '100');
  const dataSheet = sheetName ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) : null;
  const totalRows = dataSheet ? dataSheet.getLastRow() : 0;

  const rowsDone = currentRow - 2;
  const batchesDone = Math.floor(rowsDone / batchSize);
  const batchesLeft = Math.ceil((totalRows - currentRow + 1) / batchSize);
  const actualInterval = getValidInterval(interval);
  const minutesLeft = batchesLeft * actualInterval;
  const hoursLeft = (minutesLeft / 60).toFixed(1);

  const message = isRunning
    ? `🟢 Running\n\nRows sent: ${rowsDone} / ${totalRows - 1}\nBatches done: ${batchesDone}\nBatches remaining: ${batchesLeft}\nEst. time left: ${minutesLeft} min (~${hoursLeft} hrs)\n\n⚙️ Sheet: ${sheetName}\nBatch size: ${batchSize}\nInterval: every ${actualInterval} min\nWebhook: ...${webhookUrl ? webhookUrl.slice(-30) : 'Not set'}`
    : `🔴 Not running\nLast row processed: ${currentRow - 1 || 'N/A'}\n\n⚙️ Sheet: ${sheetName || 'Not set'}\nBatch size: ${batchSize}\nInterval: ${actualInterval ? `every ${actualInterval} min` : 'Not set'}\nWebhook: ...${webhookUrl ? webhookUrl.slice(-30) : 'Not set'}`;

  SpreadsheetApp.getUi().alert('📊 Pabbly Sync Status', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── STOP PROCESSING ───────────────────────────────────────────────────────────
function stopProcessing() {
  PropertiesService.getScriptProperties().setProperty(CONFIG.IS_RUNNING_KEY, 'false');
  deleteTrigger();
  SpreadsheetApp.getUi().alert('⛔ Stopped!', 'Processing has been stopped. You can restart anytime from the menu.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── LOG RESPONSE ──────────────────────────────────────────────────────────────
function logResponse(logSheet, startRow, endRow, count, result) {
  logSheet.appendRow([
    new Date().toISOString(),
    startRow === -1 ? result.status : `Rows ${startRow}–${endRow}`,
    count,
    result.status,
    result.message
  ]);
}

// ── GET OR CREATE LOG SHEET ───────────────────────────────────────────────────
function getOrCreateLogSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Batch Range', 'Row Count', 'HTTP Status', 'Response']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── VALID INTERVAL HELPER ─────────────────────────────────────────────────────
function getValidInterval(interval) {
  const validIntervals = [1, 5, 10, 15, 30];
  if (validIntervals.includes(interval)) return interval;
  return validIntervals.find(v => v >= interval) || 5;
}

// ── TRIGGER MANAGEMENT ────────────────────────────────────────────────────────
function setupTrigger() {
  deleteTrigger();
  let interval = parseInt(PropertiesService.getScriptProperties().getProperty(CONFIG.INTERVAL_KEY) || '5');
  const validInterval = getValidInterval(interval);

  if (validInterval !== interval) {
    PropertiesService.getScriptProperties().setProperty(CONFIG.INTERVAL_KEY, validInterval.toString());
    SpreadsheetApp.getUi().alert(
      `⚠️ Interval adjusted to ${validInterval} min`,
      `Apps Script only supports: 1, 5, 10, 15, 30 minutes.\nYour interval has been set to ${validInterval} min.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }

  ScriptApp.newTrigger('processSingleBatch').timeBased().everyMinutes(validInterval).create();
}

function deleteTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processSingleBatch')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
