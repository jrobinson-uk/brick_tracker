// ============================================================
// Debug.gs — API call tracking and debug log
// ============================================================

const DEBUG_TAB     = 'Debug';
const DEBUG_HEADERS = ['Timestamp', 'Function', 'Message'];

let apiCallLog_ = {};

// -------------------------------------------------------
// Debug tab setup
// -------------------------------------------------------

function setupDebugTab() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DEBUG_TAB);
  if (!sheet) sheet = ss.insertSheet(DEBUG_TAB);

  sheet.clearContents();

  const headerRange = sheet.getRange(1, 1, 1, DEBUG_HEADERS.length);
  headerRange.setValues([DEBUG_HEADERS]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground(HEADER_COLOUR);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160); // Timestamp
  sheet.setColumnWidth(2, 140); // Function
  sheet.setColumnWidth(3, 600); // Message

  return sheet;
}

function clearDebugLog() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEBUG_TAB);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, DEBUG_HEADERS.length).clearContent();
  }

  logStatus(`Debug log cleared: ${new Date().toLocaleString()}`);
}

// -------------------------------------------------------
// Core log writer — auto-creates tab on first use
// -------------------------------------------------------

function debugLog(functionName, message) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DEBUG_TAB);
  if (!sheet) sheet = setupDebugTab();

  sheet.appendRow([new Date().toLocaleString('en-GB'), functionName, message]);
}

// -------------------------------------------------------
// API call tracking
// -------------------------------------------------------

function resetApiCallLog() {
  apiCallLog_ = {};
}

function trackApiCall_(endpoint) {
  const key = endpoint.split('?')[0];
  apiCallLog_[key] = (apiCallLog_[key] || 0) + 1;
}

function flushApiCallLog(functionName) {
  const total   = Object.values(apiCallLog_).reduce((a, b) => a + b, 0);
  const detail  = Object.entries(apiCallLog_).map(([k, v]) => `${k}: ${v}`).join(', ');
  const message = `${total} API call(s) — ${detail}`;

  // Keep summary in Settings tab
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const sheet    = ss.getSheetByName(SETTINGS_TAB);
  const existing = sheet.getRange(13, 2).getValue();
  const updated  = `${new Date().toLocaleString()} — [${functionName}] ${message}` +
                   (existing ? `\n${existing}` : '');
  sheet.getRange(13, 1).setValue('API Log:');
  sheet.getRange(13, 2).setValue(updated.split('\n').slice(0, 10).join('\n'));

  // Full entry in Debug tab
  debugLog(functionName, message);

  Logger.log(`[${functionName}] ${message}`);
  apiCallLog_ = {};
}
