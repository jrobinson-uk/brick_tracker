// ============================================================
// Debug.gs — API call tracking and debug infrastructure
// ============================================================

let apiCallLog_ = {};

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
  const message = `[${functionName}] ${total} API call(s) — ${detail}`;

  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName(SETTINGS_TAB);
  const logRow  = 13;
  const existing = sheet.getRange(logRow, 2).getValue();
  const updated  = `${new Date().toLocaleString()} — ${message}` +
                   (existing ? `\n${existing}` : '');

  sheet.getRange(logRow, 1).setValue('API Log:');
  sheet.getRange(logRow, 2).setValue(updated.split('\n').slice(0, 10).join('\n'));

  Logger.log(message);
  apiCallLog_ = {};
}
