// ============================================================
// UI.gs — Menu, sidebars, status bar
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BrickLink Tracker')
    .addItem('Sync Orders',    'syncOrders')
    .addItem('Sync Inventory', 'syncInventory')
    .addItem('Log Purchase',   'logPurchase')
    .addItem('View Dashboard', 'viewDashboard')
    .addSeparator()
    .addItem('Settings',       'openSettings')
    .addSeparator()
    .addItem('Clear Debug Log', 'clearDebugLog')
    .addToUi();
}

function openSettings() {
  showSetupSidebar();
}

function showSetupSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('BrickLink Tracker Setup');
  SpreadsheetApp.getUi().showSidebar(html);
}

function logStatus(message) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_TAB);
  sheet.getRange(10, 1).setValue('Status:');
  sheet.getRange(10, 2).setValue(message);
}

// -------------------------------------------------------
// Menu stubs — filled in by later phases
// -------------------------------------------------------

function syncInventory() {
  SpreadsheetApp.getUi().alert('Sync Inventory — coming in Phase 6.');
}

function logPurchase() {
  SpreadsheetApp.getUi().alert('Log Purchase — coming in Phase 5.');
}

function viewDashboard() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName('Dashboard');
  if (dash) {
    ss.setActiveSheet(dash);
  } else {
    SpreadsheetApp.getUi().alert('Dashboard tab not set up yet — coming in Phase 8.');
  }
}
