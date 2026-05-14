// ============================================================
// Config.gs — Constants and credential management
// ============================================================

const SETTINGS_TAB    = 'Settings';
const ORDERS_TAB      = 'Orders';
const CREDS_ROW_START = 2;
const BRICKLINK_API_BASE_URL = 'https://api.bricklink.com/api/store/v1/';
const OAUTH_SIGNATURE_METHOD = 'HMAC-SHA1';
const OAUTH_VERSION   = '1.0';

const ORDERS_HEADERS = [
  'Order ID', 'Date', 'Buyer', 'Items', 'Lots',
  'Shipping Charged', 'Shipping Actual', 'Grand Total', 'Currency',
  'Payment Method', 'Status', 'Refund', 'Notes'
];

const MANUAL_COLS   = [7, 12, 13]; // Shipping Actual, Refund, Notes (1-indexed)
const MANUAL_COLOUR = '#FFF9C4';   // light yellow
const HEADER_COLOUR = '#f3f3f3';
const COL_WIDTHS    = [100, 110, 150, 60, 60, 130, 120, 110, 80, 200, 110, 90, 200];

// -------------------------------------------------------
// Credential read/write
// -------------------------------------------------------

function getCredentials() {
  const props = PropertiesService.getUserProperties();
  return {
    consumerKey:      props.getProperty('BL_CONSUMER_KEY'),
    consumerSecret:   props.getProperty('BL_CONSUMER_SECRET'),
    accessToken:      props.getProperty('BL_ACCESS_TOKEN'),
    accessTokenSecret: props.getProperty('BL_ACCESS_TOKEN_SECRET'),
    storeUsername:    props.getProperty('BL_STORE_USERNAME')
  };
}

function credentialsExist() {
  return !!PropertiesService.getUserProperties().getProperty('BL_CONSUMER_KEY');
}

function saveCredentialsFromSidebar(creds) {
  const props = PropertiesService.getUserProperties();
  if (creds.consumerKey)       props.setProperty('BL_CONSUMER_KEY',        creds.consumerKey);
  if (creds.consumerSecret)    props.setProperty('BL_CONSUMER_SECRET',     creds.consumerSecret);
  if (creds.accessToken)       props.setProperty('BL_ACCESS_TOKEN',        creds.accessToken);
  if (creds.accessTokenSecret) props.setProperty('BL_ACCESS_TOKEN_SECRET', creds.accessTokenSecret);
  if (creds.storeUsername)     props.setProperty('BL_STORE_USERNAME',      creds.storeUsername);
  logStatus(`Credentials saved: ${new Date().toLocaleString()}`);
}

function clearCredentials() {
  const props = PropertiesService.getUserProperties();
  ['BL_CONSUMER_KEY', 'BL_CONSUMER_SECRET', 'BL_ACCESS_TOKEN', 'BL_ACCESS_TOKEN_SECRET', 'BL_STORE_USERNAME']
    .forEach(k => props.deleteProperty(k));
  logStatus(`Credentials cleared: ${new Date().toLocaleString()}`);
}

// -------------------------------------------------------
// One-time Settings tab setup
// -------------------------------------------------------

function setupSettingsTab() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_TAB);

  const labels = [
    ['API Credentials — enter values in column B, then run saveCredentials()', ''],
    ['Consumer Key',    ''],
    ['Consumer Secret', ''],
    ['Token Value',     ''],
    ['Token Secret',    ''],
    ['Store Username',  ''],
  ];

  sheet.getRange(1, 1, labels.length, 2).setValues(labels);
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);

  SpreadsheetApp.getUi().alert('Settings tab ready. Fill in column B then run saveCredentials().');
}

function saveCredentials() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_TAB);
  const props = PropertiesService.getUserProperties();

  const keys = ['BL_CONSUMER_KEY', 'BL_CONSUMER_SECRET', 'BL_ACCESS_TOKEN', 'BL_ACCESS_TOKEN_SECRET', 'BL_STORE_USERNAME'];

  keys.forEach((key, i) => {
    const val = sheet.getRange(CREDS_ROW_START + i, 2).getValue().toString().trim();
    if (val && val !== '✓ Saved') {
      props.setProperty(key, val);
    }
  });

  sheet.getRange(CREDS_ROW_START, 2, 5, 1).clearContent();
  sheet.getRange(CREDS_ROW_START, 2).setValue('✓ Saved');

  SpreadsheetApp.getUi().alert('Credentials saved securely. Column B has been cleared.');
}
