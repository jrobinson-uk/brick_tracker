// ============================================================
// JR's Bricks Profit Tracker — Phase 3: Order Sync
// ============================================================

// -------------------------------------------------------
// MENU
// -------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BrickLink Tracker')
    .addItem('Sync Orders',     'syncOrders')
    .addItem('Sync Inventory',  'syncInventory')
    .addItem('Log Purchase',    'logPurchase')
    .addItem('View Dashboard',  'viewDashboard')
    .addSeparator()
    .addItem('Settings',        'openSettings')
    .addToUi();

}

// -------------------------------------------------------
// ORDER SYNC
// -------------------------------------------------------

const LAST_ORDER_SYNC_KEY = 'LAST_ORDER_SYNC';

function setupOrdersTab() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(ORDERS_TAB);
  if (!sheet) sheet = ss.insertSheet(ORDERS_TAB);

  sheet.clearContents();
  const header = sheet.getRange(1, 1, 1, ORDERS_HEADERS.length);
  header.setValues([ORDERS_HEADERS]);
  header.setFontWeight('bold');
  header.setBackground('#f3f3f3');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, ORDERS_HEADERS.length);
}

function syncOrders() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(ORDERS_TAB);
  if (!sheet) {
    setupOrdersTab();
    sheet = ss.getSheetByName(ORDERS_TAB);
  }

  const lastSync = props.getProperty(LAST_ORDER_SYNC_KEY);
  logStatus('Syncing orders…');

  try {
    const params = { direction: 'in' };
    if (lastSync) params.filed = 'false'; // only open/recent orders on incremental sync

    const data = bricklinkRequest('orders', 'GET', params);

    if (!data.meta || data.meta.code !== 200) {
      ui.alert(`⚠️ API error:\n\n${JSON.stringify(data.meta)}`);
      logStatus(`Order sync failed: ${new Date().toLocaleString()}`);
      return;
    }

    const orders = data.data || [];

    // On incremental sync, filter to orders newer than last sync
    const filtered = lastSync
      ? orders.filter(o => new Date(o.date_ordered) > new Date(lastSync))
      : orders;

    if (filtered.length === 0) {
      ui.alert('✅ Orders up to date — no new orders since last sync.');
      logStatus(`Orders synced: ${new Date().toLocaleString()} — no new orders`);
      props.setProperty(LAST_ORDER_SYNC_KEY, new Date().toISOString());
      return;
    }

    // Build rows — newest first
    filtered.sort((a, b) => new Date(b.date_ordered) - new Date(a.date_ordered));

    const rows = filtered.map(o => [
      o.order_id,
      o.date_ordered ? new Date(o.date_ordered).toLocaleDateString('en-GB') : '',
      o.buyer_name   || '',
      o.total_count  || 0,
      o.unique_count || 0,
      o.cost         ? parseFloat(o.cost.subtotal)    : '',
      o.cost         ? parseFloat(o.cost.shipping)    : '',
      o.cost         ? parseFloat(o.cost.grand_total) : '',
      o.cost         ? o.cost.currency_code           : '',
      o.payment      ? o.payment.method               : '',
      o.status       || ''
    ]);

    // Append after existing data
    const lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(lastRow + 1, 1, rows.length, ORDERS_HEADERS.length).setValues(rows);
    sheet.autoResizeColumns(1, ORDERS_HEADERS.length);

    props.setProperty(LAST_ORDER_SYNC_KEY, new Date().toISOString());

    const msg = lastSync
      ? `✅ Synced ${filtered.length} new order(s).`
      : `✅ Full sync complete — ${filtered.length} order(s) imported.`;

    ui.alert(msg);
    logStatus(`Orders synced: ${new Date().toLocaleString()} — ${filtered.length} new`);

  } catch (e) {
    ui.alert(`❌ Sync failed:\n\n${e.message}`);
    logStatus(`Order sync failed: ${new Date().toLocaleString()}`);
  }
}

// -------------------------------------------------------
// MENU STUBS — filled in by later phases
// -------------------------------------------------------

function syncInventory() {
  SpreadsheetApp.getUi().alert('Sync Inventory — coming in Phase 6.');
}

function logPurchase() {
  SpreadsheetApp.getUi().alert('Log Purchase — coming in Phase 5.');
}

function viewDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName('Dashboard');
  if (dash) {
    ss.setActiveSheet(dash);
  } else {
    SpreadsheetApp.getUi().alert('Dashboard tab not set up yet — coming in Phase 8.');
  }
}

function openSettings() {
  showSetupSidebar();
}

// -------------------------------------------------------
// SETUP SIDEBAR
// -------------------------------------------------------

function showSetupSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('BrickLink Tracker Setup');
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveCredentialsFromSidebar(creds) {
  const props = PropertiesService.getUserProperties();
  if (creds.consumerKey)       props.setProperty('BL_CONSUMER_KEY',          creds.consumerKey);
  if (creds.consumerSecret)    props.setProperty('BL_CONSUMER_SECRET',       creds.consumerSecret);
  if (creds.accessToken)       props.setProperty('BL_ACCESS_TOKEN',          creds.accessToken);
  if (creds.accessTokenSecret) props.setProperty('BL_ACCESS_TOKEN_SECRET',   creds.accessTokenSecret);
  if (creds.storeUsername)     props.setProperty('BL_STORE_USERNAME',        creds.storeUsername);
  logStatus(`Credentials saved: ${new Date().toLocaleString()}`);
}

function credentialsExist() {
  return !!PropertiesService.getUserProperties().getProperty('BL_CONSUMER_KEY');
}

function clearCredentials() {
  const props = PropertiesService.getUserProperties();
  ['BL_CONSUMER_KEY', 'BL_CONSUMER_SECRET', 'BL_ACCESS_TOKEN', 'BL_ACCESS_TOKEN_SECRET', 'BL_STORE_USERNAME']
    .forEach(k => props.deleteProperty(k));
  logStatus(`Credentials cleared: ${new Date().toLocaleString()}`);
}

// -------------------------------------------------------
// CONSTANTS
// -------------------------------------------------------

const SETTINGS_TAB  = 'Settings';
const ORDERS_TAB    = 'Orders';
const CREDS_ROW_START = 2;
const BRICKLINK_API_BASE_URL = 'https://api.bricklink.com/api/store/v1/';
const OAUTH_SIGNATURE_METHOD = 'HMAC-SHA1';
const OAUTH_VERSION = '1.0';

const ORDERS_HEADERS = [
  'Order ID', 'Date', 'Buyer', 'Items', 'Lots',
  'Subtotal', 'Shipping', 'Grand Total', 'Currency',
  'Payment Method', 'Status'
];

// -------------------------------------------------------
// SETUP — run once to label the Settings tab
// -------------------------------------------------------

function setupSettingsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

// -------------------------------------------------------
// CREDENTIALS
// -------------------------------------------------------

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

function getCredentials() {
  const props = PropertiesService.getUserProperties();
  return {
    consumerKey:     props.getProperty('BL_CONSUMER_KEY'),
    consumerSecret:  props.getProperty('BL_CONSUMER_SECRET'),
    accessToken:     props.getProperty('BL_ACCESS_TOKEN'),
    accessTokenSecret: props.getProperty('BL_ACCESS_TOKEN_SECRET'),
    storeUsername:   props.getProperty('BL_STORE_USERNAME')
  };
}

// -------------------------------------------------------
// OAuth 1.0a HELPERS
// -------------------------------------------------------

function urlEncode(str) {
  return encodeURIComponent(String(str))
    .replace(/[!'()*]/g, c => '%' + c.charCodeAt(0).toString(16).toUpperCase());
}

function normalizeParameters(params) {
  const normalized = [];
  for (const key in params) {
    if (params.hasOwnProperty(key)) {
      normalized.push(`${urlEncode(key)}=${urlEncode(params[key])}`);
    }
  }
  return normalized.sort().join('&');
}

function generateSignatureBaseString(httpMethod, baseUrl, normalizedParameters) {
  return `${urlEncode(httpMethod.toUpperCase())}&${urlEncode(baseUrl)}&${urlEncode(normalizedParameters)}`;
}

function generateHmacSha1Signature(baseString, consumerSecret, accessTokenSecret) {
  const signingKey = `${urlEncode(consumerSecret)}&${urlEncode(accessTokenSecret)}`;
  const signatureBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_1,
    baseString,
    signingKey
  );
  return Utilities.base64Encode(signatureBytes);
}

function buildAuthorizationHeader(oauthParams) {
  const headerParts = [];
  for (const key in oauthParams) {
    if (oauthParams.hasOwnProperty(key)) {
      headerParts.push(`${urlEncode(key)}="${urlEncode(oauthParams[key])}"`);
    }
  }
  return `OAuth realm="",${headerParts.sort().join(',')}`;
}

// -------------------------------------------------------
// CORE API REQUEST — used by every future phase
// -------------------------------------------------------

function bricklinkRequest(endpoint, method, queryParams, bodyParams) {
  method = method || 'GET';
  queryParams = queryParams || {};
  bodyParams = bodyParams || {};

  const creds = getCredentials();
  if (!creds.consumerKey) {
    throw new Error('No credentials found. Run saveCredentials() first.');
  }

  const fullUrl = BRICKLINK_API_BASE_URL + endpoint;
  const baseUrl = fullUrl.split('?')[0];

  const timestamp = Math.floor(Date.now() / 1000).toString();
  const nonce = Utilities.getUuid().replace(/-/g, '');

  // Merge OAuth params + query params for signing
  const allParams = {
    oauth_consumer_key:     creds.consumerKey,
    oauth_token:            creds.accessToken,
    oauth_signature_method: OAUTH_SIGNATURE_METHOD,
    oauth_timestamp:        timestamp,
    oauth_nonce:            nonce,
    oauth_version:          OAUTH_VERSION
  };

  for (const key in queryParams) {
    if (queryParams.hasOwnProperty(key)) {
      allParams[key] = queryParams[key];
    }
  }

  const normalizedParams = normalizeParameters(allParams);
  const baseString = generateSignatureBaseString(method, baseUrl, normalizedParams);
  const signature = generateHmacSha1Signature(baseString, creds.consumerSecret, creds.accessTokenSecret);

  const oauthHeaderParams = {
    oauth_consumer_key:     creds.consumerKey,
    oauth_token:            creds.accessToken,
    oauth_signature_method: OAUTH_SIGNATURE_METHOD,
    oauth_timestamp:        timestamp,
    oauth_nonce:            nonce,
    oauth_version:          OAUTH_VERSION,
    oauth_signature:        signature
  };

  const authHeader = buildAuthorizationHeader(oauthHeaderParams);

  const options = {
    method: method,
    headers: {
      'Authorization': authHeader,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (Object.keys(bodyParams).length > 0) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(bodyParams);
  }

  // Build final URL with query string
  let requestUrl = fullUrl;
  if (Object.keys(queryParams).length > 0) {
    const queryString = Object.keys(queryParams)
      .map(k => `${urlEncode(k)}=${urlEncode(queryParams[k])}`)
      .join('&');
    requestUrl = baseUrl + '?' + queryString;
  }

  const response = UrlFetchApp.fetch(requestUrl, options);
  const code = response.getResponseCode();
  const body = response.getContentText();

  Logger.log('Request URL: ' + requestUrl);
  Logger.log('Response code: ' + code);
  Logger.log('Response body: ' + body);

  return JSON.parse(body);
}

// -------------------------------------------------------
// TEST CONNECTION — Phase 1 done-when condition
// -------------------------------------------------------

function testConnection() {
  try {
    const creds = getCredentials();
    const data = bricklinkRequest('orders', 'GET', { direction: 'in' });

    if (data.meta && data.meta.code === 200) {
      SpreadsheetApp.getUi().alert(`✅ Connected!\n\nStore: ${creds.storeUsername}`);
      logStatus(`Last connection test: ${new Date().toLocaleString()} — OK (${creds.storeUsername})`);
    } else {
      SpreadsheetApp.getUi().alert(`⚠️ API responded but returned:\n\n${JSON.stringify(data.meta)}`);
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert(`❌ Failed:\n\n${e.message}`);
    logStatus(`Last connection test: ${new Date().toLocaleString()} — FAILED`);
  }
}

// -------------------------------------------------------
// STATUS BAR
// -------------------------------------------------------

function logStatus(message) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_TAB);
  sheet.getRange(10, 1).setValue('Status:');
  sheet.getRange(10, 2).setValue(message);
}