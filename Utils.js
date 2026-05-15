// ============================================================
// Utils.gs — Core API request and connection test
// ============================================================

function bricklinkRequest(endpoint, method, queryParams, bodyParams) {
  method      = method      || 'GET';
  queryParams = queryParams || {};
  bodyParams  = bodyParams  || {};

  const creds = getCredentials();
  if (!creds.consumerKey) {
    throw new Error('No credentials found. Run saveCredentials() first.');
  }

  const fullUrl = BRICKLINK_API_BASE_URL + endpoint;
  const baseUrl = fullUrl.split('?')[0];

  const timestamp = Math.floor(Date.now() / 1000).toString();
  const nonce     = Utilities.getUuid().replace(/-/g, '');

  const allParams = {
    oauth_consumer_key:     creds.consumerKey,
    oauth_token:            creds.accessToken,
    oauth_signature_method: OAUTH_SIGNATURE_METHOD,
    oauth_timestamp:        timestamp,
    oauth_nonce:            nonce,
    oauth_version:          OAUTH_VERSION
  };

  for (const key in queryParams) {
    if (queryParams.hasOwnProperty(key)) allParams[key] = queryParams[key];
  }

  const normalizedParams = normalizeParameters(allParams);
  const baseString       = generateSignatureBaseString(method, baseUrl, normalizedParams);
  const signature        = generateHmacSha1Signature(baseString, creds.consumerSecret, creds.accessTokenSecret);

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
    headers: { 'Authorization': authHeader, 'Accept': 'application/json' },
    muteHttpExceptions: true
  };

  if (Object.keys(bodyParams).length > 0) {
    options.contentType = 'application/json';
    options.payload     = JSON.stringify(bodyParams);
  }

  let requestUrl = fullUrl;
  if (Object.keys(queryParams).length > 0) {
    const queryString = Object.keys(queryParams)
      .map(k => `${urlEncode(k)}=${urlEncode(queryParams[k])}`)
      .join('&');
    requestUrl = baseUrl + '?' + queryString;
  }

  trackApiCall_(endpoint, method);
  const response = UrlFetchApp.fetch(requestUrl, options);

  Logger.log('Request URL: ' + requestUrl);
  Logger.log('Response code: ' + response.getResponseCode());
  Logger.log('Response body: ' + response.getContentText());

  return JSON.parse(response.getContentText());
}

function testConnection() {
  try {
    const creds = getCredentials();
    const data  = bricklinkRequest('orders', 'GET', { direction: 'in' });

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
