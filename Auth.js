// ============================================================
// Auth.gs — OAuth 1.0a signing
// ============================================================

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
