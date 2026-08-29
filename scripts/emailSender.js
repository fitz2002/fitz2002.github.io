/* ─────────────────────────────────────────
   emailSender.js
   Microsoft Graph API authentication
   (OAuth 2.0 Authorization Code + PKCE)
   and email sending.

   ── ONE-TIME AZURE SETUP ──
   1. portal.azure.com → App registrations → New registration.

   2. "Supported account types":
        • Pick "Accounts in any organizational directory and
          personal Microsoft accounts" to leave `tenantId` as
          'common' below, OR
        • Pick "…this organizational directory only" (single
          tenant) and set `tenantId` below to your Directory
          (tenant) ID from the app's Overview page.

   3. Authentication → Add a platform → "Single-page application".
      Redirect URI = the EXACT URL this page is served from,
      e.g.  https://your-host/confirm-updater/
      (open the deployed page and copy what this file logs to the
      browser console as "[auth] Register this redirect URI").
      No client secret — PKCE is used instead.

   4. API permissions → Add → Microsoft Graph → Delegated →
      Mail.Send → Add permissions.
      Your tenant may require an admin to click
      "Grant admin consent for <org>".

   5. Copy the "Application (client) ID" from the Overview page
      into `clientId` below.

   Once those five steps are done, no further code changes are
   needed — set `clientId` (and `tenantId` if single-tenant) and
   deploy.
───────────────────────────────────────── */

// ── The only values that must be set for a new tenant ──
const AZURE_CONFIG = {
  clientId: 'YOUR_AZURE_CLIENT_ID_HERE',

  // 'common'        → any work/school or personal Microsoft account
  // 'organizations' → any work/school account
  // '<tenant-id>'   → only your organization (required if the app
  //                   registration is single-tenant)
  tenantId: 'common',

  // Mail.Send lets the app send as the signed-in user.
  // offline_access returns a refresh token so a send batch that
  // runs longer than the ~1h access-token lifetime keeps working.
  scopes: ['https://graph.microsoft.com/Mail.Send', 'offline_access'],
};

// The redirect URI is derived from wherever this page is hosted, so
// it never needs editing — but it MUST be registered verbatim in the
// Azure app registration (step 3 above).
const AZURE_REDIRECT_URI = window.location.origin + window.location.pathname;
console.info('[auth] Register this redirect URI in Azure:', AZURE_REDIRECT_URI);

const AZURE_AUTHORITY = `https://login.microsoftonline.com/${AZURE_CONFIG.tenantId}`;

// ── In-memory token state (access token also mirrored into the
// hidden #access-token field that app.js reads) ──
let _refreshToken   = null;
let _tokenExpiresAt = 0;     // epoch ms

// ─────────────────────────────────────────
//  PKCE / crypto helpers
// ─────────────────────────────────────────
function base64UrlEncode(bytes) {
  return btoa(String.fromCharCode(...new Uint8Array(bytes)))
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/, '');
}

// Random string from the PKCE "unreserved" set (hex is a valid subset)
function randomUrlSafe(length) {
  const bytes = new Uint8Array(length);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, b => ('0' + b.toString(16)).slice(-2)).join('').slice(0, length);
}

async function sha256(text) {
  return crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
}

function tokenEndpoint() {
  return `${AZURE_AUTHORITY}/oauth2/v2.0/token`;
}

// ── Surface an auth problem in the send log + status pill ──
function reportAuthError(message) {
  if (typeof addLog === 'function') addLog('error', `Authentication failed: ${message}`);
  else console.error('[auth]', message);

  const statusEl = document.getElementById('token-status');
  if (statusEl) {
    statusEl.className = '';
    statusEl.innerHTML = '<span class="dot"></span> Not authenticated';
  }
}

// ─────────────────────────────────────────
//  Sign-in: Authorization Code flow with PKCE, via popup
// ─────────────────────────────────────────
async function authenticateOutlook() {
  if (AZURE_CONFIG.clientId.startsWith('YOUR_AZURE_CLIENT_ID')) {
    reportAuthError('AZURE_CONFIG.clientId is not set in scripts/emailSender.js.');
    return;
  }

  const codeVerifier  = randomUrlSafe(96);
  const codeChallenge = base64UrlEncode(await sha256(codeVerifier));
  const state         = randomUrlSafe(32);

  sessionStorage.setItem('msauth_verifier', codeVerifier);
  sessionStorage.setItem('msauth_state', state);

  const authUrl = `${AZURE_AUTHORITY}/oauth2/v2.0/authorize?` + new URLSearchParams({
    client_id:             AZURE_CONFIG.clientId,
    response_type:         'code',
    redirect_uri:          AZURE_REDIRECT_URI,
    response_mode:         'query',
    scope:                 AZURE_CONFIG.scopes.join(' '),
    state,
    code_challenge:        codeChallenge,
    code_challenge_method: 'S256',
    prompt:                'select_account',   // always show the account picker
  });

  const popup = window.open(authUrl, 'msauth', 'width=500,height=650');
  if (!popup) {
    reportAuthError('popup was blocked — allow popups for this site and try again.');
    return;
  }

  // Poll the popup until it redirects back to our origin with ?code=…
  const interval = setInterval(async () => {
    let search;
    try {
      if (popup.closed) { clearInterval(interval); return; }
      search = popup.location.search;   // throws while on the login.microsoftonline.com pages
    } catch {
      return; // still cross-origin — keep waiting
    }

    if (!search) return;
    const returned = new URLSearchParams(search);
    const code = returned.get('code');
    const err  = returned.get('error');
    if (!code && !err) return;

    clearInterval(interval);
    try { popup.close(); } catch {}

    if (err) {
      reportAuthError(returned.get('error_description') || err);
      return;
    }
    if (returned.get('state') !== sessionStorage.getItem('msauth_state')) {
      reportAuthError('state mismatch — aborting for safety. Please try again.');
      return;
    }

    try {
      await exchangeCodeForTokens(code);
    } catch (e) {
      reportAuthError(e.message);
    }
  }, 400);
}

// ── Swap the authorization code for tokens (PKCE, no client secret) ──
async function exchangeCodeForTokens(code) {
  const response = await fetch(tokenEndpoint(), {
    method:  'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id:     AZURE_CONFIG.clientId,
      grant_type:    'authorization_code',
      code,
      redirect_uri:  AZURE_REDIRECT_URI,
      code_verifier: sessionStorage.getItem('msauth_verifier') || '',
      scope:         AZURE_CONFIG.scopes.join(' '),
    }),
  });

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(data.error_description || data.error || `token request failed (HTTP ${response.status})`);
  }
  storeTokens(data);
}

// ── Use the refresh token to get a fresh access token ──
async function refreshAccessToken() {
  const refreshToken = _refreshToken || sessionStorage.getItem('msauth_refresh');
  if (!refreshToken) {
    throw new Error('session expired — click "Sign in with Microsoft" again.');
  }

  const response = await fetch(tokenEndpoint(), {
    method:  'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id:     AZURE_CONFIG.clientId,
      grant_type:    'refresh_token',
      refresh_token: refreshToken,
      scope:         AZURE_CONFIG.scopes.join(' '),
    }),
  });

  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    _refreshToken = null;
    sessionStorage.removeItem('msauth_refresh');
    throw new Error(data.error_description || 'could not refresh session — sign in again.');
  }
  storeTokens(data);
  return data.access_token;
}

// ── Persist a token response and update the auth indicator ──
function storeTokens(data) {
  const tokenField = document.getElementById('access-token');
  if (tokenField) tokenField.value = data.access_token || '';

  if (data.refresh_token) {
    _refreshToken = data.refresh_token;
    sessionStorage.setItem('msauth_refresh', data.refresh_token);
  }
  _tokenExpiresAt = Date.now() + ((data.expires_in || 3600) * 1000);

  onTokenPasted();
}

// ── Returns a non-expired access token, refreshing if needed ──
async function getValidAccessToken() {
  const tokenField = document.getElementById('access-token');
  const current    = tokenField ? tokenField.value.trim() : '';

  // Refresh a little early so a token doesn't expire mid-request
  if (current && Date.now() < _tokenExpiresAt - 60_000) return current;
  return refreshAccessToken();
}

// ── On page load, silently restore a session if a refresh token
// survived in sessionStorage (e.g. after a reload) ──
document.addEventListener('DOMContentLoaded', () => {
  if (sessionStorage.getItem('msauth_refresh')) {
    refreshAccessToken().catch(() => {/* stay signed out, no noise */});
  }
});

// ── Updates the auth status indicator ──
function onTokenPasted() {
  const token    = (document.getElementById('access-token')?.value || '').trim();
  const statusEl = document.getElementById('token-status');
  if (!statusEl) return;

  if (token.length > 20) {
    statusEl.className = 'authenticated';
    statusEl.innerHTML = '<span class="dot"></span> Authenticated';
  } else {
    statusEl.className = '';
    statusEl.innerHTML = '<span class="dot"></span> Not authenticated';
  }
}

// ─────────────────────────────────────────
//  Sending
// ─────────────────────────────────────────

// ── Reads a File as base64 for the Graph API's fileAttachment format ──
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = () => resolve(reader.result.split(',')[1]); // strip the "data:...;base64," prefix
    reader.onerror = () => reject(new Error(`Failed to read attachment "${file.name}"`));
    reader.readAsDataURL(file);
  });
}

// ── Sends one email via Microsoft Graph API ──
// `attachments` is an array of { label, file } — only entries with a
// real File attached are sent. The `token` argument is optional; a
// valid token is fetched/refreshed automatically. Returns true on
// success, throws an Error on failure.
async function sendEmail(recipients, cc, subject, bodyText, token, attachments = []) {
  const filesToSend = attachments.filter(a => a.file);

  // Graph's sendMail accepts small attachments inline as base64; larger
  // files (~3MB+) would need an upload session instead, but rooming
  // lists and similar prep-day docs are well under that in practice.
  const graphAttachments = await Promise.all(filesToSend.map(async (a) => ({
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: a.file.name,
    contentBytes: await fileToBase64(a.file),
  })));

  const payload = {
    message: {
      subject:      subject,
      body:         { contentType: 'Text', content: bodyText },
      toRecipients: recipients.map(addr => ({ emailAddress: { address: addr } })),
      ccRecipients: cc.map(addr => ({ emailAddress: { address: addr } })),
      ...(graphAttachments.length ? { attachments: graphAttachments } : {}),
    },
    saveToSentItems: true,
  };

  const postOnce = async (bearer) => fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method:  'POST',
    headers: {
      'Authorization': `Bearer ${bearer}`,
      'Content-Type':  'application/json',
    },
    body: JSON.stringify(payload),
  });

  let bearer   = (token && token.trim()) || await getValidAccessToken();
  let response = await postOnce(bearer);

  // Access token expired mid-batch — refresh once and retry
  if (response.status === 401) {
    bearer   = await refreshAccessToken();
    response = await postOnce(bearer);
  }

  if (!response.ok) {
    const errorBody = await response.json().catch(() => ({ error: { message: response.statusText } }));
    throw new Error(errorBody.error?.message || `HTTP ${response.status}`);
  }

  return true;
}
