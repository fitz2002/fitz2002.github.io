/* ─────────────────────────────────────────
   emailSender.js
   Microsoft Graph API authentication
   and email sending.

   TO SET UP:
     1. Register an app in Azure Portal
        (portal.azure.com → App registrations)
     2. Under "Supported account types", select:
        "Accounts in any organizational directory
        and personal Microsoft accounts"
        (this enables the 'common' endpoint below)
     3. Under "Authentication", add a
        Single-page app redirect URI pointing
        to this page's URL
     4. Grant the "Mail.Send" delegated permission
     5. Replace AZURE_CLIENT_ID below with your
        app's client ID — that's the only value
        you need to change
───────────────────────────────────────── */

// ── Only this value needs to be set — tenant stays 'common' ──
const AZURE_CLIENT_ID = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a';

// 'common' allows any Microsoft account (personal Outlook,
// work/school Office 365) to sign in — do not change this.
const AZURE_TENANT_ID = 'common';

// ── Opens the Microsoft OAuth2 popup and captures the token ──
function authenticateOutlook() {
  const redirectUri = encodeURIComponent(window.location.href.split('?')[0]);

  // offline_access keeps the session alive; Mail.Send is the only
  // permission needed to send emails on behalf of the signed-in user.
  const scope = encodeURIComponent('https://graph.microsoft.com/Mail.Send offline_access');

  const authUrl = [
    `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/authorize`,
    `?client_id=${AZURE_CLIENT_ID}`,
    `&response_type=token`,
    `&redirect_uri=${redirectUri}`, //url of my page whereever that ends up
    `&scope=${scope}`,
    `&response_mode=fragment`,
    `&prompt=select_account`,   // Always show account picker so different users can log in
  ].join('');

  const popup = window.open(authUrl, 'msauth', 'width=500,height=650');

  // Poll the popup for the access token in the URL hash
  const interval = setInterval(() => {
    try {
      const hash = popup.location.hash;
      if (hash && hash.includes('access_token')) {
        const params = new URLSearchParams(hash.substring(1));
        const token  = params.get('access_token');

        document.getElementById('access-token').value = token;
        onTokenPasted();

        popup.close();
        clearInterval(interval);
      }
    } catch (e) {
      // Cross-origin error while popup is loading — safe to ignore
    }

    if (popup.closed) clearInterval(interval);
  }, 500);
}

// ── Updates the auth status indicator when a token is entered or captured ──
function onTokenPasted() {
  const token    = document.getElementById('access-token').value.trim();
  const statusEl = document.getElementById('token-status');

  if (token.length > 20) {
    statusEl.className = 'authenticated';
    statusEl.innerHTML = '<span class="dot"></span> Authenticated';
  } else {
    statusEl.className = '';
    statusEl.innerHTML = '<span class="dot"></span> Not authenticated';
  }
}

// ── Sends one email via Microsoft Graph API ──
// Returns true on success, throws an Error on failure
async function sendEmail(recipients, cc, subject, bodyText, token) {
  const payload = {
    message: {
      subject:      subject,
      body:         { contentType: 'Text', content: bodyText },
      toRecipients: recipients.map(addr => ({ emailAddress: { address: addr } })),
      ccRecipients: cc.map(addr => ({ emailAddress: { address: addr } })),
    },
    saveToSentItems: true,
  };

  const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
    method:  'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type':  'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const errorBody = await response.json().catch(() => ({ error: { message: response.statusText } }));
    throw new Error(errorBody.error?.message || `HTTP ${response.status}`);
  }

  return true;
}
