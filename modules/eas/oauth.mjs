/**
 * Microsoft 365 OAuth 2.0 for EAS.
 *
 * Reuses the legacy EAS-4-TbSync community Azure AD application
 * (2980deeb-7460-4723-864a-f9b0f10cd992) which is registered with the
 * `https://login.microsoftonline.com/common/oauth2/nativeclient` redirect
 * URI. This means we can NOT use `browser.identity.launchWebAuthFlow`
 * (that API forces its own redirect URL and would require a different
 * Azure AD app). Instead we drive the consent popup ourselves: open the
 * auth URL in a popup window, watch `tabs.onUpdated` until the active
 * tab navigates to `…/nativeclient?code=…`, parse the code from the URL,
 * and exchange it at the token endpoint. Mirrors what Mozilla's
 * `OAuth2.sys.mjs` did internally for the legacy add-on.
 *
 * Refresh tokens persist in `account.custom.refreshToken`. The OAuth
 * client ID lives in `browser.storage.local["oauth.clientID"]` - a
 * single global value shared by every account in the profile; empty or
 * missing falls back to DEFAULT_OAUTH_CLIENT_ID. Scope is derived from
 * `account.custom.servertype` via `scopeForServertype`. The same
 * `servertype` field is the discriminator for "is this an OAuth
 * account?" (see `isOAuthAccount`). Access tokens are kept in an
 * in-memory cache only; we refresh transparently on expiry. `primeAuth`
 * seeds the auth cache at the top of every on* hook that hits the
 * network.
 */

import { ERR, withCode } from "../../vendor/tbsync/provider.mjs";

/**
 * Default Application (client) ID, registered in Microsoft Entra ID with
 * the nativeclient redirect URI.
 */
export const DEFAULT_OAUTH_CLIENT_ID = "2980deeb-7460-4723-864a-f9b0f10cd992";

/**
 * Global storage.local key that holds the user's custom OAuth client ID.
 * One value shared by every account in the profile; empty or missing
 * means "use DEFAULT_OAUTH_CLIENT_ID".
 */
const CUSTOM_OAUTH_CLIENT_ID_STORAGE_KEY = "oauth.clientID";

/**
 * Resolve the OAuth client ID. Reads `oauth.clientID` from
 * `browser.storage.local`; falls back to DEFAULT_OAUTH_CLIENT_ID when missing
 * or empty. Called from every call site that builds a Microsoft request.
 */
export async function getGlobalClientID() {
  try {
    const rv = await browser.storage.local.get({
      [CUSTOM_OAUTH_CLIENT_ID_STORAGE_KEY]: "",
    });
    const v = rv[CUSTOM_OAUTH_CLIENT_ID_STORAGE_KEY];
    if (typeof v === "string" && v.trim()) return v.trim();
  } catch {
    /* fall through */
  }
  return DEFAULT_OAUTH_CLIENT_ID;
}

const AUTH_ENDPOINT =
  "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
const TOKEN_ENDPOINT =
  "https://login.microsoftonline.com/common/oauth2/v2.0/token";
const REDIRECT_URI =
  "https://login.microsoftonline.com/common/oauth2/nativeclient";

/** Microsoft scope strings keyed by `account.custom.servertype`. Only
 *  the OAuth-capable setup types (`office365`, `personal-ms`) appear
 *  here. */
const SCOPE_BY_SERVERTYPE = {
  office365: "offline_access https://outlook.office.com/.default",
  "personal-ms":
    "offline_access https://outlook.office.com/EAS.AccessAsUser.All",
};

/** Look up the OAuth scope for a given setup-type. Throws if the type
 *  is not OAuth-capable. */
export function scopeForServertype(servertype) {
  const scope = SCOPE_BY_SERVERTYPE[servertype];
  if (!scope)
    throw new Error(`OAuth scope unknown for servertype '${servertype}'`);
  return scope;
}

/** Returns true iff the account's `servertype` is one of the OAuth-capable
 *  setup types. */
export function isOAuthAccount(custom) {
  return custom?.servertype != null && custom.servertype in SCOPE_BY_SERVERTYPE;
}

/** Refresh 30 s before the token expires. */
const REFRESH_SKEW_MS = 30_000;

/** accountId → { token, expiresAt } */
const accessTokenCache = new Map();

/** accountId → { refreshToken, servertype }. */
const authCache = new Map();

export function primeAuth(accountId, { refreshToken, servertype }) {
  if (!refreshToken || !servertype) return;
  authCache.set(accountId, { refreshToken, servertype });
}

export function forgetAuth(accountId) {
  authCache.delete(accountId);
  accessTokenCache.delete(accountId);
}

export function invalidateAccessToken(accountId) {
  accessTokenCache.delete(accountId);
}

export function primeAccessToken(accountId, token, expiresIn) {
  accessTokenCache.set(accountId, {
    token,
    expiresAt: Date.now() + (expiresIn ?? 3600) * 1000,
  });
}

// ── Interactive sign-in ────────────────────────────────────────────────────

/**
 * Open the Microsoft consent popup and return the resulting tokens.
 *
 *   loginHint   pre-selects an account on the consent screen
 *   servertype  "office365" | "personal-ms" - drives scope
 *
 * The OAuth client ID is read from the global `oauth.clientID` slot
 * (storage.local), with the hardcoded community ID as fallback. */
export async function startAuth({ loginHint, servertype }) {
  const clientID = await getGlobalClientID();
  const scope = scopeForServertype(servertype);
  const state = crypto.randomUUID();

  const authUrl = new URL(AUTH_ENDPOINT);
  authUrl.searchParams.set("client_id", clientID);
  authUrl.searchParams.set("redirect_uri", REDIRECT_URI);
  authUrl.searchParams.set("response_type", "code");
  authUrl.searchParams.set("response_mode", "query");
  authUrl.searchParams.set("scope", scope);
  authUrl.searchParams.set("state", state);
  if (loginHint) authUrl.searchParams.set("login_hint", loginHint);

  const responseUrl = await runConsentPopup(authUrl.toString());

  // Parse the redirect URL - Microsoft echoes the code back as a query
  // string on the nativeclient page.
  const parsed = new URL(responseUrl);
  const error = parsed.searchParams.get("error");
  if (error) {
    const desc = parsed.searchParams.get("error_description") ?? "";
    throw withCode(
      new Error(`Microsoft returned: ${error} ${desc}`.trim()),
      ERR.AUTH,
    );
  }
  const code = parsed.searchParams.get("code");
  const returnedState = parsed.searchParams.get("state");
  if (!code)
    throw withCode(new Error("No authorization code in response"), ERR.AUTH);
  if (returnedState !== state)
    throw withCode(new Error("OAuth state mismatch (possible CSRF)"), ERR.AUTH);

  const tokens = await exchangeCode({ clientID, code, scope });
  if (!tokens.refresh_token) {
    throw withCode(new Error("No refresh_token in token response"), ERR.AUTH);
  }
  const authenticatedUserEmail =
    decodeIdTokenEmail(tokens.id_token) ?? loginHint ?? null;
  return {
    refreshToken: tokens.refresh_token,
    accessToken: tokens.access_token,
    expiresIn: tokens.expires_in,
    authenticatedUserEmail,
  };
}

/** Exchange the authorization code for tokens. */
async function exchangeCode({ clientID, code, scope }) {
  const body = new URLSearchParams({
    client_id: clientID,
    grant_type: "authorization_code",
    code,
    redirect_uri: REDIRECT_URI,
    scope,
  });
  const resp = await fetch(TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  const text = await resp.text().catch(() => "");
  if (!resp.ok) {
    throw withCode(
      new Error(`Token exchange failed (${resp.status}): ${text}`),
      ERR.AUTH,
    );
  }
  try {
    return JSON.parse(text);
  } catch {
    throw withCode(new Error("Invalid token-exchange response"), ERR.AUTH);
  }
}

// ── Refresh-on-demand ─────────────────────────────────────────────────────

/** Returns a valid access token. Refreshes transparently using the cached
 *  refresh token when the in-memory access token is expired or missing.
 *  `primeAuth(accountId, …)` must have been called first. */
export async function getAccessToken(accountId) {
  const cached = accessTokenCache.get(accountId);
  if (cached && cached.expiresAt > Date.now() + REFRESH_SKEW_MS) {
    return cached.token;
  }
  const auth = authCache.get(accountId);
  if (!auth) {
    throw withCode(
      new Error("OAuth auth not primed - call primeAuth first"),
      ERR.AUTH,
    );
  }
  const fresh = await refreshAccessToken(auth);
  accessTokenCache.set(accountId, {
    token: fresh.access_token,
    expiresAt: Date.now() + (fresh.expires_in ?? 3600) * 1000,
  });
  // Microsoft sometimes rotates refresh tokens; capture the new one if so.
  if (fresh.refresh_token && fresh.refresh_token !== auth.refreshToken) {
    auth.refreshToken = fresh.refresh_token;
    authCache.set(accountId, auth);
  }
  return fresh.access_token;
}

/** Exchange a refresh token for a fresh access token. Throws ERR.AUTH on
 *  invalid_grant (the user revoked access or password changed). The
 *  client ID is resolved from the global `oauth.clientID` slot at every
 *  call; the scope is derived from the per-account servertype. */
export async function refreshAccessToken({ refreshToken, servertype }) {
  const clientID = await getGlobalClientID();
  const scope = scopeForServertype(servertype);
  const body = new URLSearchParams({
    client_id: clientID,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope,
  });
  const resp = await fetch(TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  const text = await resp.text().catch(() => "");
  if (!resp.ok) {
    const isInvalidGrant = resp.status === 400 && /invalid_grant/.test(text);
    throw withCode(
      new Error(`Token refresh failed (${resp.status}): ${text}`),
      isInvalidGrant ? ERR.AUTH : ERR.NETWORK,
    );
  }
  try {
    return JSON.parse(text);
  } catch {
    throw withCode(new Error("Invalid token-refresh response"), ERR.NETWORK);
  }
}

/** Pulls the email out of the id_token's payload (when present). The
 *  id_token is a JWT; we don't verify its signature here - we only use
 *  it as a UX hint to display which account was authenticated. */
function decodeIdTokenEmail(idToken) {
  if (!idToken || typeof idToken !== "string") return null;
  const parts = idToken.split(".");
  if (parts.length < 2) return null;
  try {
    // JWT base64url → base64
    const b64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padded = b64 + "===".slice((b64.length + 3) % 4);
    const json = atob(padded);
    const payload = JSON.parse(decodeURIComponent(escape(json)));
    return payload.email ?? payload.preferred_username ?? payload.upn ?? null;
  } catch {
    return null;
  }
}

// ── Popup-driven consent flow (the nativeclient dance) ────────────────────

/**
 * Open `authUrl` in a popup window and resolve with the URL the popup
 * ends up on once Microsoft redirects to the nativeclient endpoint.
 * Throws ERR.CANCELLED if the user closes the window first.
 */
async function runConsentPopup(authUrl) {
  const popup = await browser.windows.create({
    url: authUrl,
    type: "popup",
    width: 500,
    height: 750,
  });

  return new Promise((resolve, reject) => {
    let done = false;

    const cleanup = () => {
      browser.tabs.onUpdated.removeListener(onUpdated);
      browser.windows.onRemoved.removeListener(onClosed);
    };

    const finish = (fn, value) => {
      if (done) return;
      done = true;
      cleanup();
      try {
        browser.windows.remove(popup.id);
      } catch {
        /* already gone */
      }
      fn(value);
    };

    const onUpdated = (_tabId, changeInfo, tab) => {
      if (tab.windowId !== popup.id) return;
      const url = changeInfo.url ?? tab.url;
      if (typeof url === "string" && url.startsWith(REDIRECT_URI)) {
        finish(resolve, url);
      }
    };

    const onClosed = (windowId) => {
      if (windowId !== popup.id) return;
      finish(reject, withCode(new Error("Sign-in cancelled"), ERR.CANCELLED));
    };

    browser.tabs.onUpdated.addListener(onUpdated);
    browser.windows.onRemoved.addListener(onClosed);
  });
}
