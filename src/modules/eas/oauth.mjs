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
 * account?" (see `isOAuthAccount`).
 *
 * In-memory state is one **session** per account. A session is a sign-in
 * lineage: Microsoft rotates the refresh token on every redemption, so one
 * consent produces a chain of tokens, and the current access token, the
 * refresh in flight, and the stored refresh token are all facts about that
 * chain. The invariant the shape enforces: nothing from one lineage can
 * leak into another. Callers run concurrently - the GAL search and the
 * free/busy lookup fire one request per keystroke - and every historical
 * bug in this file was such a leak: two callers redeeming the same
 * refresh token, a stale storage snapshot overwriting a rotated one, a
 * late-landing refresh burying the token a re-sign-in had just stored.
 *
 * Hence: a session is created whole (`primeAuth`, which seeds and never
 * replaces), discarded whole (`forgetAuth` - the account was disabled,
 * deleted, or re-authenticated, the cases where storage really is newer),
 * and written only by its own refresh, which holds the object itself. A
 * refresh that outlives its session writes into an orphan nobody can
 * reach. `primeAuth` runs at the top of every on* hook that hits the
 * network, and per request from the GAL and free/busy paths.
 *
 * A rotation is announced through `setRotationSink` so it reaches
 * `account.custom` whichever caller triggered it - the background is an
 * event page, and a suspend takes these sessions with it, so the next wake
 * primes from storage.
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
 * Global storage.local key that routes consent through the system browser
 * instead of the popup below. Absent or anything but `true` means the
 * popup, which is the right answer for almost everybody: it keeps the
 * sign-in inside Thunderbird and needs no copying and pasting.
 *
 * The popup is a Thunderbird content window, and Thunderbird's chrome
 * implements no WebAuthn prompt, so a tenant that requires a passkey or a
 * security key cannot finish consent there - the ceremony waits for a UI
 * nobody shows, and the page falls back to asking for a password the
 * tenant has disabled. Thunderbird's own mail accounts escape the same
 * trap through `mailnews.oauth.useExternalBrowser`, but that pref is read
 * by `OAuth2.sys.mjs`, which serves accounts this add-on does not own;
 * setting it changes nothing here. This option is the equivalent for EAS
 * accounts, deliberately separate so that a profile can keep the smoother
 * popup for mail while an EAS account that needs a passkey goes out.
 */
const USE_EXTERNAL_BROWSER_STORAGE_KEY = "oauth.useExternalBrowser";

/** True iff the user asked for the external-browser route. Storage being
 *  unreachable answers "no": the popup is the route that always works, and
 *  an unreadable option is not a reason to demand a paste. */
async function useExternalBrowser() {
  try {
    const rv = await browser.storage.local.get({
      [USE_EXTERNAL_BROWSER_STORAGE_KEY]: false,
    });
    return rv[USE_EXTERNAL_BROWSER_STORAGE_KEY] === true;
  } catch {
    return false;
  }
}

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

/** accountId → session. See the module docstring: a session is one
 *  sign-in lineage, created whole, discarded whole, and written only by
 *  its own refresh. Shape:
 *
 *    { refreshToken, servertype,   // the lineage's current link + scope
 *      accessToken, expiresAt,     // what the lineage has redeemed
 *      pending }                   // the one refresh in flight, or null
 */
const sessions = new Map();

/** Called with (accountId, refreshToken) when the live session's token
 *  rotates, so the new one can reach storage. Set once by the provider;
 *  absent in tests and before construction. */
let rotationSink = null;

export function setRotationSink(fn) {
  rotationSink = typeof fn === "function" ? fn : null;
}

/** Open a session for this account from what storage holds - unless one
 *  is already live. An existing session has redeemed its way past the
 *  stored token, so replacing it would revive a link the server has
 *  already retired; and the concurrent callers this module serves each
 *  read storage at their own moment, so stale snapshots arriving here is
 *  the normal case, not a corner. `forgetAuth` is how a caller says the
 *  stored token really is the newer one. */
export function primeAuth(accountId, { refreshToken, servertype }) {
  if (!refreshToken || !servertype) return;
  if (sessions.has(accountId)) return;
  sessions.set(accountId, {
    refreshToken,
    servertype,
    accessToken: null,
    expiresAt: 0,
    pending: null,
  });
}

/** End the account's session - lineage, access token, in-flight slot, all
 *  of it. A refresh still in the air keeps its orphaned object and settles
 *  for its own callers, but can no longer be seen through the map. */
export function forgetAuth(accountId) {
  sessions.delete(accountId);
}

/** Drop the redeemed access token but keep the lineage, so the next call
 *  refreshes. A refresh already in flight is left alone on purpose: it is
 *  fetching a new token, which is exactly what a caller who just saw a
 *  401 wants to join. */
export function invalidateAccessToken(accountId) {
  const session = sessions.get(accountId);
  if (!session) return;
  session.accessToken = null;
  session.expiresAt = 0;
}

/** Hand the session an access token obtained elsewhere - the interactive
 *  sign-in returns one, and keeping it saves the next request a refresh
 *  round-trip. Without a session this is a no-op: an access token with no
 *  lineage behind it could never be renewed. */
export function primeAccessToken(accountId, token, expiresIn) {
  const session = sessions.get(accountId);
  if (!session) return;
  session.accessToken = token;
  session.expiresAt = Date.now() + (expiresIn ?? 3600) * 1000;
}

// ── Interactive sign-in ────────────────────────────────────────────────────

/**
 * Open the Microsoft consent popup and return the resulting tokens.
 *
 *   loginHint        pre-selects an account on the consent screen
 *   servertype       "office365" | "personal-ms" - drives scope
 *   onWindowCreated  called with the popup's windowId once it exists, so a
 *                    caller can hand it to `registerAccountWindow` and let
 *                    the host raise the popup if the user clicks again
 *
 * The OAuth client ID is read from the global `oauth.clientID` slot
 * (storage.local), with the hardcoded community ID as fallback. */
export async function startAuth({ loginHint, servertype, onWindowCreated }) {
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

  const responseUrl = (await useExternalBrowser())
    ? await runExternalConsent(authUrl.toString(), state, onWindowCreated)
    : await runConsentPopup(authUrl.toString(), onWindowCreated);

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

/** Returns a valid access token for the account's session, refreshing
 *  transparently when the one in hand is expired or missing. `primeAuth`
 *  must have opened a session first.
 *
 *  At most one refresh per session is ever in flight; concurrent callers
 *  join it and share its answer. Everything the refresh learns is written
 *  to the session it started from - never through the map - so a refresh
 *  that lands after its session ended settles its own callers and changes
 *  nothing else. */
export async function getAccessToken(accountId) {
  const session = sessions.get(accountId);
  if (!session) {
    throw withCode(
      new Error("OAuth auth not primed - call primeAuth first"),
      ERR.AUTH,
    );
  }
  if (session.accessToken && session.expiresAt > Date.now() + REFRESH_SKEW_MS) {
    return session.accessToken;
  }
  session.pending ??= (async () => {
    try {
      const fresh = await refreshAccessToken(session);
      session.accessToken = fresh.access_token;
      session.expiresAt = Date.now() + (fresh.expires_in ?? 3600) * 1000;
      // Microsoft sometimes rotates refresh tokens; move the lineage on.
      if (fresh.refresh_token && fresh.refresh_token !== session.refreshToken) {
        session.refreshToken = fresh.refresh_token;
        // Only the live session speaks to storage: writing a rotation that
        // belongs to an ended lineage would leave `custom.refreshToken` on
        // the dead chain, and the next start would prime from it. Not
        // awaited, and a rejection is caught rather than left to the
        // process - the sink is an async function handed in from outside,
        // and one that threw would slip past a plain try/catch.
        if (sessions.get(accountId) === session) {
          try {
            Promise.resolve(rotationSink?.(accountId, fresh.refresh_token)).catch(
              (err) =>
                console.debug("[eas] refresh-token rotation sink failed:", err),
            );
          } catch (err) {
            console.debug("[eas] refresh-token rotation sink threw:", err);
          }
        }
      }
      return fresh.access_token;
    } finally {
      session.pending = null;
    }
  })();
  return session.pending;
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
async function runConsentPopup(authUrl, onWindowCreated) {
  const popup = await browser.windows.create({
    url: authUrl,
    type: "popup",
    width: 500,
    height: 750,
  });

  // Reported before the await below, so a caller tracking this window has
  // it registered for the whole time the popup is up. Never let a failing
  // callback take down a sign-in that is otherwise fine.
  try {
    onWindowCreated?.(popup.id);
  } catch (err) {
    console.debug("[eas] onWindowCreated callback failed:", err);
  }

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
      } catch (err) {
        // Window already gone — common when the user closes it manually.
        console.debug(
          `[eas] OAuth popup windows.remove(${popup.id}) failed:`,
          err,
        );
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

// ── External-browser consent (the paste dance) ────────────────────────────

/**
 * The other route, for tenants whose sign-in the popup above cannot
 * perform - see USE_EXTERNAL_BROWSER_STORAGE_KEY. The authorization page
 * goes to the system browser, where a passkey works because the browser
 * has the WebAuthn UI Thunderbird lacks, and where the user may already
 * hold a session and never be asked anything.
 *
 * What that costs is the watcher. `tabs.onUpdated` sees Thunderbird's own
 * tabs and nothing else, so the redirect - a blank `…/nativeclient?code=…`
 * page in a foreign application - has to come back by hand: the user
 * copies the address bar into a Thunderbird dialog. Thunderbird's own
 * external flow avoids the paste by listening on a loopback socket and
 * redirecting there instead, which needs `http://localhost` registered on
 * the Azure application; the community client ID this add-on ships is
 * registered for `nativeclient` only, so that door is shut for anyone who
 * cannot register an application of their own.
 *
 * The dialog is not trusted to be right, only to be the user's. Whether a
 * pasted URL may end a sign-in is decided here, in the same terms
 * `runConsentPopup` decides it for a navigation, and a URL that may not is
 * refused with a reason rather than an exception - the dialog stays up and
 * the user tries again.
 */
const PASTE_DIALOG_PATH = "dialogs/oauth-paste/oauth-paste.html";

/** handoff token → the sign-in waiting on it. One entry per live external
 *  consent; the token is what the dialog quotes back, so a window left
 *  over from an abandoned attempt cannot answer for the current one. */
const pendingExternal = new Map();

async function runExternalConsent(authUrl, state, onWindowCreated) {
  const token = crypto.randomUUID();
  // Registered before the window exists. The dialog can message us the
  // moment it loads, and a handoff we have not recorded yet reads as
  // expired - which would strand a sign-in that is perfectly fine.
  const entry = { state, authUrl, windowId: null, resolve: null, reject: null };
  pendingExternal.set(token, entry);

  let resolveResult;
  let rejectResult;
  const result = new Promise((resolve, reject) => {
    resolveResult = resolve;
    rejectResult = reject;
  });

  // One settlement per handoff, whoever gets there first: the paste, the
  // Cancel button, the window closing, or a browser that would not open.
  let done = false;
  const onClosed = (windowId) => {
    if (windowId !== entry.windowId) return;
    entry.reject(withCode(new Error("Sign-in cancelled"), ERR.CANCELLED));
  };
  const settle = (fn, value) => {
    if (done) return;
    done = true;
    pendingExternal.delete(token);
    browser.windows.onRemoved.removeListener(onClosed);
    if (entry.windowId != null) {
      // Already gone when the user is the one who closed it; that is the
      // common case, not a fault.
      Promise.resolve(browser.windows.remove(entry.windowId)).catch((err) =>
        console.debug(
          `[eas] OAuth paste dialog windows.remove(${entry.windowId}) failed:`,
          err,
        ),
      );
    }
    fn(value);
  };
  entry.resolve = (responseUrl) => settle(resolveResult, responseUrl);
  entry.reject = (err) => settle(rejectResult, err);
  browser.windows.onRemoved.addListener(onClosed);

  try {
    const url = new URL(browser.runtime.getURL(PASTE_DIALOG_PATH));
    url.searchParams.set("token", token);
    // The dialog goes up first. It is the only way back from a browser
    // Thunderbird cannot see, so a failure to open it must not leave the
    // user signing in with nowhere to put the result.
    const dialog = await browser.windows.create({
      url: url.toString(),
      type: "popup",
      width: 560,
      height: 480,
    });
    entry.windowId = dialog.id;
    // Reported before the browser opens, so a caller tracking this window
    // has it for the whole time it is up. Never let a failing callback take
    // down a sign-in that is otherwise fine.
    try {
      onWindowCreated?.(dialog.id);
    } catch (err) {
      console.debug("[eas] onWindowCreated callback failed:", err);
    }
    await browser.windows.openDefaultBrowser(authUrl);
  } catch (err) {
    entry.reject(
      withCode(
        new Error(
          `Could not start sign-in in your default browser: ${err?.message ?? String(err)}`,
        ),
        ERR.AUTH,
      ),
    );
  }

  return result;
}

/**
 * The paste dialog offering a URL as the end of its sign-in. Returns a
 * verdict rather than throwing: everything a user can get wrong here is
 * ordinary - the wrong window copied, a link from an attempt they
 * abandoned, the page copied before sign-in finished - and each deserves
 * its own sentence in the dialog rather than a failed sign-in.
 *
 * `accepted` means the URL was handed to the waiting sign-in, not that the
 * sign-in succeeded: a redirect carrying Microsoft's own `error` is
 * accepted here and fails there, where the message it carries can be
 * reported properly.
 */
export async function completeExternalConsent({ token, url } = {}) {
  const entry = pendingExternal.get(token);
  if (!entry) return { accepted: false, reason: "expired" };

  let parsed;
  try {
    parsed = new URL(String(url ?? "").trim());
  } catch {
    return { accepted: false, reason: "notAUrl" };
  }
  if (!parsed.href.startsWith(REDIRECT_URI)) {
    return { accepted: false, reason: "notRedirect" };
  }
  if (!parsed.searchParams.get("code") && !parsed.searchParams.get("error")) {
    return { accepted: false, reason: "noCode" };
  }
  // The same CSRF check `startAuth` makes, made early so a link from an
  // earlier attempt is a sentence in the dialog rather than a dead sign-in.
  if (parsed.searchParams.get("state") !== entry.state) {
    return { accepted: false, reason: "staleLink" };
  }

  entry.resolve(parsed.href);
  return { accepted: true };
}

/** Send the user to the authorization page again - the browser did not
 *  come up, or they closed the tab before finishing. The URL is the one
 *  handed out the first time on purpose: a freshly built one would carry a
 *  new `state`, and the paste the user is midway through would be refused
 *  as stale. */
export async function reopenExternalConsent({ token } = {}) {
  const entry = pendingExternal.get(token);
  if (!entry) return { accepted: false, reason: "expired" };
  try {
    await browser.windows.openDefaultBrowser(entry.authUrl);
  } catch (err) {
    console.debug("[eas] could not reopen the default browser:", err);
    return { accepted: false, reason: "noBrowser" };
  }
  return { accepted: true };
}

/** The dialog's Cancel button. Closing the window does the same thing;
 *  this is the deterministic way, since a content page is not always
 *  allowed to close the window it lives in. */
export async function cancelExternalConsent({ token } = {}) {
  const entry = pendingExternal.get(token);
  if (!entry) return { accepted: false, reason: "expired" };
  entry.reject(withCode(new Error("Sign-in cancelled"), ERR.CANCELLED));
  return { accepted: true };
}
