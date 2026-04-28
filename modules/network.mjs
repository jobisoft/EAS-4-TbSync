/**
 * EAS HTTP transport. Two functions:
 *
 *   easOptions({account})
 *       OPTIONS probe - discovers the server's MS-ASProtocolVersions and
 *       MS-ASProtocolCommands. Used at first sync to negotiate the AS
 *       version we'll use for subsequent commands. Honours
 *       isOAuthAccount(account.custom) so OAuth accounts use Bearer.
 *
 *   easRequest({account, command, body, asVersion?})
 *       Single-shot WBXML POST. Encodes the EAS query string from the
 *       account's custom fields, attaches basic-auth or Bearer + AS
 *       protocol headers, sends the WBXML body, decodes the response,
 *       and parses it into an XML Document for the caller to query.
 *
 * Failure modes are surfaced as `EasHttpError` with a stable `code` so
 * upstream callers can branch on auth vs provision-required vs generic
 * HTTP errors without re-parsing status fields.
 */

import { decodeWBXML } from "./wbxml.mjs";
import { getAccessToken, invalidateAccessToken, isOAuthAccount } from "./eas/oauth.mjs";

const DEFAULT_USER_AGENT = "Thunderbird ActiveSync";
const CUSTOM_USER_AGENT_STORAGE_KEY = "tbsync.useragent";

const DEFAULT_DEVICE_TYPE = "TbSync";
const CUSTOM_DEVICE_TYPE_STORAGE_KEY = "tbsync.type";

const DEFAULT_CONNECTION_TIMEOUT_MS = 90_000;
const CUSTOM_CONNECTION_TIMEOUT_STORAGE_KEY = "timeout";

/** Hostnames where Microsoft's load balancer needs the
 *  `DefaultAnchorMailbox` cookie to route the request to the right
 *  mailbox tenant. Without it requests bounce around and EAS commands
 *  silently fail. The legacy add-on pre-sets this cookie because
 *  Firefox drops the server's own response cookie (SameSite=None
 *  without a partition key, etc.). */
const ANCHOR_MAILBOX_HOSTS = new Set(["eas.outlook.com"]);

/** Stable error codes thrown by this module. */
export const NET_ERR = {
  AUTH: "E:AUTH",
  PROVISION_REQUIRED: "E:PROVISION_REQUIRED",
  HOST_REDIRECT: "E:HOST_REDIRECT",
  HTTP: "E:HTTP",
  NETWORK: "E:NETWORK",
};

export class EasHttpError extends Error {
  constructor(code, status, options = {}) {
    super(options.message ?? `EAS transport error ${code} (HTTP ${status})`, { cause: options.cause });
    this.name = "EasHttpError";
    this.code = code;
    this.status = status;
    if (options.newLocation) this.newLocation = options.newLocation;
  }
}

/** First four bytes of every EAS WBXML response: version 1.3, public id 1,
 *  UTF-8, empty string table. Anything else is junk (HTML error page, JSON
 *  blob from a misconfigured server, etc.) and we reject it before feeding
 *  it to the decoder. */
const WBXML_MAGIC = [0x03, 0x01, 0x6A, 0x00];

// ── Public API ────────────────────────────────────────────────────────────

export async function easOptions({ account }) {
  const custom = account?.custom ?? {};
  if (!custom.server) throw new Error("easOptions: account.custom.server is missing");
  await ensureAnchorMailboxCookie(account);
  const authHeader = await buildAuthHeader(account);
  const headers = new Headers({
    Authorization: authHeader,
    "User-Agent": await getUserAgent(),
  });
  const resp = await fetchWithTimeout(custom.server, { method: "OPTIONS", headers });
  if (resp.status === 401 || resp.status === 403) {
    throw new EasHttpError(NET_ERR.AUTH, resp.status);
  }
  if (resp.status === 451) throw redirectError(resp);
  if (!resp.ok) throw new EasHttpError(NET_ERR.HTTP, resp.status);
  return {
    versions: parseList(resp.headers.get("MS-ASProtocolVersions")),
    commands: parseList(resp.headers.get("MS-ASProtocolCommands")),
  };
}

export async function easRequest({ account, command, body, asVersion }) {
  const custom = account?.custom ?? {};
  if (!custom.server) throw new Error("easRequest: account.custom.server is missing");
  if (!custom.user) throw new Error("easRequest: account.custom.user is missing");
  if (!custom.deviceId) throw new Error("easRequest: account.custom.deviceId is missing");

  const url = new URL(custom.server);
  url.searchParams.set("Cmd", command);
  url.searchParams.set("User", custom.user);
  url.searchParams.set("DeviceId", custom.deviceId);
  url.searchParams.set("DeviceType", await getDeviceType());

  await ensureAnchorMailboxCookie(account);

  const send = async (authHeader, retryOnAuth) => {
    const headers = new Headers({
      "Content-Type": "application/vnd.ms-sync.wbxml",
      Accept: "application/vnd.ms-sync.wbxml",
      Authorization: authHeader,
      "MS-ASProtocolVersion": asVersion ?? custom.asversion ?? "14.1",
      "User-Agent": await getUserAgent(),
    });
    // Once the user (or a server-driven 449) has flipped `provision: true`,
    // legacy sends `X-MS-PolicyKey` on every command - including the
    // bootstrap value `"0"` during the very first Provision request.
    // Omitting it is what trips up some servers during Provision iter 0.
    if (custom.provision === true) {
      headers.set("X-MS-PolicyKey", custom.policykey ?? "0");
    }

    const resp = await fetchWithTimeout(url, { method: "POST", headers, body });

    if (resp.status === 401 || resp.status === 403) {
      // OAuth-specific recovery: cached access token may be stale despite
      // not being expired (server-side revocation, clock skew). Invalidate
      // and retry once with a freshly-refreshed token before bubbling up
      // an E:AUTH that would disable the account. 403 also belongs here -
      // some servers return it for token-related authorization failures.
      if (retryOnAuth && isOAuthAccount(custom)) {
        invalidateAccessToken(account.accountId);
        const fresh = await buildAuthHeader(account);
        return send(fresh, /* retryOnAuth */ false);
      }
      throw new EasHttpError(NET_ERR.AUTH, resp.status);
    }
    if (resp.status === 449) throw new EasHttpError(NET_ERR.PROVISION_REQUIRED, 449);
    if (resp.status === 451) throw redirectError(resp);
    if (!resp.ok) throw new EasHttpError(NET_ERR.HTTP, resp.status);

    const buf = new Uint8Array(await resp.arrayBuffer());
    if (buf.length === 0) return { xml: "", doc: null };

    if (!hasWbxmlMagic(buf)) {
      const head = [...buf.slice(0, 4)].map(b => b.toString(16).padStart(2, "0")).join(" ");
      throw new EasHttpError(NET_ERR.HTTP, resp.status, {
        message: `Response is not WBXML (first bytes: ${head})`,
      });
    }

    const xml = decodeWBXML(buf);
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    return { xml, doc };
  };

  const authHeader = await buildAuthHeader(account);
  return send(authHeader, /* retryOnAuth */ true);
}

async function getDeviceType() {
  try {
    const rv = await browser.storage.local.get({ [CUSTOM_DEVICE_TYPE_STORAGE_KEY]: "" });
    const v = rv[CUSTOM_DEVICE_TYPE_STORAGE_KEY];
    if (typeof v === "string" && v.trim()) {
      return v.trim();
    }
  } catch { /* fall through */ }
  return DEFAULT_DEVICE_TYPE;
}

async function getUserAgent() {
  try {
    const rv = await browser.storage.local.get({ [CUSTOM_USER_AGENT_STORAGE_KEY]: "" });
    const v = rv[CUSTOM_USER_AGENT_STORAGE_KEY];
    if (typeof v === "string" && v.trim()) {
      return v.trim();
    }
  } catch { /* fall through */ }
  return DEFAULT_USER_AGENT;
}

async function getConnectionTimeout() {
  try {
    const rv = await browser.storage.local.get({ [CUSTOM_CONNECTION_TIMEOUT_STORAGE_KEY]: null });
    const v = rv[CUSTOM_CONNECTION_TIMEOUT_STORAGE_KEY];
    if (typeof v === "number" && Number.isFinite(v) && v > 0) {
      return v;
    }
  } catch { /* fall through */ }
  return DEFAULT_CONNECTION_TIMEOUT_MS;
}

/** Pre-set the `DefaultAnchorMailbox` cookie that Microsoft's load
 *  balancer needs to route EAS requests to the right mailbox. Only
 *  applies to hosts in `ANCHOR_MAILBOX_HOSTS` (currently just
 *  `eas.outlook.com`). The cookie value is the user's email - same
 *  thing the legacy add-on does at network.js:69-84. The browser will
 *  attach the cookie automatically on the next fetch to that host.
 *  Failures are swallowed: a missing `cookies` permission shouldn't
 *  prevent sync against servers that don't need this cookie. */
async function ensureAnchorMailboxCookie(account) {
  const custom = account?.custom ?? {};
  if (!custom.server || !custom.user) return;
  let host;
  try { host = new URL(custom.server).hostname; }
  catch { return; }
  if (!ANCHOR_MAILBOX_HOSTS.has(host)) return;
  try {
    await browser.cookies.set({
      url: `https://${host}/`,
      name: "DefaultAnchorMailbox",
      value: custom.user,
      domain: host,
      path: "/",
      secure: true,
      httpOnly: false,
      sameSite: "no_restriction",
      // Far-future expiry; the cookie is harmless to leave around and
      // setting it on every request keeps the value aligned if the
      // user's email changes mid-account.
      expirationDate: Math.floor(Date.now() / 1000) + (365 * 24 * 3600),
    });
  } catch (err) {
    console.warn("[eas-4-tbsync] cookies.set DefaultAnchorMailbox failed:", err?.message ?? err);
  }
}

/** Build the Authorization header for a given account. OAuth accounts
 *  use a cached Bearer (refreshed transparently by `getAccessToken`);
 *  basic-auth accounts concatenate user:password. */
async function buildAuthHeader(account) {
  const custom = account?.custom ?? {};
  if (isOAuthAccount(custom)) {
    const token = await getAccessToken(account.accountId);
    return `Bearer ${token}`;
  }
  if (!custom.user) throw new Error("buildAuthHeader: account.custom.user is missing");
  if (custom.password == null) throw new Error("buildAuthHeader: account.custom.password is missing");
  return basicAuthHeader(custom.user, custom.password);
}

// ── Helpers ───────────────────────────────────────────────────────────────

/** Wrap fetch in an AbortController so we don't hang forever on a black-
 *  hole connection. Retries once on a transient network error before
 *  giving up - covers brief Wi-Fi drops, DNS hiccups, server bounces.
 *  AbortError (timeout) and other fetch errors both map to E:NETWORK. */
const NETWORK_RETRY_DELAY_MS = 500;

async function fetchWithTimeout(url, init) {
  const timeout = await getConnectionTimeout();

  for (let attempt = 0; attempt < 2; attempt++) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeout);
    try {
      return await fetch(url, { ...init, signal: controller.signal });
    } catch (err) {
      clearTimeout(timer);
      const isTimeout = err.name === "AbortError";
      if (attempt === 0) {
        // Brief pause before the retry so we don't immediately re-hit a
        // half-closed socket.
        await new Promise(r => setTimeout(r, NETWORK_RETRY_DELAY_MS));
        continue;
      }
      if (isTimeout) {
        throw new EasHttpError(NET_ERR.NETWORK, 0, { message: "Connection timeout" });
      }
      throw new EasHttpError(NET_ERR.NETWORK, 0, { cause: err });
    } finally {
      clearTimeout(timer);
    }
  }
  // Unreachable: the loop either returns or throws.
  throw new EasHttpError(NET_ERR.NETWORK, 0, { message: "fetchWithTimeout: exhausted retries" });
}

function redirectError(resp) {
  const newLocation = resp.headers.get("X-MS-Location");
  return new EasHttpError(NET_ERR.HOST_REDIRECT, 451, {
    message: newLocation ? `Server moved to ${newLocation}` : "Server requested a redirect (no X-MS-Location)",
    newLocation: newLocation ?? null,
  });
}

function hasWbxmlMagic(buf) {
  if (buf.length < WBXML_MAGIC.length) return false;
  for (let i = 0; i < WBXML_MAGIC.length; i++) {
    if (buf[i] !== WBXML_MAGIC[i]) return false;
  }
  return true;
}

/** RFC 7617 basic-auth header. UTF-8 → byte string → base64; this avoids
 *  the well-known `btoa` problem with non-ASCII characters in the
 *  username or password. */
function basicAuthHeader(user, password) {
  const utf8 = new TextEncoder().encode(`${user}:${password}`);
  let bin = "";
  for (const b of utf8) bin += String.fromCharCode(b);
  return "Basic " + btoa(bin);
}

function parseList(headerValue) {
  if (!headerValue) return [];
  return headerValue.split(",").map(s => s.trim()).filter(Boolean);
}
