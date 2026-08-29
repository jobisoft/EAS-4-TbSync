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
import {
  getAccessToken,
  invalidateAccessToken,
  isOAuthAccount,
} from "./eas/oauth.mjs";
import { reportEventLog } from "./eas-event-log.mjs";
import {
  ANCHOR_MAILBOX_HOSTS,
  ANCHOR_MAILBOX_MARKER,
} from "./anchor-mailbox.mjs";
import { normalizeCustomServerUrl } from "./eas/server-url.mjs";

const DEFAULT_USER_AGENT = "Thunderbird ActiveSync";
const CUSTOM_USER_AGENT_STORAGE_KEY = "tbsync.useragent";

const DEFAULT_DEVICE_TYPE = "TbSync";
const CUSTOM_DEVICE_TYPE_STORAGE_KEY = "tbsync.type";

const DEFAULT_DEVICE_OS = "Unknown";
const CUSTOM_DEVICE_OS_STORAGE_KEY = "tbsync.os";

// Map browser.runtime.getPlatformInfo()'s `os` enum to the legacy
// Services.appinfo.OS values. Servers that key on the OS string (e.g.
// the mobile-device list in Exchange admin) expect these legacy
// tokens; values from the deprecated `navigator.platform` (Win32 /
// MacIntel / Linux x86_64) are seen as "unknown" by some of them.
// Anything outside the map falls through to the raw enum.
const LEGACY_OS_MAP = {
  win: "WINNT",
  linux: "Linux",
  mac: "Darwin",
  openbsd: "OpenBSD",
  android: "Android",
  cros: "CrOS",
};

const DEFAULT_CONNECTION_TIMEOUT_MS = 90_000;
const CUSTOM_CONNECTION_TIMEOUT_STORAGE_KEY = "timeout";

/** Stable error codes thrown by this module. */
export const NET_ERR = {
  AUTH: "E:AUTH",
  PROVISION_REQUIRED: "E:PROVISION_REQUIRED",
  HOST_REDIRECT: "E:HOST_REDIRECT",
  HTTP: "E:HTTP",
  NETWORK: "E:NETWORK",
  // Its own code rather than E:NETWORK: the host treats E:NETWORK as
  // predefined and renders that code instead of our message, which is
  // the only thing naming the offending address. It also keeps the
  // account clear of the Autodiscover re-run E:NETWORK triggers.
  INVALID_SERVER_URL: "E:INVALID_SERVER_URL",
};

export class EasHttpError extends Error {
  constructor(code, status, options = {}) {
    // Messages from this class can surface in the manager UI (the host's
    // sync-coordinator falls back to `err.message` for codes outside its
    // PREDEFINED_ERROR_CODES set, which includes E:HTTP / E:HOST_REDIRECT
    // / E:PROVISION_REQUIRED). Default message is provider-localized via
    // `eas.network.error.transport`; callers may override with a more
    // specific localized message.
    super(
      options.message ??
        browser.i18n.getMessage("eas.network.error.transport", [
          String(status),
        ]),
      { cause: options.cause },
    );
    this.name = "EasHttpError";
    this.code = code;
    this.status = status;
    if (options.newLocation) this.newLocation = options.newLocation;
    if (options.retryAfterMs != null) this.retryAfterMs = options.retryAfterMs;
  }
}

/** First four bytes of every EAS WBXML response: version 1.3, public id 1,
 *  UTF-8, empty string table. Anything else is junk (HTML error page, JSON
 *  blob from a misconfigured server, etc.) and we reject it before feeding
 *  it to the decoder. */
const WBXML_MAGIC = [0x03, 0x01, 0x6a, 0x00];

/** The URL to contact. Only custom-mode accounts hold an address the
 *  user typed, so only they are normalized - every other mode stores a
 *  complete URL and is passed through exactly as before.
 *
 *  Checked here as well as in the dialogs because storage can hold
 *  anything: the migration writes `custom.server`, so do the provider's
 *  `createAccountFromSetup` and `saveAccountFromConfig`, neither of
 *  which validates, and a profile can arrive from a backup. The dialogs
 *  are one writer among several, so the value is verified where it is
 *  used rather than only where it is entered. */
function easUrlFor(custom, context) {
  if (!custom?.server)
    throw new Error(`${context}: account.custom.server is missing`);
  if (custom.servertype !== "custom") return custom.server;

  const url = normalizeCustomServerUrl(custom.server);
  if (url) return url;
  throw new EasHttpError(NET_ERR.INVALID_SERVER_URL, 0, {
    message: browser.i18n.getMessage("setup.error.serverInvalid", [
      String(custom.server),
    ]),
  });
}

// ── Public API ────────────────────────────────────────────────────────────

/**
 * One OPTIONS request, logged like every other exchange.
 *
 * Request headers are deliberately not logged - `Authorization` is one of
 * them. Response headers are, because the version and command lists live
 * there, and because a server that fails here usually names its own trace
 * id (`X-Request-Id` and friends), which is what lets its administrator
 * find the failure from the other side.
 */
async function sendOptions({ account }) {
  const custom = account?.custom ?? {};
  const serverUrl = easUrlFor(custom, "easOptions");

  const headers = new Headers({
    Authorization: await buildAuthHeader(account),
    "User-Agent": await getUserAgent(),
  });
  stampAnchorMailbox(headers, custom);
  reportEventLog({
    level: "debug",
    accountId: account?.accountId,
    message: "[eas:net] send OPTIONS",
    details: serverUrl,
  });

  const { resp } = await fetchWithTimeout(
    serverUrl,
    { method: "OPTIONS", headers },
    syncSignalFor(account?.accountId),
  );
  let buf = null;
  try {
    buf = new Uint8Array(await resp.arrayBuffer());
  } catch {
    /* a body we cannot read is not a reason to lose the status */
  }
  reportEventLog({
    level: "debug",
    accountId: account?.accountId,
    message: `[eas:net] receive OPTIONS (HTTP ${resp.status})`,
    details: formatHeadersForLog(resp.headers),
  });
  if (!resp.ok) logRecvError({ account, command: "OPTIONS", resp, buf });

  return {
    resp,
    versions: parseList(resp.headers.get("MS-ASProtocolVersions")),
    commands: parseList(resp.headers.get("MS-ASProtocolCommands")),
  };
}

export async function easOptions({ account }) {
  const custom = account?.custom ?? {};
  const serverUrl = easUrlFor(custom, "easOptions");
  // OPTIONS runs once per connect, so this is the one place that can
  // name the effective URL without repeating it on every request.
  if (serverUrl !== custom.server) {
    reportEventLog({
      level: "info",
      accountId: account?.accountId,
      message: browser.i18n.getMessage("eas.network.info.effectiveServerUrl", [
        serverUrl,
      ]),
    });
  }
  const { resp, versions, commands } = await sendOptions({ account });
  if (resp.status === 401 || resp.status === 403) {
    throw new EasHttpError(NET_ERR.AUTH, resp.status);
  }
  if (resp.status === 451) throw redirectError(resp);
  if (!resp.ok) throw new EasHttpError(NET_ERR.HTTP, resp.status);
  return { versions, commands };
}

/** How this module reaches the running sync's AbortSignal.
 *
 *  Wired once by the provider rather than threaded through eight call sites -
 *  and threading would have missed the ones that matter, since Provision,
 *  Settings and FolderSync all run inside a sync too. */
let syncSignalFor = () => null;
export function setSyncSignalResolver(fn) {
  syncSignalFor = typeof fn === "function" ? fn : () => null;
}

export async function easRequest({ account, command, body, asVersion }) {
  const custom = account?.custom ?? {};
  const serverUrl = easUrlFor(custom, "easRequest");
  if (!custom.user)
    throw new Error("easRequest: account.custom.user is missing");
  if (!custom.deviceId)
    throw new Error("easRequest: account.custom.deviceId is missing");

  const url = new URL(serverUrl);
  url.searchParams.set("Cmd", command);
  url.searchParams.set("User", custom.user);
  url.searchParams.set("DeviceId", custom.deviceId);
  url.searchParams.set("DeviceType", await getDeviceType());

  const send = async (authHeader, retryOnAuth) => {
    const headers = new Headers({
      "Content-Type": "application/vnd.ms-sync.wbxml",
      Authorization: authHeader,
      "MS-ASProtocolVersion": asVersion ?? custom.asversion ?? "14.1",
      "User-Agent": await getUserAgent(),
    });
    stampAnchorMailbox(headers, custom);
    // Once the user (or a server-driven 449) has flipped `provision: true`,
    // legacy sends `X-MS-PolicyKey` on every command - including the
    // bootstrap value `"0"` during the very first Provision request.
    // Omitting it is what trips up some servers during Provision iter 0.
    if (custom.provision === true) {
      headers.set("X-MS-PolicyKey", custom.policykey ?? "0");
    }

    logSendXML({ account, command, body });
    const { resp, buf: rawBuf } = await fetchWithTimeout(
      url,
      { method: "POST", headers, body },
      syncSignalFor(account?.accountId),
      { readBody: true },
    );

    // Log what a failing response actually said - headers and body - at
    // debug, before the throw ladder decides what the failure means. The
    // retried 401 lands here too; one debug line for a recovered auth blip
    // is cheap, and its absence made real prompts undiagnosable.
    if (!resp.ok) logRecvError({ account, command, resp, buf: rawBuf });

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
    if (resp.status === 449)
      throw new EasHttpError(NET_ERR.PROVISION_REQUIRED, 449);
    if (resp.status === 451) throw redirectError(resp);
    if (resp.status === 503) {
      // "Retry later", by the book: [MS-ASCMD] 2.2.2 has pre-14.0 servers
      // answer HTTP 503 where 14.0+ answers Status 111
      // (ServerErrorRetryLater), and Exchange throttling uses it too. Carry
      // the server's own Retry-After when it states one, so the sync layer
      // can pause autosync for the time the server asked for, not a guess.
      throw new EasHttpError(NET_ERR.HTTP, 503, {
        retryAfterMs: parseRetryAfterMs(resp.headers.get("Retry-After")),
      });
    }
    if (!resp.ok) throw new EasHttpError(NET_ERR.HTTP, resp.status);

    const buf = rawBuf ?? new Uint8Array(0);
    if (buf.length === 0) return { xml: "", doc: null };

    if (!hasWbxmlMagic(buf)) {
      logRecvText({ account, command, status: resp.status, buf });
      const head = [...buf.slice(0, 4)]
        .map((b) => b.toString(16).padStart(2, "0"))
        .join(" ");
      throw new EasHttpError(NET_ERR.HTTP, resp.status, {
        message: browser.i18n.getMessage("eas.network.error.responseNotWbxml", [
          head,
        ]),
      });
    }

    const xml = decodeWBXML(buf);
    logRecvXML({ account, command, xml });
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    return { xml, doc };
  };

  const authHeader = await buildAuthHeader(account);
  return send(authHeader, /* retryOnAuth */ true);
}

/* ── Wire-level debug logging ───────────────────────────────────────── */

// Every WBXML send and receive emits a `level: "debug"` event-log entry
// with the decoded XML in `details`. The host applies its own capture
// gate (`tbsync.settings.logLevel`) and silently drops these unless the
// user has enabled debug-level capture, so the calls are unconditional
// here. Decoding the outbound body is a tiny cost the host's drop path
// accepts in exchange for keeping the wire layer free of host config.

function logSendXML({ account, command, body }) {
  let xml;
  try {
    xml = decodeWBXML(body);
  } catch {
    xml = "<decode-failed>";
  }
  reportEventLog({
    level: "debug",
    accountId: account?.accountId,
    message: `[eas:net] send ${command}`,
    details: xml,
  });
}

function logRecvXML({ account, command, xml }) {
  reportEventLog({
    level: "debug",
    accountId: account?.accountId,
    message: `[eas:net] receive ${command}`,
    details: xml,
  });
}

/** Response headers as "Name: value" lines, for the error log.
 *
 *  Set-Cookie is withheld - a session cookie in an event log is a
 *  credential in every bug report the log is attached to. Everything else
 *  goes through: Retry-After, X-MS-*, WWW-Authenticate and whatever else
 *  the server volunteers are exactly what a report needs. */
export function formatHeadersForLog(headers) {
  const lines = [];
  for (const [name, value] of headers?.entries?.() ?? []) {
    if (name.toLowerCase() === "set-cookie") continue;
    lines.push(`${name}: ${value}`);
  }
  return lines.join("\n");
}

/** One entry for a response we are about to throw on. Until this existed,
 *  every receive log line required a healthy response - the throw ladder
 *  ran first - so the reports that most needed the wire (an HTTP 500 with
 *  the server's own explanation in the body, a 503 with a Retry-After)
 *  carried none of it.
 *
 *  Info, not debug like the other wire lines: those fire on every healthy
 *  request and are most of the log's bytes, while a failing response is
 *  rare and exactly what a bug report needs - so it must be captured at
 *  the default log level, without asking the reporter to raise verbosity
 *  and fail again.
 *
 *  Logged whole. The log holds one record per entry and rolls by entry
 *  count, so a long body costs nothing a short one does not - and a
 *  truncated server error page is exactly the one that stops explaining
 *  itself at the interesting line. */
function logRecvError({ account, command, resp, buf }) {
  let body = "";
  if (buf?.length) {
    try {
      body = new TextDecoder("utf-8", { fatal: false }).decode(buf);
    } catch {
      body = "<decode-failed>";
    }
  }
  reportEventLog({
    level: "info",
    accountId: account?.accountId,
    message: `[eas:net] receive ${command} failed (HTTP ${resp.status})`,
    details: [formatHeadersForLog(resp.headers), body]
      .filter(Boolean)
      .join("\n\n"),
  });
}

function logRecvText({ account, command, status, buf }) {
  let text;
  try {
    text = new TextDecoder("utf-8", { fatal: false }).decode(buf);
  } catch {
    text = "<decode-failed>";
  }
  reportEventLog({
    level: "debug",
    accountId: account?.accountId,
    message: `[eas:net] receive ${command} non-WBXML (HTTP ${status})`,
    details: text,
  });
}

async function getDeviceType() {
  try {
    const rv = await browser.storage.local.get({
      [CUSTOM_DEVICE_TYPE_STORAGE_KEY]: "",
    });
    const v = rv[CUSTOM_DEVICE_TYPE_STORAGE_KEY];
    if (typeof v === "string" && v.trim()) {
      return v.trim();
    }
  } catch {
    /* fall through */
  }
  return DEFAULT_DEVICE_TYPE;
}

export async function getUserAgent() {
  try {
    const rv = await browser.storage.local.get({
      [CUSTOM_USER_AGENT_STORAGE_KEY]: "",
    });
    const v = rv[CUSTOM_USER_AGENT_STORAGE_KEY];
    if (typeof v === "string" && v.trim()) {
      return v.trim();
    }
  } catch {
    /* fall through */
  }
  return DEFAULT_USER_AGENT;
}

/** Platform-derived Device OS string (no storage override). Exported so
 *  the options dialog can populate the input's placeholder with the
 *  value the user would inherit if they leave the override empty. */
export async function getDefaultDeviceOs() {
  try {
    const { os } = await browser.runtime.getPlatformInfo();
    return LEGACY_OS_MAP[os] ?? os ?? DEFAULT_DEVICE_OS;
  } catch {
    return DEFAULT_DEVICE_OS;
  }
}

/** Effective Device OS sent on the wire. Storage override wins;
 *  otherwise falls back to the platform-mapped default. */
export async function getDeviceOs() {
  try {
    const rv = await browser.storage.local.get({
      [CUSTOM_DEVICE_OS_STORAGE_KEY]: "",
    });
    const v = rv[CUSTOM_DEVICE_OS_STORAGE_KEY];
    if (typeof v === "string" && v.trim()) {
      return v.trim();
    }
  } catch {
    /* fall through */
  }
  return getDefaultDeviceOs();
}

async function getConnectionTimeout() {
  try {
    const rv = await browser.storage.local.get({
      [CUSTOM_CONNECTION_TIMEOUT_STORAGE_KEY]: null,
    });
    const v = rv[CUSTOM_CONNECTION_TIMEOUT_STORAGE_KEY];
    if (typeof v === "number" && Number.isFinite(v) && v > 0) {
      return v;
    }
  } catch {
    /* fall through */
  }
  return DEFAULT_CONNECTION_TIMEOUT_MS;
}

/** Stamp the per-request `X-EAS-Anchor-Mailbox` marker that the
 *  webRequest listener in `anchor-mailbox.mjs` translates into a
 *  `Cookie: DefaultAnchorMailbox=<user>` header on the wire. Doing this
 *  per-request avoids the shared cookie-jar race that would otherwise
 *  let two `personal-ms` accounts overwrite each other's anchor value
 *  during concurrent autosync. Hosts not in `ANCHOR_MAILBOX_HOSTS` are
 *  left alone - other EAS servers don't need (and ignore) the cookie. */
function stampAnchorMailbox(headers, custom) {
  if (!custom?.server || !custom.user) return;
  let host;
  try {
    host = new URL(custom.server).hostname;
  } catch {
    return;
  }
  if (!ANCHOR_MAILBOX_HOSTS.has(host)) return;
  headers.set(ANCHOR_MAILBOX_MARKER, custom.user);
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
  if (!custom.user)
    throw new Error("buildAuthHeader: account.custom.user is missing");
  if (custom.password == null)
    throw new Error("buildAuthHeader: account.custom.password is missing");
  return basicAuthHeader(custom.user, custom.password);
}

/** How long to pause autosync on a retry-later signal when the server
 *  does not say. Legacy paused 30 minutes on Sync Status 110; kept. */
export const RETRY_LATER_BACKOFF_MS = 30 * 60 * 1000;

/** The Retry-After header, as milliseconds from now, or null.
 *
 *  RFC 9110 allows delay-seconds or an HTTP-date; both occur in the wild.
 *  Clamped to [1 min, 4 h]: below a minute a pause is not worth
 *  recording, and above four hours a confused server would silence an
 *  account for a day on one header nobody can see. */
export function parseRetryAfterMs(header) {
  if (!header) return null;
  const s = header.trim();
  let ms = null;
  if (/^\d+$/.test(s)) {
    ms = Number(s) * 1000;
  } else {
    const date = Date.parse(s);
    if (!Number.isNaN(date)) ms = date - Date.now();
  }
  if (ms == null) return null;
  return Math.min(Math.max(ms, 60_000), 4 * 60 * 60 * 1000);
}

// ── Helpers ───────────────────────────────────────────────────────────────

/** Wrap fetch in an AbortController so we don't hang forever on a black-
 *  hole connection. Retries once on a transient network error before
 *  giving up - covers brief Wi-Fi drops, DNS hiccups, server bounces.
 *  AbortError (timeout) and other fetch errors both map to E:NETWORK. */
const NETWORK_RETRY_DELAY_MS = 500;

async function fetchWithTimeout(url, init, cancelSignal = null, opts = {}) {
  const timeout = await getConnectionTimeout();
  // Read the body inside the guarded window when asked. Headers arriving
  // says nothing about the body: a server can answer 200 and then stall the
  // stream, and a read outside this loop would hang past both the timeout
  // and the cancel signal. (Before the cancel work, the timer happened to
  // stay armed across the caller's read and covered this by accident.)
  const { readBody = false } = opts;

  for (let attempt = 0; attempt < 2; attempt++) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeout);
    // The host asking us to stop has to reach the socket, not just the loop
    // between requests: a server that never answers is exactly the case the
    // Disconnect button exists for, and waiting out the connection timeout
    // first would make the button feel broken.
    const onCancel = () => controller.abort();
    if (cancelSignal) {
      if (cancelSignal.aborted) controller.abort();
      else cancelSignal.addEventListener("abort", onCancel, { once: true });
    }
    try {
      const resp = await fetch(url, { ...init, signal: controller.signal });
      if (!readBody) return { resp, buf: null };
      // Error responses are read too: a 500 from a non-Microsoft server
      // usually carries its own explanation as HTML or text, and dropping
      // it unread is why error reports used to arrive with nothing but the
      // status number. A failed body read on an error response is not
      // worth failing over - the status is already in hand.
      let buf = null;
      try {
        buf = new Uint8Array(await resp.arrayBuffer());
      } catch (err) {
        if (resp.ok) throw err;
      }
      return { resp, buf };
    } catch (err) {
      clearTimeout(timer);
      // A cancellation is not a network failure and must not be retried or
      // reported as one. Rethrown as-is so the AbortError name survives to
      // the port, where it becomes E:CANCELLED rather than E:PROVIDER_FAULT.
      if (cancelSignal?.aborted) throw err;
      const isTimeout = err.name === "AbortError";
      if (attempt === 0) {
        // Brief pause before the retry so we don't immediately re-hit a
        // half-closed socket.
        await new Promise((r) => setTimeout(r, NETWORK_RETRY_DELAY_MS));
        continue;
      }
      if (isTimeout) {
        throw new EasHttpError(NET_ERR.NETWORK, 0, {
          message: "Connection timeout",
        });
      }
      throw new EasHttpError(NET_ERR.NETWORK, 0, { cause: err });
    } finally {
      clearTimeout(timer);
      // Every request of a sync shares one signal, so a listener left behind
      // on the success path accumulates for the length of the sync.
      cancelSignal?.removeEventListener("abort", onCancel);
    }
  }
  // Unreachable: the loop either returns or throws.
  throw new EasHttpError(NET_ERR.NETWORK, 0, {
    message: browser.i18n.getMessage("eas.network.error.retriesExhausted"),
  });
}

function redirectError(resp) {
  const newLocation = resp.headers.get("X-MS-Location");
  if (!newLocation) {
    // Per MS-ASHTTP, a 451 MUST carry an X-MS-Location pointing to the
    // new EAS endpoint. Without it the response is malformed - fail loud
    // as a generic HTTP error rather than dressing it up as a redirect
    // that the caller can't act on anyway.
    return new EasHttpError(NET_ERR.HTTP, 451, {
      message: browser.i18n.getMessage(
        "eas.network.error.redirectMissingLocation",
      ),
    });
  }
  return new EasHttpError(NET_ERR.HOST_REDIRECT, 451, {
    message: `Server moved to ${newLocation}`,
    newLocation,
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
  // Dedupe: some EAS frontends (notably Office 365) emit
  // MS-ASProtocolCommands twice in the OPTIONS reply, which the Headers
  // API joins with ",". The list is a set in spirit.
  return [
    ...new Set(
      headerValue
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean),
    ),
  ];
}
