/**
 * Per-request `DefaultAnchorMailbox` cookie injection.
 *
 * Microsoft's load balancer at `eas.outlook.com` needs an explicit
 * `Cookie: DefaultAnchorMailbox=<user-email>` value on every EAS
 * request to route the call to the correct backend mailbox tenant.
 *
 * Two complications:
 *
 *  1. The legacy provider relied on Firefox's container infrastructure
 *     (`userContextId`) to keep one account's cookie jar isolated from
 *     another's. That XPCOM-only mechanism isn't reachable from a
 *     MailExtension background page. A naive `browser.cookies.set` into
 *     the shared jar races between concurrent syncs of two `personal-ms`
 *     accounts (autosync fires both in parallel - see the host's
 *     `onAutosyncTick`).
 *
 *  2. Microsoft sends the cookie back with `SameSite=None` but no
 *     `Secure` partition key, so Firefox drops it; we have to inject it
 *     ourselves.
 *
 * The fix: `easRequest` / `easOptions` stamp the per-request mailbox
 * value into a private `X-EAS-Anchor-Mailbox` header before fetch. This
 * `webRequest.onBeforeSendHeaders` listener picks it up, removes the
 * marker, and writes a `Cookie: DefaultAnchorMailbox=<value>` header on
 * the outbound request. No shared cookie-jar state, no race window,
 * one cookie per request scoped to its issuing account.
 */

/** Hosts where Microsoft's load balancer demands the anchor-mailbox
 *  cookie. Other EAS servers don't need it (and ignore it if sent). */
export const ANCHOR_MAILBOX_HOSTS = new Set(["eas.outlook.com"]);

/** Private request marker. The listener strips this before send so it
 *  never reaches the wire. */
export const ANCHOR_MAILBOX_MARKER = "X-EAS-Anchor-Mailbox";

const FILTER_URLS = Array.from(ANCHOR_MAILBOX_HOSTS).map(
  (h) => `https://${h}/*`,
);

export function installAnchorMailboxInjector() {
  if (!browser.webRequest?.onBeforeSendHeaders) {
    console.warn(
      "[eas-4-tbsync] webRequest API unavailable; anchor-mailbox cookie injection disabled",
    );
    return;
  }
  browser.webRequest.onBeforeSendHeaders.addListener(
    rewriteHeaders,
    { urls: FILTER_URLS },
    ["blocking", "requestHeaders"],
  );
}

function rewriteHeaders(details) {
  const headers = details.requestHeaders;
  if (!Array.isArray(headers)) return {};

  let mailbox = null;
  const filtered = [];
  for (const h of headers) {
    if (
      h.name &&
      h.name.toLowerCase() === ANCHOR_MAILBOX_MARKER.toLowerCase()
    ) {
      mailbox = h.value ?? null;
    } else {
      filtered.push(h);
    }
  }
  // Not one of our requests → pass through unchanged. (Other extensions
  // or the user's own browsing wouldn't carry the marker, and we don't
  // want to invent cookie state for them.)
  if (!mailbox) return {};

  // Replace any pre-existing DefaultAnchorMailbox in the Cookie header
  // (defensive - should not happen now that we no longer set the cookie
  // ourselves, but another extension might).
  const ours = `DefaultAnchorMailbox=${mailbox}`;
  const cookieHeader = filtered.find((h) => h.name?.toLowerCase() === "cookie");
  if (cookieHeader) {
    const stripped = String(cookieHeader.value || "")
      .split(/;\s*/)
      .filter((p) => p && !/^DefaultAnchorMailbox=/i.test(p))
      .join("; ");
    cookieHeader.value = stripped ? `${stripped}; ${ours}` : ours;
  } else {
    filtered.push({ name: "Cookie", value: ours });
  }
  return { requestHeaders: filtered };
}
