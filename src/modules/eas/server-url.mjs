/**
 * Turn what the user typed into the URL a custom-mode account contacts.
 *
 * Only `servertype === "custom"` holds a user-entered address. Every
 * other mode composes or receives a complete URL - the OAuth hosts are
 * built from a constant, `auto` stores what Autodiscover returned, and a
 * migrated v4 account is assembled by `liftHostAndHttpsToServer` - so
 * nothing here applies to them.
 *
 * v4 had a `host` field plus an HTTPS checkbox and composed the URL at
 * request time, which meant a bare hostname was always a valid answer.
 * v5 replaced both with one text field and dropped the composition, so
 * `my.server.tld` stopped working (issue #338). Both of v4's composing
 * rules are restored verbatim:
 *
 *   - a missing scheme means https (createAccount.js:251-254)
 *   - the path ends at /Microsoft-Server-ActiveSync unless it already
 *     does, compared case-sensitively (network.js:103-104)
 *
 * Where this deliberately departs from v4: a user name, password, query
 * string or fragment is refused rather than carried along. v4 kept all
 * of them, because `stripAutodiscoverUrl` reduced the typed URL with
 * `u.split("//")[1]` and stored whatever followed the scheme. None of
 * them have a meaning for an ActiveSync endpoint - credentials belong in
 * the user and password fields, and `easRequest` builds the query string
 * itself - so accepting them only defers a confusing failure.
 */

const ENDPOINT_PATH = "/Microsoft-Server-ActiveSync";
const HAS_SCHEME = /^[a-z][a-z0-9+.-]*:\/\//i;

/**
 * @param {string} input raw `custom.server`
 * @returns {string|null} the URL to contact, or null if unusable
 */
export function normalizeCustomServerUrl(input) {
  const trimmed = String(input ?? "").trim();
  if (!trimmed) return null;

  // A bare host is shorthand, not an error - it was the only thing v4's
  // host field could hold.
  const withScheme = HAS_SCHEME.test(trimmed) ? trimmed : `https://${trimmed}`;

  let parsed;
  try {
    parsed = new URL(withScheme);
  } catch {
    return null;
  }
  if (!parsed.hostname) return null;
  if (parsed.protocol !== "https:" && parsed.protocol !== "http:") return null;
  // Prefixing `https://` above turns the local part of an email address
  // into a user name, so this is also what catches `me@example.com`
  // typed into the server field instead of the user field.
  if (parsed.username || parsed.password) return null;
  if (parsed.search || parsed.hash) return null;

  let path = parsed.pathname;
  while (path.endsWith("/")) path = path.slice(0, -1);
  if (!path.endsWith(ENDPOINT_PATH)) path += ENDPOINT_PATH;

  return parsed.origin + path;
}
