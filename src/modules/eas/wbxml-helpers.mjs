/**
 * Tiny shared helpers for navigating decoded EAS responses. The decoded
 * Document loses the WBXML namespace structure (each codepage's tags
 * become plain elements), so a path-anchored child walk is the safest
 * way to read fields when the same tag name appears at multiple depths
 * (e.g. `Status` exists under both `Provision` and `Provision.Policies.Policy`).
 */

/** Walk from the document root down through `path` (an array of tag names),
 *  matching direct children at each step. Returns the trimmed text content
 *  of the leaf, or null if any step doesn't match. */
export function readPath(doc, path) {
  let node = doc?.documentElement;
  if (!node) return null;
  for (const tag of path) {
    node = childByTag(node, tag);
    if (!node) return null;
  }
  return decodeText(node.textContent);
}

/** Like `readPath` but starts at the given element rather than the root.
 *  Handy when the caller has already located a parent (e.g. each `Add`
 *  inside `FolderSync.Changes`). */
export function readPathFrom(node, path) {
  for (const tag of path) {
    if (!node) return null;
    node = childByTag(node, tag);
  }
  return decodeText(node?.textContent);
}

/** Return the decoded text of every direct child of `node` whose tag
 *  matches `tag`. Used for repeating-element containers like
 *  `<Categories><Category>foo</Category>…</Categories>` and
 *  `<Children><Child>…</Child>…</Children>`. Goes through `decodeText`
 *  so the wbxml decoder's `encodeURIComponent` escape and the per-byte
 *  UTF-8 reinterpretation both unwind correctly. Empty / whitespace-only
 *  results are skipped. */
export function readChildTexts(node, tag) {
  const result = [];
  if (!node?.children) return result;
  for (const c of node.children) {
    if (c.tagName === tag) {
      const v = decodeText(c.textContent);
      if (v) result.push(v);
    }
  }
  return result;
}

/** Inverse of the WBXML decoder's `encodeURIComponent` round-trip in
 *  [modules/wbxml.mjs](../wbxml.mjs). The decoder builds a string with
 *  one JS code unit per raw WBXML byte, then `encodeURIComponent`s it
 *  for safe XML embedding. `decodeURIComponent` here recovers that
 *  byte-per-code-unit string; we then interpret the byte sequence as
 *  UTF-8 (per MS-ASWBXML 2.1.2.2) to get the proper Unicode string.
 *  Without the second step "ü" (UTF-8 0xC3 0xBC) would surface as
 *  "Ã¼". ASCII bytes are unchanged by either step. */
function decodeText(text) {
  if (text == null) return null;
  if (text === "") return "";
  let raw;
  try {
    raw = decodeURIComponent(text);
  } catch {
    return text;
  }
  try {
    const bytes = Uint8Array.from(raw, (c) => c.charCodeAt(0) & 0xff);
    return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
  } catch {
    return raw;
  }
}

/** The first child element with this tag, or null. Takes a missing node
 *  so callers can walk a path without a guard at every step. */
export function childByTag(node, tag) {
  if (!node?.children) return null;
  for (const c of node.children) {
    if (c.tagName === tag) return c;
  }
  return null;
}

/**
 * Is this EAS date value the Windows FILETIME zero?
 *
 * FILETIME counts 100-nanosecond intervals from 1601-01-01T00:00:00Z, so a
 * field that was never set serialises as exactly that instant. Kerio
 * Connect sends it for a contact's Anniversary that has never been set,
 * and for a Birthday the user has just cleared - which is how a cleared
 * birthday came back as 1 January 1601 and could never be cleared again.
 *
 * ActiveSync has no sentinel for "no value": an unset element is simply
 * omitted, so a value that can only be a null leaking through
 * serialisation is read as absent.
 *
 * The whole day is matched, not the instant alone. A server rendering the
 * epoch in its own zone produces 1600-12-31T23:00:00Z or 1601-01-01T01:00Z
 * rather than the epoch itself, and those are the same non-value.
 */
export function isFiletimeZero(value) {
  if (!value) return false;
  const m = /^(\d{4})-?(\d{2})-?(\d{2})/.exec(String(value));
  if (!m) return false;
  const [, y, mo, d] = m;
  const day = `${y}${mo}${d}`;
  return day === "16010101" || day === "16001231";
}
