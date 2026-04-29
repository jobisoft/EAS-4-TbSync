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

function childByTag(node, tag) {
  for (const c of node.children) {
    if (c.tagName === tag) return c;
  }
  return null;
}
