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
 *  [modules/wbxml.mjs](../wbxml.mjs) — turns "Feiertage%20in%20Deutschland"
 *  back into "Feiertage in Deutschland". All leaf text in the synthesized
 *  XML went through `encodeURIComponent`, so a single decode here is the
 *  right inverse for everything (numeric tokens like "1"/"8"/"13" round-trip
 *  unchanged; multi-byte UTF-8 is reassembled correctly). */
function decodeText(text) {
  if (text == null) return null;
  if (text === "") return "";
  try { return decodeURIComponent(text); }
  catch { return text; }
}

function childByTag(node, tag) {
  for (const c of node.children) {
    if (c.tagName === tag) return c;
  }
  return null;
}
