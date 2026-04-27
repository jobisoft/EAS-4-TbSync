/**
 * WBXML 1.3 codec for EAS traffic.
 *
 * Decode: Uint8Array (binary WBXML from server) → XML string the caller
 *   parses with DOMParser to query. Inline strings are URI-component
 *   encoded so any byte value round-trips as printable XML text.
 *
 * Encode: `createWBXML(initialNamespace)` returns a builder with the
 *   methods `switchpage(name)`, `otag(tag)`, `atag(tag, content?)`,
 *   `ctag()`, `append(bytes)`, `getBytes()`. Builder writes to an
 *   auto-growing byte buffer; `getBytes()` snapshots to Uint8Array.
 *
 * Ported from legacy EAS-4-TbSync `content/includes/wbxmltools.js`
 * (MPL-2.0) with these adjustments:
 *   - binary I/O is Uint8Array, not charcode-as-string;
 *   - inline string encoding uses TextEncoder (UTF-8 bytes);
 *   - inline string decoding preserves raw bytes via encodeURIComponent
 *     on a byte-string synthesized from the bytes (matches legacy behaviour).
 */

import { CODEPAGES, NAMESPACES, NAMESPACE_INDEX, TOKENS_BY_NAME } from "./wbxml-codepages.mjs";

const TEXT_ENCODER = new TextEncoder();

/** WBXML header: 0x03 version 1.3, 0x01 unknown public id, 0x6A UTF-8
 *  charset, 0x00 empty string table. */
const HEADER = new Uint8Array([0x03, 0x01, 0x6A, 0x00]);

// ── Decode ────────────────────────────────────────────────────────────────

/**
 * Decode a WBXML byte buffer into an XML string. The first four bytes are
 * the standard EAS WBXML header (fixed: version 1.3, public id 1, UTF-8,
 * empty string table). Inline-string content is `encodeURIComponent`-
 * escaped so the returned XML parses without fiddling with binary bytes.
 */
export function decodeWBXML(bytes) {
  if (!(bytes instanceof Uint8Array)) {
    bytes = new Uint8Array(bytes);
  }
  let pos = 4; // skip fixed 4-byte header
  let codepage = 0;
  let mainCodepage = null;
  const tagStack = [];
  let xml = "";

  while (pos < bytes.length) {
    const data = bytes[pos];
    const token = data & 0x3F;
    const hasContent = (data & 0x40) !== 0;
    const hasAttributes = (data & 0x80) !== 0;

    switch (token) {
      case 0x00: // SWITCH_PAGE: next byte = new codepage index
        pos += 1;
        codepage = bytes[pos] & 0xFF;
        break;

      case 0x01: // END: close most-recently-opened tag
        xml += tagStack.pop();
        break;

      case 0x02: // ENTITY: unsupported in EAS traffic
        throw new Error("wbxml: ENTITY token is not supported");

      case 0x03: { // STR_I: inline string, NUL-terminated
        let end = pos + 1;
        while (end < bytes.length && bytes[end] !== 0x00) end++;
        // Preserve original byte values through encodeURIComponent by
        // building a code-unit-per-byte string. apostrophe isn't encoded
        // by encodeURIComponent, so we patch it by hand.
        let s = "";
        for (let i = pos + 1; i < end; i++) s += String.fromCharCode(bytes[i]);
        xml += encodeURIComponent(s).replace(/'/g, "%27");
        pos = end;
        break;
      }

      // Unsupported global tokens (not used by EAS):
      case 0x04: case 0x40: case 0x41: case 0x42: case 0x43: case 0x44:
      case 0x80: case 0x81: case 0x82: case 0x83: case 0x84:
      case 0xC0: case 0xC1: case 0xC2: case 0xC3: case 0xC4:
        throw new Error(`wbxml: global token 0x${token.toString(16)} not supported`);

      default: {
        // Regular tag. Emit namespace attribute the first time we see a
        // non-main codepage; the main codepage attaches once on the root.
        const needNs = codepage !== mainCodepage;
        const ns = needNs ? ` xmlns='${getNamespace(codepage)}'` : "";
        if (mainCodepage === null) mainCodepage = codepage;

        const tagName = getTagName(codepage, token);
        if (!hasContent) {
          xml += `<${tagName}${ns}/>`;
        } else {
          xml += `<${tagName}${ns}>`;
          tagStack.push(`</${tagName}>`);
        }
        // We don't emit a dedicated "unknown token" diagnostic here;
        // `getTagName` falls back to a synthetic label so downstream
        // queries can still find the element.
        break;
      }
    }
    pos += 1;
    // `hasAttributes` is set but we don't encounter attribute lists in
    // EAS traffic (all server responses use content-only tokens). If a
    // server ever sends one, we'd need an attribute-parsing branch here.
    void hasAttributes;
  }

  return xml === "" ? "" : `<?xml version="1.0" encoding="utf-8"?>${xml}`;
}

function getNamespace(codepage) {
  return NAMESPACES[codepage] ?? `UnknownCodePage${codepage}`;
}

function getTagName(codepage, token) {
  const page = CODEPAGES[codepage];
  if (page && token in page) return page[token];
  return `Unknown.${codepage}.${token}`;
}

// ── Encode (builder) ──────────────────────────────────────────────────────

/**
 * Build a WBXML byte buffer incrementally. The builder holds a growable
 * array of bytes; call `getBytes()` to snapshot into a `Uint8Array`
 * suitable for a fetch body.
 *
 * Usage:
 *   const w = createWBXML("FolderHierarchy");
 *   w.otag("FolderSync");
 *     w.atag("SyncKey", "0");
 *   w.ctag();
 *   return w.getBytes();
 */
export function createWBXML(initialNamespace = "") {
  /** @type {number[]} */
  const bytes = [];
  let codepage = 0;

  for (const b of HEADER) bytes.push(b);
  if (initialNamespace) {
    const idx = NAMESPACE_INDEX.get(initialNamespace);
    if (idx === undefined) throw new Error(`wbxml: unknown namespace '${initialNamespace}'`);
    codepage = idx;
  }

  const tokenFor = (name) => {
    const page = TOKENS_BY_NAME[codepage];
    if (!page || !(name in page)) {
      throw new Error(`wbxml: unknown tag '${name}' in codepage '${NAMESPACES[codepage]}'`);
    }
    return page[name];
  };

  return {
    switchpage(name) {
      const idx = NAMESPACE_INDEX.get(name);
      if (idx === undefined) throw new Error(`wbxml: unknown namespace '${name}'`);
      codepage = idx;
      bytes.push(0x00, idx);
    },

    /** Open a tag with content (expects a matching `ctag()`). */
    otag(name) {
      bytes.push(tokenFor(name) | 0x40);
    },

    /** Close the most-recently-opened content tag. */
    ctag() {
      bytes.push(0x01);
    },

    /** Emit a complete `<tag>content</tag>`. With empty content, emits
     *  the short self-closing form `<tag/>`. */
    atag(name, content = "") {
      const token = tokenFor(name);
      if (content === "") {
        bytes.push(token);
        return;
      }
      bytes.push(token | 0x40, 0x03);
      const utf8 = TEXT_ENCODER.encode(String(content));
      for (const b of utf8) bytes.push(b);
      bytes.push(0x00, 0x01);
    },

    /** Append raw WBXML bytes (rarely needed; reserved for edge cases
     *  where opaque sub-documents are spliced in). */
    append(raw) {
      const arr = raw instanceof Uint8Array ? raw : new Uint8Array(raw);
      for (const b of arr) bytes.push(b);
    },

    getBytes() {
      return new Uint8Array(bytes);
    },
  };
}
