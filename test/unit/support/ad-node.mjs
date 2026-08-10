/**
 * Turn a wire capture - an XML snippet pasted VERBATIM from the Event
 * Log's decoded WBXML, e.g. one `<ApplicationData>…</ApplicationData>`
 * or `<Exception>…</Exception>` element - into the lightweight
 * `{tagName, textContent, children}` node shape that wbxml-helpers'
 * `childByTag` / `readPathFrom` / `readChildTexts` consume. A user's log
 * capture becomes a regression fixture unchanged, which is the point:
 * any mismatch with what the real decoder produces shows up as a failing
 * fixture, never as a silently divergent hand-built one.
 *
 * Deliberately NOT a general XML parser, so it can stay small enough to
 * trust by reading: namespace prefixes are stripped from tag names,
 * attributes are skipped wholesale (EAS carries everything as element
 * text; the codec never reads an attribute), and there is no entity,
 * CDATA or comment handling - decoded-WBXML captures contain none of
 * those, their values travel percent-encoded and are undone by
 * wbxml-helpers' own `decodeText`. Feeding it anything fancier than a
 * capture is a bug in the test, and it throws rather than guesses.
 *
 * Wrapper elements get NO `textContent` - production wrappers do carry
 * DOM-concatenated descendant text, but nothing ever reads a wrapper
 * through `decodeText`, and omitting it keeps `readPathFrom` on the
 * `childByTag` path a capture is meant to exercise. A leaf's text is
 * preserved exactly; an empty leaf (`<A/>` or `<A></A>`) yields `""`,
 * matching what the real decoder produces for every empty element.
 * Non-whitespace text BESIDE child elements throws - captures never
 * contain it, and throwing is also what surfaces a `>` hiding inside an
 * attribute value instead of silently mis-parsing the capture.
 *
 * Idea and shape from PR #345 (tomaskovacik); reimplemented here for
 * the node:test layer.
 */

export function parseAdNode(xml) {
  const src = String(xml)
    .replace(/^\s*<\?xml[^?]*\?>\s*/i, "")
    .trim();
  let i = 0;

  const fail = (what) => {
    throw new Error(
      `parseAdNode: ${what} at offset ${i} (…${src.slice(Math.max(0, i - 20), i + 20)}…)`,
    );
  };

  function readTagName() {
    const start = i;
    while (i < src.length && !/[\s/>]/.test(src[i])) i++;
    if (i === start) fail("empty tag name");
    const raw = src.slice(start, i);
    const colon = raw.lastIndexOf(":");
    return colon === -1 ? raw : raw.slice(colon + 1);
  }

  function parseElement() {
    if (src[i] !== "<") fail("expected '<'");
    i++;
    const tagName = readTagName();
    // Skip attributes. Quoted values in captures never contain '>'.
    while (i < src.length && src[i] !== ">") i++;
    if (i >= src.length) fail("unterminated tag");
    const selfClosing = src[i - 1] === "/";
    i++; // consume '>'
    if (selfClosing) return { tagName, textContent: "", children: [] };

    const children = [];
    let text = "";
    while (i < src.length) {
      if (src[i] === "<") {
        if (src[i + 1] === "/") {
          i += 2;
          const closing = readTagName();
          if (closing !== tagName) {
            fail(`mismatched </${closing}> for <${tagName}>`);
          }
          while (i < src.length && src[i] !== ">") i++;
          i++;
          if (children.length) {
            if (text.trim() !== "") {
              fail(`text beside child elements in <${tagName}>`);
            }
            return { tagName, children };
          }
          return { tagName, textContent: text, children };
        }
        children.push(parseElement());
      } else {
        text += src[i];
        i++;
      }
    }
    fail(`missing </${tagName}>`);
  }

  const node = parseElement();
  if (i < src.length && src.slice(i).trim() !== "") {
    fail("trailing content after the root element");
  }
  return node;
}

/** Build a node by hand where no capture exists: `el("Subject", "x")`
 *  is a leaf, `el("Exceptions", [children])` a wrapper. Same shape as
 *  `parseAdNode`'s output. */
export function el(tagName, value) {
  return Array.isArray(value)
    ? { tagName, children: value }
    : { tagName, textContent: value, children: [] };
}
