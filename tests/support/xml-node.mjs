/**
 * Minimal recursive-descent parser turning a captured EAS wire XML
 * snippet - pasted verbatim from the TbSync debug log, e.g. a single
 * `<ApplicationData>...</ApplicationData>` or `<Exception>...</Exception>`
 * element - into the lightweight `{tagName, textContent, children}` node
 * shape that `wbxml-helpers.mjs`'s `childByTag`/`readPathFrom`/
 * `readChildTexts` consume.
 *
 * This is deliberately NOT a general XML parser: no entities, no CDATA,
 * no comments, attribute values are skipped rather than parsed (the
 * codec never reads them - EAS carries everything as element text,
 * percent-encoded the same way `decodeText` in wbxml-helpers.mjs
 * expects). That narrow scope is what makes it safe to hand-roll instead
 * of pulling in a real XML/DOM dependency: it only has to match what the
 * real WBXML decoder already hands the codec, and real captured payloads
 * are exactly what we test against, so any mismatch shows up immediately
 * as a failing fixture rather than a silent divergence.
 */

export function parseAdNode(xml) {
  const src = xml.replace(/^\s*<\?xml[^?]*\?>\s*/i, "").trim();
  let i = 0;

  function skipWs() {
    while (i < src.length && /\s/.test(src[i])) i++;
  }

  function readTagName() {
    const start = i;
    while (i < src.length && !/[\s/>]/.test(src[i])) i++;
    const raw = src.slice(start, i);
    const colon = raw.lastIndexOf(":");
    return colon === -1 ? raw : raw.slice(colon + 1);
  }

  function parseElement() {
    if (src[i] !== "<") {
      throw new Error(`parseAdNode: expected '<' at offset ${i}`);
    }
    i++;
    const tagName = readTagName();
    skipWs();
    while (i < src.length && src[i] !== ">") i++;
    const selfClosing = src[i - 1] === "/";
    i++; // consume '>'
    if (selfClosing) {
      return { tagName, textContent: "", children: [] };
    }
    const children = [];
    let text = "";
    while (i < src.length) {
      if (src[i] === "<") {
        if (src[i + 1] === "/") {
          i += 2;
          readTagName();
          skipWs();
          if (src[i] === ">") i++;
          break;
        }
        children.push(parseElement());
      } else {
        text += src[i];
        i++;
      }
    }
    return { tagName, textContent: text, children };
  }

  skipWs();
  return parseElement();
}
