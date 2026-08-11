/**
 * The AirSyncBase Body ⇆ iCal DESCRIPTION mapping, shared by the calendar
 * and task codecs (VEVENT and VTODO both carry notes the same way).
 *
 * A rich note lives on the DESCRIPTION property: the plain text is its
 * value, the HTML is an `ALTREP="data:text/html,…"` parameter beside it -
 * the shape Thunderbird's own `descriptionHTML` setter writes and its
 * editor reads (`CalItemBase.sys.mjs`). The tooltip reads only the value.
 * So the two must agree: whenever the value changes, the ALTREP has to be
 * rewritten or cleared, or the editor keeps rendering a stale note while
 * the tooltip shows the new one (issue #347).
 *
 * The Body reaching this module is one the sync runner has already judged:
 * it reads NativeBodyType, which says what the server actually stores, and
 * re-fetches an item held as HTML before decoding it. So a Type-1 payload
 * here means the note really is plain text - it becomes the value and any
 * ALTREP goes - and a Type-2 payload really is the note's own markup, which
 * becomes the ALTREP beside a plain-text rendering for the tooltip.
 *
 * Outbound is where formatting travels: a DESCRIPTION carrying an ALTREP
 * goes out as Type 2 with that HTML, so what the user authored in
 * Thunderbird reaches the server; a note without one goes out as Type 1.
 */

import { readPathFrom } from "./wbxml-helpers.mjs";
import {
  asNoteText,
  rememberNoteLineEndings,
  restoreNoteLineEndings,
} from "./note-text.mjs";

const HTML_ALTREP_PREFIX = "data:text/html,";

// ical.js normalises parameter names to lower case on parse and looks them
// up case-sensitively, so every access has to use "altrep" - "ALTREP"
// silently misses on any blob that has been through ICAL.parse, which is
// every stored one. ical.js still serialises it as the canonical ALTREP.
const ALTREP = "altrep";

const asText = (v) => (v == null ? "" : String(v));

/** Read the inbound `<Body>` into `comp`'s DESCRIPTION, ALTREP and all.
 *  Merge-aware: an absent (or empty) Body leaves the note - value AND its
 *  ALTREP - untouched, so a delta that does not mention the note changes
 *  nothing. `useAirSyncBase` is false only on 2.5, where Body is a plain
 *  Calendar-page element with no Type (always plain text). Async: the
 *  HTML→text conversion is a WebExtension API. */
export async function readBodyIntoDescription(
  comp,
  adNode,
  { useAirSyncBase, nativePlainText = null },
) {
  const data = useAirSyncBase
    ? readPathFrom(adNode, ["Body", "Data"])
    : readPathFrom(adNode, ["Body"]);
  if (!data) return;

  const type = useAirSyncBase ? readPathFrom(adNode, ["Body", "Type"]) : "1";

  // Only Type 2 is HTML. Everything else - Type 1, and the RTF (3) / MIME
  // (4) forms we never request and a conformant server therefore never
  // sends - is treated as plain text: value verbatim, ALTREP cleared.
  if (type === "2") {
    // HTML: the value is a plain-text rendering (the tooltip's source), the
    // HTML rides along as the ALTREP (the editor's source). The rendering is
    // the server's own when the sync runner was given one - text we were
    // handed rather than computed, and what other clients show. Converting
    // locally is the fallback for HTML that arrived with no plain
    // counterpart.
    const text =
      nativePlainText ??
      (await messenger.messengerUtilities.convertToPlainText(data, {}));
    // The server's line endings are recorded here as well as on the plain
    // branch. The ALTREP is what travels while it exists, so this changes
    // nothing today - but the moment a user clears the formatting, the value
    // becomes what goes back, and it has to go back in the server's shape
    // like any other note. Read from whichever payload the server sent.
    rememberNoteLineEndings(comp, nativePlainText ?? data);
    // Trailing whitespace comes off, and only here: this value is what the
    // tooltip shows while the ALTREP beside it is what travels, so trimming
    // it edits nothing the server compares against. The plain branch below
    // leaves its value alone - there it IS the note.
    comp.updatePropertyWithValue(
      "description",
      asNoteText(text).replace(/\s+$/, ""),
    );
    comp
      .getFirstProperty("description")
      ?.setParameter(ALTREP, HTML_ALTREP_PREFIX + encodeURIComponent(data));
    return;
  }

  // Plain text. The sync runner re-fetches an item whose NativeBodyType says
  // HTML, so this is normally the note itself rather than a flattening of
  // one. Normally, not always: an embedded exception carries no
  // NativeBodyType of its own, and the revert path asks for plain text
  // deliberately. Both readers get it - the value for the tooltip, and
  // no ALTREP so the editor falls back to that same text. Leaving a stale
  // ALTREP here is what made the editor and the tooltip disagree (#347).
  rememberNoteLineEndings(comp, data);
  comp.updatePropertyWithValue("description", asNoteText(data));
  comp.getFirstProperty("description")?.removeParameter(ALTREP);
}

/** Emit `<Body>` from `comp`'s DESCRIPTION. A note carrying an HTML ALTREP
 *  goes out as Type 2 with that HTML, so formatting survives the round
 *  trip; a plain note goes out as Type 1. `homePage` is the codepage the
 *  caller is on and must be returned to (Calendar / Tasks). */
export function appendBodyFromDescription(builder, comp, asVersion, homePage) {
  const prop = comp.getFirstProperty("description");
  const value = prop ? asText(prop.getFirstValue()) : "";
  const altrep = prop?.getParameter(ALTREP);
  // An ALTREP we did not write - an item imported from elsewhere - may not
  // be valid percent-encoding, and `decodeURIComponent` throws on that.
  // Thrown from here it would escape the whole push and block every other
  // pending edit in the folder, so one unreadable note would cost the user
  // all of them. The note's plain text still goes out as Type 1.
  let html = null;
  if (typeof altrep === "string" && altrep.startsWith(HTML_ALTREP_PREFIX)) {
    try {
      html = decodeURIComponent(altrep.slice(HTML_ALTREP_PREFIX.length));
    } catch {
      html = null;
    }
  }

  // A plain note goes back in the line endings the server uses. The two
  // steps are one operation: collapse whatever the value holds to bare
  // newlines - including CRLFs a user just typed, which would otherwise be
  // doubled - then expand every newline once. Idempotent, so a value that is
  // already CRLF survives unchanged. The HTML branch needs none of this: its
  // ALTREP is stored exactly as received and travels the same way.
  const plain = restoreNoteLineEndings(comp, value);

  // 16.1 omits an empty Body rather than sending a blank one.
  if (asVersion === "16.1" && !value && !html) return;
  if (asVersion === "2.5") {
    // 2.5 has no AirSyncBase Body / HTML - plain text only.
    builder.atag("Body", plain);
    return;
  }

  const [type, out] = html ? ["2", html] : ["1", plain];
  builder.switchpage("AirSyncBase");
  builder.otag("Body");
  builder.atag("Type", type);
  if (asVersion !== "16.1") {
    builder.atag("EstimatedDataSize", String(out.length));
  }
  builder.atag("Data", out);
  builder.ctag();
  builder.switchpage(homePage);
}
