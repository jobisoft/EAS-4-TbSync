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
 * We request HTML from the server (`BodyPreference Type=2` for calendar
 * and tasks), so a Type-2 payload carries the server's real formatting; we
 * store it in the ALTREP and derive the plain value with Thunderbird's own
 * converter. A Type-1 payload is authoritative plain text and clears any
 * ALTREP. Outbound mirrors it: a note with an ALTREP goes back as Type 2.
 */

import { readPathFrom } from "./wbxml-helpers.mjs";

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
  { useAirSyncBase },
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
    // HTML: the value is a plain-text rendering (the tooltip's source),
    // the HTML rides along as the ALTREP (the editor's source).
    const text = await messenger.messengerUtilities.convertToPlainText(
      data,
      {},
    );
    comp.updatePropertyWithValue("description", asText(text));
    comp
      .getFirstProperty("description")
      ?.setParameter(ALTREP, HTML_ALTREP_PREFIX + encodeURIComponent(data));
    return;
  }

  // Plain text is the whole truth - clear any ALTREP a prior in-Thunderbird
  // edit left behind, or the editor would keep showing that stale HTML.
  comp.updatePropertyWithValue("description", data);
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

  // 16.1 omits an empty Body rather than sending a blank one.
  if (asVersion === "16.1" && !value && !html) return;
  if (asVersion === "2.5") {
    // 2.5 has no AirSyncBase Body / HTML - plain text only.
    builder.atag("Body", value);
    return;
  }

  const [type, out] = html ? ["2", html] : ["1", value];
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
