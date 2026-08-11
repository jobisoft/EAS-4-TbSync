/**
 * Line endings for a note, shared by every codec that carries one.
 *
 * EAS servers send CRLF: Exchange separates the lines of a note with one and
 * ends a flattened note with one. Neither iCalendar nor vCard can hold a
 * carriage return in a TEXT value - RFC 5545 §3.3.11 and RFC 6350 §3.4 give
 * an escape for a line break and none for a CR - and ical.js writes one
 * straight through, so the property ends mid-value for any reader that
 * splits on CR and the rest of the note becomes a line belonging to nothing.
 * That is issue #262, and a blob Thunderbird itself refuses to parse back.
 *
 * So a note is normalised as it arrives. That alone would leave the stored
 * value different from the bytes the server holds, and a value the server did
 * not give us is an edit it acts on - every multi-line note would be
 * rewritten on the next push. The shape the server used is therefore
 * remembered on the item and restored on the way out.
 *
 * Per item, not per account: it is evidence from that item's own payload
 * rather than a guess about the server, and an account can hold items from
 * more than one source.
 */

const asText = (v) => (v == null ? "" : String(v));

/** ical.js lower-cases property names on parse and looks them up
 *  case-sensitively, so every access uses the lower-case form; it still
 *  serialises the canonical upper-case name. */
export const CRLF_COMPAT = "x-eas-crlf-note";

/** A note as an iCalendar or vCard TEXT value can hold it. */
export const asNoteText = (v) => asText(v).replace(/\r\n?/g, "\n");

/** Record whether this item's server uses CRLF, from what it just sent. */
export function rememberNoteLineEndings(comp, raw) {
  if (asText(raw).includes("\r")) {
    comp.updatePropertyWithValue(CRLF_COMPAT, "1");
  } else {
    comp.removeAllProperties(CRLF_COMPAT);
  }
}

/** The note in the line endings its server uses.
 *
 *  The two steps are one operation: collapse whatever the value holds to bare
 *  newlines - including CRLFs the user just typed, which would otherwise be
 *  doubled - then expand every newline once. Idempotent, so a value that
 *  already carries CRLF survives unchanged, and a server that sends bare
 *  newlines never has CRLF invented for it. */
export function restoreNoteLineEndings(comp, value) {
  const text = asText(value);
  if (comp.getFirstPropertyValue(CRLF_COMPAT) !== "1") return text;
  return text.replace(/\r\n/g, "\n").replace(/\n/g, "\r\n");
}
