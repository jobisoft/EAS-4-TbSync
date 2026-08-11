/**
 * Line endings for notes, shared by the calendar and contact codecs.
 *
 * Neither iCalendar nor vCard can hold a carriage return in a TEXT value, so
 * a note is normalised as it arrives - and the shape the server used is then
 * remembered and restored, or every multi-line note would go back in bytes
 * the server never sent, which it reads as an edit.
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import ICAL from "../../src/vendor/ical.min.js";
import {
  asNoteText,
  rememberNoteLineEndings,
  restoreNoteLineEndings,
  CRLF_COMPAT,
} from "../../src/modules/eas/note-text.mjs";

const comp = () => new ICAL.Component("vevent");

test("a carriage return never survives normalisation", () => {
  assert.equal(asNoteText("a\r\nb"), "a\nb", "CRLF collapses");
  assert.equal(asNoteText("a\rb"), "a\nb", "a lone CR becomes a newline");
  assert.equal(asNoteText("a\nb"), "a\nb", "bare newlines are left alone");
  assert.equal(asNoteText(null), "", "a missing note is empty, not null");
});

test("the server's shape is remembered from its own payload", () => {
  const crlf = comp();
  rememberNoteLineEndings(crlf, "one\r\ntwo");
  assert.equal(crlf.getFirstPropertyValue(CRLF_COMPAT), "1", "CRLF is marked");

  const lf = comp();
  rememberNoteLineEndings(lf, "one\ntwo");
  assert.equal(lf.getFirstPropertyValue(CRLF_COMPAT), null, "bare LF is not");

  // A server that stops sending CRLF clears the mark rather than leaving a
  // stale one behind.
  rememberNoteLineEndings(crlf, "one\ntwo");
  assert.equal(
    crlf.getFirstPropertyValue(CRLF_COMPAT),
    null,
    "the mark is dropped",
  );
});

test("restoring is idempotent, and invents nothing", () => {
  const crlf = comp();
  rememberNoteLineEndings(crlf, "one\r\ntwo");

  assert.equal(
    restoreNoteLineEndings(crlf, "one\ntwo"),
    "one\r\ntwo",
    "the server gets its CRLF back",
  );
  assert.equal(
    restoreNoteLineEndings(crlf, "one\r\ntwo"),
    "one\r\ntwo",
    "a value that already carries CRLF is unchanged - no doubling",
  );
  assert.equal(
    restoreNoteLineEndings(crlf, "one\r\ntwo\nthree"),
    "one\r\ntwo\r\nthree",
    "a line the user just typed expands exactly once",
  );

  const lf = comp();
  assert.equal(
    restoreNoteLineEndings(lf, "one\ntwo"),
    "one\ntwo",
    "no CRLF is invented for a server that never used it",
  );
});
