/**
 * What a create hook records when the calendar already holds the item.
 *
 * Thunderbird calls it a create whenever a write arrives carrying its own
 * id - an .ics import, another add-on, and above all iTIP re-creating a
 * meeting when the user answers an invitation to one already in the
 * calendar. Recording that as a new item sends an `<Add>` for something
 * the server has: mechnich's server answered Status 7 ("conflict, my copy
 * wins") and the edit never landed; Exchange, measured on a live account,
 * accepts it and keeps a second copy under the same UID.
 *
 * The stamp decides, not the presence of a prior copy - see the tests at
 * the bottom for why the difference matters.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

const { opForCreatedItem } = await import(
  "../../src/modules/calendar-provider.mjs"
);

const ics = (...extra) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    "BEGIN:VEVENT",
    "UID:create-op-uid",
    "DTSTAMP:20260801T090000Z",
    "DTSTART:20260810T140000Z",
    "DTEND:20260810T150000Z",
    "SUMMARY:create-op",
    ...extra,
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");

test("an item the server already knows is an update, not a create", () => {
  // The stamp is only ever written after the server has named the item, so
  // its presence is the one dependable statement that a push must address
  // the existing copy rather than make a new one.
  assert.equal(opForCreatedItem(ics("X-EAS-SERVERID:11:14")), "updated");
});

test("a local-only item stays a create", () => {
  // It exists here, but the server has never seen it. Calling this an
  // update would leave the push with no address - `dropUnsatisfiableEntry`
  // - and the edit would be discarded instead of sent.
  assert.equal(opForCreatedItem(ics()), "created");
});

test("no prior copy at all is a create", () => {
  assert.equal(opForCreatedItem(null), "created");
  assert.equal(opForCreatedItem(undefined), "created");
  assert.equal(opForCreatedItem(""), "created");
});

test("an unreadable prior is a create, never an update", () => {
  // A blob we cannot parse cannot testify that the server knows the item.
  // Erring towards "create" costs a duplicate at worst; erring the other
  // way discards the user's edit.
  assert.equal(opForCreatedItem("this is not iCalendar"), "created");
});

test("the stamp is read from the series, not from an override", () => {
  // A recurring item arrives with its overrides beside it. Only the master
  // carries the identity, and an override must not be able to speak for it
  // either way.
  const series = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    "BEGIN:VEVENT",
    "UID:create-op-uid",
    "DTSTART:20260810T140000Z",
    "RRULE:FREQ=DAILY;COUNT=3",
    "X-EAS-SERVERID:11:14",
    "END:VEVENT",
    "BEGIN:VEVENT",
    "UID:create-op-uid",
    "RECURRENCE-ID:20260811T140000Z",
    "DTSTART:20260811T160000Z",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
  assert.equal(opForCreatedItem(series), "updated");
});
