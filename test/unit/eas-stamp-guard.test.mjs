/**
 * Unit tests for holding our own stamps against writers that are not us.
 *
 * The rule under test: a local item never invents a server identity, and
 * never loses one it already had.
 *
 * `X-EAS-SERVERID` and its siblings are an item's identity on the server
 * and our record of what the server said about it. Nothing outside this
 * add-on writes them, and nothing outside it can: our own sync writes go to
 * `<id>#cache`, which fires no item hooks. So every write reaching a hook is
 * somebody else's, and any difference in these properties is damage -
 * Thunderbird rebuilding an item from an emailed invitation, a copied event
 * carrying the identity of the one it came from, an import overwriting an
 * item that is already synced.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";

import { pinEasStamps } from "../../src/modules/eas/calendar-codec.mjs";

/** A VEVENT carrying whatever stamps a test needs. */
function event(stamps = [], { uid = "u-1", summary = "probe" } = {}) {
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//eas-test//EN",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    "DTSTAMP:20260801T120000Z",
    "DTSTART:20260901T100000Z",
    `SUMMARY:${summary}`,
    ...stamps,
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
}

const stampsIn = (ical) =>
  ical
    .split(/\r?\n/)
    .filter((l) => /^X-EAS-/i.test(l))
    .sort();

test("a new item cannot bring a server identity with it", () => {
  // Copy and paste, or an import of something that was never ours: the
  // item arrives wearing an identity that belongs to a different item.
  const out = pinEasStamps({ builtIcal: event(["X-EAS-SERVERID:stolen"]) });
  assert.deepEqual(stampsIn(out), []);
});

test("a genuinely new item is left alone", () => {
  const ical = event();
  assert.deepEqual(stampsIn(pinEasStamps({ builtIcal: ical })), []);
});

test("stamps the writer dropped come back", () => {
  // Accepting an invitation from the message makes Thunderbird rebuild the
  // item from the mail's iCalendar, which carries none of our properties.
  const out = pinEasStamps({
    builtIcal: event([], { summary: "rebuilt" }),
    priorIcal: event(["X-EAS-SERVERID:sid-1", "X-EAS-MEETINGSTATUS:3"]),
  });
  assert.deepEqual(stampsIn(out), [
    "X-EAS-MEETINGSTATUS:3",
    "X-EAS-SERVERID:sid-1",
  ]);
  assert.match(out, /SUMMARY:rebuilt/, "the writer's own change survives");
});

test("stamps the writer altered are reverted", () => {
  const out = pinEasStamps({
    builtIcal: event(["X-EAS-SERVERID:tampered"]),
    priorIcal: event(["X-EAS-SERVERID:sid-1"]),
  });
  assert.deepEqual(stampsIn(out), ["X-EAS-SERVERID:sid-1"]);
});

test("stamps the writer invented are dropped", () => {
  const out = pinEasStamps({
    builtIcal: event(["X-EAS-SERVERID:sid-1", "X-EAS-RESPONSETYPE:3"]),
    priorIcal: event(["X-EAS-SERVERID:sid-1"]),
  });
  assert.deepEqual(stampsIn(out), ["X-EAS-SERVERID:sid-1"]);
});

test("an unknown x-eas- property is covered too", () => {
  // Matched by prefix rather than by a list, so a stamp added later needs
  // nobody to remember this file.
  const out = pinEasStamps({
    builtIcal: event([]),
    priorIcal: event(["X-EAS-SOMETHING-NEW:7"]),
  });
  assert.deepEqual(stampsIn(out), ["X-EAS-SOMETHING-NEW:7"]);
});

test("an override never keeps a stamp", () => {
  // `stampEasServerId` writes to the master and the reader skips
  // overrides, so a stamp on one is somebody else's doing either way.
  const withOverride = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//eas-test//EN",
    "BEGIN:VEVENT",
    "UID:u-1",
    "DTSTAMP:20260801T120000Z",
    "DTSTART:20260901T100000Z",
    "RRULE:FREQ=DAILY;COUNT=3",
    "SUMMARY:master",
    "END:VEVENT",
    "BEGIN:VEVENT",
    "UID:u-1",
    "RECURRENCE-ID:20260902T100000Z",
    "DTSTAMP:20260801T120000Z",
    "DTSTART:20260902T110000Z",
    "SUMMARY:override",
    "X-EAS-SERVERID:on-the-override",
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
  const out = pinEasStamps({
    builtIcal: withOverride,
    priorIcal: event(["X-EAS-SERVERID:sid-1"]),
  });
  assert.deepEqual(stampsIn(out), ["X-EAS-SERVERID:sid-1"]);
  assert.match(out, /RECURRENCE-ID/, "the override itself survives");
});

test("nothing to do returns the input untouched", () => {
  // The common path: an ordinary edit to an item whose stamps nobody
  // disturbed. Same string back, so the caller can skip the write.
  const ical = event(["X-EAS-SERVERID:sid-1"]);
  assert.equal(pinEasStamps({ builtIcal: ical, priorIcal: ical }), ical);
});

test("a save is never failed over this", () => {
  // Unparseable in, unparseable out - refusing the user's edit because we
  // could not read it would be far worse than a missing stamp.
  assert.equal(pinEasStamps({ builtIcal: "not a calendar" }), "not a calendar");
  assert.equal(pinEasStamps({ builtIcal: "" }), "");
});
