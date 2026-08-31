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

import {
  easStampsAgree,
  pinEasStamps,
} from "../../src/modules/eas/calendar-codec.mjs";

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

test("a task keeps its stamps too", () => {
  // One calendar type serves events and tasks, so a VTODO reaches this the
  // same way a VEVENT does. The strip half always covered both; the restore
  // half looked for a VEVENT master and found none, so editing a task
  // deleted its stamps outright - X-EAS-SERVERID included.
  const todo = (stamps = []) =>
    [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "PRODID:-//eas-test//EN",
      "BEGIN:VTODO",
      "UID:t-1",
      "DTSTAMP:20260801T120000Z",
      "DTSTART:20260901T080000Z",
      "SUMMARY:probe",
      ...stamps,
      "END:VTODO",
      "END:VCALENDAR",
    ].join("\r\n");

  const out = pinEasStamps({
    builtIcal: todo([]),
    priorIcal: todo(["X-EAS-SERVERID:sid-1", "X-EAS-DEADOCCUR:1"]),
  });
  assert.deepEqual(stampsIn(out), [
    "X-EAS-DEADOCCUR:1",
    "X-EAS-SERVERID:sid-1",
  ]);
});

/* ------------------------------------------------------------------ *
 * What counts as somebody writing to a stamp.
 *
 * The guard has one question: does the incoming item carry the stamps we
 * stored for it. `pinEasStamps` cannot answer it - it discards the incoming
 * stamps unread and writes the stored ones over the top, so what it returns
 * is a function of the stored copy alone. Reading the answer out of it
 * instead compares the two documents, and a document differs for reasons
 * that are not stamps: how it is serialised, where in a component a stamp
 * sits, the order two of them are held in, which components exist.
 *
 * Each of those is a writer reported for something they did not do, so each
 * is pinned here.
 * ------------------------------------------------------------------ */

const SERVERID = "X-EAS-SERVERID:U2f1ad:57a543ff";
const MEETING = "X-EAS-MEETINGSTATUS:0";

/** What the guard does: repair, then ask whether a stamp moved. */
function agree(incoming, stored) {
  return easStampsAgree(
    incoming,
    pinEasStamps({ builtIcal: incoming, priorIcal: stored }),
  );
}

/** A VEVENT whose properties come out in the order given, so a test can put
 *  a stamp somewhere other than last. */
function ordered(lines, { uid = "u-1" } = {}) {
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    "DTSTART:20260901T100000Z",
    ...lines,
    "END:VEVENT",
    "END:VCALENDAR",
  ].join("\r\n");
}

/** A series, optionally with an occurrence override, and stamps placed on
 *  whichever of the two the test names. */
function series({ masterStamps = [], overrideStamps = null } = {}) {
  const master = [
    "BEGIN:VEVENT",
    "UID:u-r",
    "DTSTART:20260901T100000Z",
    "RRULE:FREQ=DAILY;COUNT=3",
    ...masterStamps,
    "END:VEVENT",
  ];
  const override =
    overrideStamps === null
      ? []
      : [
          "BEGIN:VEVENT",
          "UID:u-r",
          "RECURRENCE-ID:20260902T100000Z",
          "DTSTART:20260902T110000Z",
          ...overrideStamps,
          "END:VEVENT",
        ];
  return ["BEGIN:VCALENDAR", "VERSION:2.0", ...master, ...override, "END:VCALENDAR"].join(
    "\r\n",
  );
}

test("serialisation is not a stamp", () => {
  const stored = event([SERVERID]);
  for (const [name, incoming] of [
    ["LF line endings", stored.replace(/\r\n/g, "\n")],
    [
      "a long line left unfolded",
      stored.replace(
        "SUMMARY:probe",
        "ORGANIZER;RSVP=FALSE;CN=John Bieling;PARTSTAT=ACCEPTED;ROLE=CHAIR:mailto:john.bieling@ekir.de",
      ),
    ],
  ]) {
    assert.ok(agree(incoming, stored), name);
  }
});

test("where a stamp sits in its component is not a stamp", () => {
  const stored = ordered(["SUMMARY:probe", SERVERID]);
  const incoming = ordered([SERVERID, "SUMMARY:probe"]);
  assert.ok(agree(incoming, stored));
});

test("the order two stamps are held in is not a stamp", () => {
  const stored = ordered(["SUMMARY:probe", MEETING, SERVERID]);
  const incoming = ordered(["SUMMARY:probe", SERVERID, MEETING]);
  assert.ok(agree(incoming, stored));
});

test("a component appearing or disappearing is not a stamp", () => {
  const withoutOverride = series({ masterStamps: [SERVERID] });
  const withOverride = series({ masterStamps: [SERVERID], overrideStamps: [] });
  assert.ok(agree(withOverride, withoutOverride), "an override was added");
  assert.ok(agree(withoutOverride, withOverride), "an override was removed");
});

test("a stamp altered, removed or invented is a stamp", () => {
  const stored = event([SERVERID]);
  const cases = {
    altered: event(["X-EAS-SERVERID:U2f1ad:somebody-elses"]),
    removed: event([]),
    invented: event([SERVERID, "X-EAS-RESPONSETYPE:3"]),
  };
  for (const [name, incoming] of Object.entries(cases)) {
    assert.equal(agree(incoming, stored), false, name);
  }
});

test("a stamp moving between a series and an occurrence is a stamp", () => {
  const onMaster = series({ masterStamps: [SERVERID], overrideStamps: [] });
  const onOverride = series({ masterStamps: [], overrideStamps: [SERVERID] });
  assert.equal(agree(onOverride, onMaster), false);
});

test("with nothing stored, any stamp is one it has no claim to", () => {
  assert.ok(agree(event([]), null));
  assert.equal(agree(event([SERVERID]), null), false);
});

test("a document with nothing to guard is left alone", () => {
  assert.ok(agree("not a calendar", event([SERVERID])));
  assert.ok(
    agree(
      ["BEGIN:VCALENDAR", "VERSION:2.0", "END:VCALENDAR"].join("\r\n"),
      event([SERVERID]),
    ),
  );
});
