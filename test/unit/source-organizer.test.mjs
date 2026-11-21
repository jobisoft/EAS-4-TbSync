/**
 * An item that names an organizer and invites nobody.
 *
 * A webinar or booking service hands out an "add to calendar" file that
 * names itself as ORGANIZER and lists no ATTENDEE. On 16.x the push held such
 * an item back as somebody else's meeting, found no answer to send, and
 * dropped the queued edit - the event never reached the server.
 *
 * It is rescued at the door, in the item hook, by taking the organizer out.
 * Only an item with an attendee on no component qualifies; everything else
 * is left exactly as it was, because those are invitations or meetings and
 * have their own handling.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import {
  isReceivedMeeting,
  withoutSourceOrganizer,
} from "../../src/modules/eas/calendar-codec.mjs";

const ME = "user@example.invalid";
const SERVICE = "noreply@booking.example.invalid";

const vevent = (...lines) => [
  "BEGIN:VEVENT",
  "UID:webinar@booking.example.invalid",
  "DTSTAMP:20261121T120000Z",
  "DTSTART:20261214T100000Z",
  "DTEND:20261214T103000Z",
  "SUMMARY:Webinar",
  ...lines,
  "END:VEVENT",
];

const cal = (...events) =>
  [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    ...events.flat(),
    "END:VCALENDAR",
    "",
  ].join("\r\n");

const organizerLine = `ORGANIZER;CN=Booking Service:mailto:${SERVICE}`;
const attendeeLine = `ATTENDEE;CN=User;PARTSTAT=ACCEPTED:mailto:${ME}`;
const count = (ical, name) =>
  ical.split("\r\n").filter((l) => l.startsWith(name)).length;

test("an organizer that invites nobody is taken out", () => {
  const out = withoutSourceOrganizer(cal(vevent(organizerLine)));
  assert.ok(out, "the item was rewritten");
  assert.equal(count(out, "ORGANIZER"), 0);
  assert.match(out, /SUMMARY:Webinar/, "nothing else was touched");
  assert.match(out, /UID:webinar@booking\.example\.invalid/);
});

test("every component of a series loses it, the overrides too", () => {
  const master = vevent(organizerLine, "RRULE:FREQ=WEEKLY;COUNT=3");
  const override = vevent(
    "RECURRENCE-ID:20261221T100000Z",
    "DTSTART:20261221T110000Z",
    organizerLine,
  );
  const out = withoutSourceOrganizer(cal(master, override));
  assert.equal(count(out, "ORGANIZER"), 0);
  assert.equal(count(out, "BEGIN:VEVENT"), 2, "the override is still there");
  assert.match(out, /RRULE:FREQ=WEEKLY;COUNT=3/);
});

test("an attendee on the master leaves the item alone", () => {
  assert.equal(
    withoutSourceOrganizer(cal(vevent(organizerLine, attendeeLine))),
    null,
  );
});

test("an attendee on one occurrence only is an invitation, and is left alone", () => {
  // Its answer is what the item is held back for, so it must stay a meeting.
  const master = vevent(organizerLine, "RRULE:FREQ=WEEKLY;COUNT=3");
  const override = vevent(
    "RECURRENCE-ID:20261221T100000Z",
    organizerLine,
    attendeeLine,
  );
  assert.equal(withoutSourceOrganizer(cal(master, override)), null);
});

test("a guest list without the user is still somebody else's meeting", () => {
  const guest = "ATTENDEE;CN=Guest:mailto:guest@example.invalid";
  assert.equal(withoutSourceOrganizer(cal(vevent(organizerLine, guest))), null);
});

test("an item with no organizer has nothing to take out", () => {
  assert.equal(withoutSourceOrganizer(cal(vevent())), null);
});

test("an item the server has described is not rewritten", () => {
  const described = cal(vevent(organizerLine, "X-EAS-MEETINGSTATUS:3"));
  assert.equal(withoutSourceOrganizer(described), null);
});

test("what cannot be read is not rewritten", () => {
  assert.equal(withoutSourceOrganizer("not a calendar"), null);
  assert.equal(withoutSourceOrganizer(""), null);
  assert.equal(withoutSourceOrganizer(undefined), null);
  assert.equal(withoutSourceOrganizer(cal()), null);
});

test("a second pass finds nothing left to do", () => {
  const once = withoutSourceOrganizer(cal(vevent(organizerLine)));
  assert.equal(withoutSourceOrganizer(once), null);
});

test("what comes out no longer reads as somebody else's meeting", () => {
  const before = cal(vevent(organizerLine));
  assert.equal(
    isReceivedMeeting(before, ME),
    true,
    "the problem: the foreign organizer made it a meeting",
  );
  assert.equal(isReceivedMeeting(withoutSourceOrganizer(before), ME), false);
});

test("both hooks drop the organizer after the stamps and before the shape", async () => {
  // The siblings `conformItem` splits off are made from what it is handed,
  // and the queue entry is filed from it, so the order is the point - read
  // out of the provider rather than restated here.
  const { readFileSync } = await import("node:fs");
  const src = readFileSync(
    new URL("../../src/modules/calendar-provider.mjs", import.meta.url),
    "utf8",
  );
  for (const hook of ["onItemCreated", "onItemUpdated"]) {
    const body = src.slice(src.indexOf(`provider.${hook}.addListener`));
    const guard = body.indexOf("guardStamps(");
    const drop = body.indexOf("dropSourceOrganizer(");
    const conform = body.indexOf("conformItem(");
    assert.ok(drop > 0, `${hook} no longer takes the organizer out`);
    assert.ok(guard > 0 && guard < conform, `${hook}: stamps before shape`);
    assert.ok(
      drop < conform,
      `${hook} conforms the item before the organizer is out of it`,
    );
  }
});
