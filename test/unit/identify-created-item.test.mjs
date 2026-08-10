/**
 * An item created in Thunderbird's own dialog reaches the provider hook
 * with NO id: core decides an edit is an *addition* precisely by the
 * absence of one, and the platform mints its id only after the hook has
 * answered. Our queue files against an id, so such an item used to be
 * dropped without a trace and never synced.
 *
 * `identify` gives it one and hands back props the platform rebuilds the
 * item from, so its own fallback id never happens and the id we queued
 * against is the id the item ends up carrying. Live proof is suite 2.8,
 * which creates an item with no UID and follows it to the server; this
 * pins the arithmetic underneath.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

import ICAL from "../../src/vendor/ical.min.js";
import { identify } from "../../src/modules/calendar-provider.mjs";

const UID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;

/** What `convertItem(item, {returnFormat: "ical"})` hands the hook. */
const props = (icalLines, overrides = {}) => ({
  type: "event",
  id: null,
  format: "ical",
  item: [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//test//EN",
    "BEGIN:VEVENT",
    "DTSTAMP:20260801T120000Z",
    ...icalLines,
    "END:VEVENT",
    "END:VCALENDAR",
    "",
  ].join("\r\n"),
  ...overrides,
});

const uidIn = (ics, comp = "vevent") =>
  new ICAL.Component(ICAL.parse(ics))
    .getFirstSubcomponent(comp)
    ?.getFirstPropertyValue("uid");

test("an id-less event gets a UID, in the props and in the iCal alike", () => {
  const out = identify(props(["SUMMARY:Zahnarzt", "DTSTART:20261111T123000Z"]));
  assert.ok(out, "expected props to hand back");
  assert.match(out.id, UID_RE);
  assert.equal(
    uidIn(out.item),
    out.id,
    "the iCal must carry the same UID - the platform rebuilds the item " +
      "from it, and a mismatch would queue against an id no item has",
  );
  assert.equal(
    out.type,
    "event",
    "the type must survive: it is what makes the platform rebuild",
  );
  assert.match(out.item, /SUMMARY:Zahnarzt/, "the user's content is untouched");
});

test("a task is identified the same way", () => {
  const out = identify({
    ...props(["SUMMARY:call the dentist"]),
    type: "task",
    item: [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "BEGIN:VTODO",
      "DTSTAMP:20260801T120000Z",
      "SUMMARY:call the dentist",
      "END:VTODO",
      "END:VCALENDAR",
      "",
    ].join("\r\n"),
  });
  assert.ok(out, "a VTODO must be identified too");
  assert.equal(uidIn(out.item, "vtodo"), out.id);
});

test("an item that already has an id is left entirely alone", () => {
  const withId = props(["UID:keep-me@example.org", "SUMMARY:x"], {
    id: "keep-me@example.org",
  });
  assert.equal(
    identify(withId),
    null,
    "returning null means the hook passes the item through untouched",
  );
});

test("two creates never collide", () => {
  const a = identify(props(["SUMMARY:a"]));
  const b = identify(props(["SUMMARY:b"]));
  assert.notEqual(a.id, b.id);
});

test("an unparseable or empty blob falls back rather than throwing", () => {
  // The user's save is being held while this runs. Whatever is wrong with
  // the blob, losing their event over it would be worse.
  assert.equal(
    identify({ type: "event", id: null, item: "not iCalendar" }),
    null,
  );
  assert.equal(identify({ type: "event", id: null, item: "" }), null);
  assert.equal(
    identify({
      type: "event",
      id: null,
      item: "BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n",
    }),
    null,
    "a VCALENDAR with no event or task component",
  );
  assert.equal(identify(null), null);
  assert.equal(identify(undefined), null);
});

test("a recurring create gets ONE uid across master and every override", () => {
  // Dropping an .ics with modified occurrences, or dragging such an
  // event onto the calendar, clears the id on the whole series at once -
  // master and overrides arrive id-less together. Stamping only the
  // first would orphan the overrides: RFC 5545 binds an override to its
  // master by UID, and the platform refuses to rebuild an exception
  // whose id does not match its parent's, so the user's save fails
  // outright. This is the case the first cut of `identify` got wrong.
  const out = identify({
    type: "event",
    id: null,
    format: "ical",
    item: [
      "BEGIN:VCALENDAR",
      "VERSION:2.0",
      "BEGIN:VTIMEZONE",
      "TZID:Europe/Berlin",
      "BEGIN:STANDARD",
      "DTSTART:19701025T030000",
      "TZOFFSETFROM:+0200",
      "TZOFFSETTO:+0100",
      "END:STANDARD",
      "END:VTIMEZONE",
      "BEGIN:VEVENT",
      "DTSTAMP:20260801T120000Z",
      "SUMMARY:weekly",
      "DTSTART;TZID=Europe/Berlin:20261201T090000",
      "RRULE:FREQ=WEEKLY;COUNT=5",
      "END:VEVENT",
      "BEGIN:VEVENT",
      "DTSTAMP:20260801T120000Z",
      "SUMMARY:weekly (moved)",
      "RECURRENCE-ID;TZID=Europe/Berlin:20261215T090000",
      "DTSTART;TZID=Europe/Berlin:20261215T110000",
      "END:VEVENT",
      "END:VCALENDAR",
      "",
    ].join("\r\n"),
  });
  assert.ok(out, "expected props to hand back");
  const events = new ICAL.Component(ICAL.parse(out.item)).getAllSubcomponents(
    "vevent",
  );
  assert.equal(events.length, 2, "both components must survive");
  assert.deepEqual(
    events.map((v) => v.getFirstPropertyValue("uid")),
    [out.id, out.id],
    "master and override must carry the same uid",
  );
  assert.ok(
    new ICAL.Component(ICAL.parse(out.item)).getFirstSubcomponent("vtimezone"),
    "the VTIMEZONE must survive the round trip",
  );
  assert.equal(
    out.format,
    "ical",
    "the format must survive - propsToItem needs it",
  );
});
