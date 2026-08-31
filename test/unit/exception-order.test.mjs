/**
 * The order modified occurrences are written in.
 *
 * A server applies each exception against the occurrences as they stand at
 * that moment and refuses one that would carry an occurrence past a
 * sibling that has not moved yet - silently, with Status 1, so the
 * exception is just absent on the next pull. Measured on Z-Push both ways:
 * two occurrences moved later survive only when the latest is written
 * first, two moved earlier only when the earliest is.
 *
 * Asserted on the decoded wire, never on the blob: the blob order is the
 * input to the thing under test.
 */

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import {
  appendApplicationDataFromIcal,
  listInstanceCommands,
} from "../../src/modules/eas/calendar-codec.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";

before(() => ensureLoaded());

/** A series with `moves` applied as overrides, listed in the given blob
 *  order so the test can prove the writer does not simply keep it. */
function series({ rule, moves }) {
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "BEGIN:VEVENT",
    "UID:order@eas-test.invalid",
    "DTSTAMP:20260801T120000Z",
    "DTSTART:20261106T090000Z",
    "DTEND:20261106T100000Z",
    "SUMMARY:order probe",
    `RRULE:${rule}`,
    "END:VEVENT",
  ];
  for (const [slot, to] of moves) {
    lines.push(
      "BEGIN:VEVENT",
      "UID:order@eas-test.invalid",
      "DTSTAMP:20260801T120000Z",
      `RECURRENCE-ID:${slot}T090000Z`,
      `DTSTART:${to}T090000Z`,
      `DTEND:${to}T100000Z`,
      `SUMMARY:moved to ${to}`,
      "END:VEVENT",
    );
  }
  return lines.concat("END:VCALENDAR").join("\r\n");
}

/** The ExceptionStartTimes in the order they reach the wire, <=14.x. */
function embeddedOrder(ical) {
  const w = createWBXML("AirSync");
  w.otag("ApplicationData");
  appendApplicationDataFromIcal({
    builder: w,
    ical,
    asVersion: "14.1",
    defaultTimezone: "UTC",
    suppressExceptions: false,
    userEmail: "user@example.invalid",
    fallbackOrganizerName: null,
    eventLog: null,
  });
  w.switchpage("AirSync");
  w.ctag();
  return [
    ...decodeWBXML(w.getBytes()).matchAll(/<ExceptionStartTime[^>]*>([^<]+)</g),
  ].map((m) => m[1]);
}

/** The InstanceIds in the order they reach the wire, 16.1. */
const instanceOrder = (ical) =>
  listInstanceCommands({
    blob: ical,
    serverID: "srv-1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    userEmail: "user@example.invalid",
    fallbackOrganizerName: null,
    eventLog: null,
  }).map((c) => c.instanceId);

// Rule: 6, 7, 8 Nov. Both moves are LATER, and both are listed in the
// blob earliest-first - the order that loses one of them.
const MOVED_LATER = series({
  rule: "FREQ=DAILY;COUNT=3",
  moves: [
    ["20261107", "20261119"],
    ["20261108", "20261225"],
  ],
});

// Rule: 6 Nov, 6 Dec, 6 Jan. Both moves are EARLIER, listed latest-first -
// again the order that loses one.
const MOVED_EARLIER = series({
  rule: "FREQ=MONTHLY;COUNT=3",
  moves: [
    ["20270106", "20261119"],
    ["20261206", "20261112"],
  ],
});

test("occurrences moved later are written latest-first", () => {
  // Written earliest-first, moving 7 Nov to 19 Nov carries it past the
  // 8 Nov occurrence, which has not moved yet, and the server drops it.
  assert.deepEqual(embeddedOrder(MOVED_LATER), [
    "20261108T090000Z",
    "20261107T090000Z",
  ]);
});

test("occurrences moved earlier are written earliest-first", () => {
  // The mirror: moving 6 Jan back to 19 Nov carries it past 6 Dec.
  assert.deepEqual(embeddedOrder(MOVED_EARLIER), [
    "20261206T090000Z",
    "20270106T090000Z",
  ]);
});

test("16.x instance commands take the same order", () => {
  // Same rule, same reason - only the spelling of the identifier differs.
  assert.deepEqual(instanceOrder(MOVED_LATER), [
    "20261108T090000Z",
    "20261107T090000Z",
  ]);
  assert.deepEqual(instanceOrder(MOVED_EARLIER), [
    "20261206T090000Z",
    "20270106T090000Z",
  ]);
});

test("an occurrence that has not moved is written too", () => {
  const edited = series({
    rule: "FREQ=DAILY;COUNT=3",
    moves: [["20261107", "20261107"]],
  });
  assert.deepEqual(embeddedOrder(edited), ["20261107T090000Z"]);
});
