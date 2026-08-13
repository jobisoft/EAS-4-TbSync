/**
 * Unit tests for the recurrence elements neither codec models.
 *
 * `Regenerate`, `DeadOccur` and `CalendarType` have no representation in a
 * VTODO and no control in Thunderbird's task UI. They are still the user's
 * data, and a push rebuilds `<Recurrence>` from the RRULE - so anything not
 * re-emitted is dropped from the block we send.
 *
 * That matters because the server **replaces** the block rather than merging
 * it. Measured on both servers: send a weekly rule, change it to daily, and
 * the omitted `DayOfWeek` is gone from the server's own copy on 14.1 and on
 * 16.1 alike. Exchange Online states `Regenerate` and `DeadOccur` on every
 * task recurrence it sends, so they are demonstrably part of what a push
 * overwrites - and without this, renaming a regenerating task in Thunderbird
 * converts it to a fixed schedule.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test, before } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import {
  applicationDataToIcal,
  appendApplicationDataFromIcal,
} from "../../src/modules/eas/task-codec.mjs";
import {
  applicationDataToIcal as eventToIcal,
  appendApplicationDataFromIcal as appendEvent,
} from "../../src/modules/eas/calendar-codec.mjs";
import { createWBXML, decodeWBXML } from "../../src/modules/wbxml.mjs";
import { ensureLoaded } from "../../src/modules/eas/timezone-mapping.mjs";
import { parseAdNode } from "./support/ad-node.mjs";

before(() => ensureLoaded());

/** A task ApplicationData whose Recurrence carries `extra` verbatim. */
function taskAd(extra = "") {
  return `<ApplicationData>
  <Subject xmlns='Tasks'>regenerating</Subject>
  <UtcStartDate xmlns='Tasks'>2026-09-01T08:00:00.000Z</UtcStartDate>
  <StartDate xmlns='Tasks'>2026-09-01T08:00:00.000Z</StartDate>
  <UtcDueDate xmlns='Tasks'>2026-09-01T09:00:00.000Z</UtcDueDate>
  <DueDate xmlns='Tasks'>2026-09-01T09:00:00.000Z</DueDate>
  <Complete xmlns='Tasks'>0</Complete>
  <Recurrence xmlns='Tasks'>${extra}<Type xmlns='Tasks'>1</Type><Start xmlns='Tasks'>2026-09-01T08:00:00.000Z</Start><DayOfWeek xmlns='Tasks'>4</DayOfWeek><Interval xmlns='Tasks'>1</Interval><Occurrences xmlns='Tasks'>4</Occurrences></Recurrence>
</ApplicationData>`;
}

const inbound = (extra) =>
  applicationDataToIcal({
    adNode: parseAdNode(taskAd(extra)),
    serverID: "1:9",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: "regen@eas-test.invalid",
  });

/** Push `ical` and return the decoded `<Recurrence>`, namespaces stripped. */
function pushedRecurrence(ical) {
  const w = createWBXML("AirSync");
  w.otag("ApplicationData");
  appendApplicationDataFromIcal({
    builder: w,
    ical,
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
  });
  w.switchpage("AirSync");
  w.ctag();
  const found = /<Recurrence.*?<\/Recurrence>/s.exec(decodeWBXML(w.getBytes()));
  return found ? found[0].replace(/ xmlns='[^']*'/g, "") : null;
}

test("what the server states is parked on the item", async () => {
  const ical = await inbound(
    "<Regenerate xmlns='Tasks'>1</Regenerate><DeadOccur xmlns='Tasks'>0</DeadOccur>",
  );
  assert.match(ical, /X-EAS-REGENERATE:1/i);
  assert.match(ical, /X-EAS-DEADOCCUR:0/i);
  assert.match(ical, /RRULE:/, "the rule itself still maps as before");
});

test("and is handed back on the next push", async () => {
  // The case that used to destroy it: the item goes out again for a reason
  // that has nothing to do with its recurrence.
  const ical = await inbound(
    "<Regenerate xmlns='Tasks'>1</Regenerate><DeadOccur xmlns='Tasks'>1</DeadOccur>",
  );
  const sent = pushedRecurrence(ical.replace("regenerating", "renamed"));
  assert.match(sent, /<Regenerate>1<\/Regenerate>/);
  assert.match(sent, /<DeadOccur>1<\/DeadOccur>/);
  assert.match(sent, /<Type>1<\/Type>/, "the mapped elements are unaffected");
  assert.match(sent, /<DayOfWeek>4<\/DayOfWeek>/);
});

test("they precede Type, where 16.1 puts them", async () => {
  // The one detail with no second source: the Tasks codepage assigns their
  // tokens after the qualifying elements, while the block Exchange Online
  // sends puts them first. This follows the server, being the only sample
  // of a full server-authored Recurrence we have.
  const ical = await inbound("<Regenerate xmlns='Tasks'>1</Regenerate>");
  const sent = pushedRecurrence(ical);
  assert.ok(
    sent.indexOf("<Regenerate>") < sent.indexOf("<Type>"),
    `Regenerate must precede Type, got ${sent}`,
  );
});

test("a task the server says nothing about carries nothing", async () => {
  const ical = await inbound("");
  assert.doesNotMatch(ical, /X-EAS-REGENERATE/i);
  const sent = pushedRecurrence(ical);
  assert.doesNotMatch(sent, /Regenerate|DeadOccur|CalendarType/);
});

test("a stamp the server stops sending is dropped, not re-asserted", async () => {
  // Otherwise an item would keep telling the server something it has
  // already stopped saying - forever, on every push.
  const withFlag = await inbound("<Regenerate xmlns='Tasks'>1</Regenerate>");
  assert.match(withFlag, /X-EAS-REGENERATE:1/i);

  const without = await applicationDataToIcal({
    adNode: parseAdNode(taskAd("")),
    existingIcal: withFlag,
    serverID: "1:9",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: "regen@eas-test.invalid",
  });
  assert.doesNotMatch(without, /X-EAS-REGENERATE/i);
});

test("the non-Gregorian pair rides along on the same rails", async () => {
  const ical = await inbound(
    "<CalendarType xmlns='Tasks'>1</CalendarType><IsLeapMonth xmlns='Tasks'>0</IsLeapMonth>",
  );
  assert.match(ical, /X-EAS-CALENDARTYPE:1/i);
  assert.match(ical, /X-EAS-ISLEAPMONTH:0/i);
  const sent = pushedRecurrence(ical);
  assert.match(sent, /<CalendarType>1<\/CalendarType>/);
  assert.match(sent, /<IsLeapMonth>0<\/IsLeapMonth>/);
});

test("FirstDayOfWeek becomes WKST and goes back as the server sent it", async () => {
  // The one element of the set iCalendar can express. It is mapped inbound
  // so the rule expands correctly, and still carried verbatim outbound:
  // Thunderbird stamps `weekStart` from the `calendar.week.start`
  // preference whenever a rule is authored and offers no control for it, so
  // deriving the push from the local WKST would overwrite the server's
  // value with a profile setting the user never chose.
  const ical = await inbound(
    "<FirstDayOfWeek xmlns='Tasks'>0</FirstDayOfWeek>",
  );
  assert.match(ical, /WKST=SU/, "0 is Sunday, in both numbering schemes");
  assert.match(ical, /X-EAS-FIRSTDAYOFWEEK:0/i);

  const sent = pushedRecurrence(ical);
  assert.match(sent, /<FirstDayOfWeek>0<\/FirstDayOfWeek>/);
  assert.ok(
    sent.indexOf("<FirstDayOfWeek>") > sent.indexOf("<Type>"),
    `FirstDayOfWeek goes last, where both servers put it: ${sent}`,
  );
});

test("a Monday week start survives even though the rule cannot show it", async () => {
  // MO is RFC 5545's default, so ical.js parses `WKST=MO` and then leaves it
  // out of the serialised rule - the semantics are right, the text just does
  // not say so. The stamp is what carries it back to the server, which is
  // the second reason the push does not read the local WKST.
  const ical = await inbound(
    "<FirstDayOfWeek xmlns='Tasks'>1</FirstDayOfWeek>",
  );
  assert.doesNotMatch(ical, /WKST=/, "the default is not serialised");
  assert.match(ical, /X-EAS-FIRSTDAYOFWEEK:1/i);
  assert.match(pushedRecurrence(ical), /<FirstDayOfWeek>1<\/FirstDayOfWeek>/);
});

test("a rule authored here sends our own week start", async () => {
  // No stamp means the server never stated one, so the local rule is the
  // only statement of intent there is and it goes out. ical.js always has a
  // `wkst` - Monday when the rule does not say - so something is always
  // sent rather than leaving the server to guess.
  const ical = await inbound("");
  assert.doesNotMatch(ical, /X-EAS-FIRSTDAYOFWEEK/i, "nothing was stamped");
  assert.match(pushedRecurrence(ical), /<FirstDayOfWeek>1<\/FirstDayOfWeek>/);
});

test("but the server's own value wins over ours", async () => {
  // The asymmetry that matters: once the mailbox has said what its week
  // starts on, an edit made here must not replace it with a Thunderbird
  // preference the user was never shown.
  const ical = await inbound(
    "<FirstDayOfWeek xmlns='Tasks'>4</FirstDayOfWeek>",
  );
  assert.match(ical, /WKST=TH/);
  assert.match(pushedRecurrence(ical), /<FirstDayOfWeek>4<\/FirstDayOfWeek>/);
});

test("an event round-trips them through the same shared helper", async () => {
  // The calendar codec had the identical gap. In practice a 16.1 server
  // sends only FirstDayOfWeek on an ordinary event recurrence - measured -
  // and that one is deliberately not preserved, so this uses the
  // non-Gregorian pair, which is what would arrive for a rule we cannot
  // author from here.
  const ad = `<ApplicationData>
  <Subject xmlns='Calendar'>recurring</Subject>
  <StartTime xmlns='Calendar'>20260901T080000Z</StartTime>
  <EndTime xmlns='Calendar'>20260901T090000Z</EndTime>
  <Recurrence xmlns='Calendar'><CalendarType xmlns='Calendar'>1</CalendarType><IsLeapMonth xmlns='Calendar'>0</IsLeapMonth><Type xmlns='Calendar'>1</Type><DayOfWeek xmlns='Calendar'>4</DayOfWeek><Interval xmlns='Calendar'>1</Interval><Occurrences xmlns='Calendar'>3</Occurrences></Recurrence>
</ApplicationData>`;
  const ical = await eventToIcal({
    adNode: parseAdNode(ad),
    serverID: "e:1",
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
    uid: "recur@eas-test.invalid",
  });
  assert.match(ical, /X-EAS-CALENDARTYPE:1/i);
  assert.match(ical, /X-EAS-ISLEAPMONTH:0/i);

  const w = createWBXML("AirSync");
  w.otag("ApplicationData");
  appendEvent({
    builder: w,
    ical,
    asVersion: "16.1",
    defaultTimezone: "UTC",
    syncRecurrence: true,
  });
  w.switchpage("AirSync");
  w.ctag();
  const sent = /<Recurrence.*?<\/Recurrence>/s
    .exec(decodeWBXML(w.getBytes()))[0]
    .replace(/ xmlns='[^']*'/g, "");
  assert.match(sent, /<CalendarType>1<\/CalendarType>/);
  assert.match(sent, /<IsLeapMonth>0<\/IsLeapMonth>/);
  assert.ok(sent.indexOf("<CalendarType>") < sent.indexOf("<Type>"));
});
