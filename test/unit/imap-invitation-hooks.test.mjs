import { test, mock, beforeEach } from "node:test";
import assert from "node:assert/strict";
import ICAL from "../../src/vendor/ical.min.js";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
let stored, records, prior;
const listeners = {};
const listener = name => ({ addListener(callback) { listeners[name] = callback; } });
browser.storage = { local: {
  async get(defaults) { return { ...structuredClone(defaults), ...structuredClone(stored) }; },
  async set(values) { Object.assign(stored, structuredClone(values)); },
  async remove(key) { delete stored[key]; },
} };
messenger.calendar.provider = Object.fromEntries([
  "onItemCreated", "onItemUpdated", "onItemRemoved", "onSync", "onResetSync",
].map(name => [name, listener(name)]));
messenger.calendar.calendars = {
  onUpdated: listener("calendarUpdated"), onRemoved: listener("calendarRemoved"),
};
messenger.calendar.items = { async get() { return prior ? { item: prior } : null; } };
mock.module(new URL("../../src/vendor/tbsync/change-queue.mjs", import.meta.url).href, {
  namedExports: {
    lookupBinding: async () => ({ sessionId: "session", accountId: "account" }),
    rememberBindings: async () => {},
    localQueue: () => ({ async record(entry) {
      records.push({ ...entry, source: structuredClone(stored["eas.clientInvitations.calendar"]) });
      return records.length;
    } }),
  },
});
mock.module(new URL("../../src/modules/sync-after-change.mjs", import.meta.url).href, {
  namedExports: { armSyncAfterChange() {}, quietSecondsFor: async () => 0 },
});
const { registerCalendarProvider } = await import("../../src/modules/calendar-provider.mjs");
registerCalendarProvider();
const { isClientInvitation, withClientInvitationLock } = await import("../../src/modules/eas/client-invitations.mjs");
const blob = (extra = [], sequence = 1) => [
  "BEGIN:VCALENDAR", "VERSION:2.0", "BEGIN:VEVENT", "UID:source@example.org",
  "DTSTAMP:20261021T100000Z", "DTSTART:20261022T100000Z", "DTEND:20261022T110000Z",
  "SUMMARY:meeting", "ORGANIZER:mailto:sender@example.org",
  "ATTENDEE;PARTSTAT=ACCEPTED:mailto:me@example.org", `SEQUENCE:${sequence}`,
  ...extra, "END:VEVENT", "END:VCALENDAR", "",
].join("\r\n");
const item = value => ({ id: "source@example.org", type: "event", format: "ical", item: value });
const create = (value, options = { invitation: true }) =>
  listeners.onItemCreated({ id: "calendar" }, item(value), options);
beforeEach(() => {
  stored = { "eas.imapInvitations.calendar": { user: "me@example.org" } };
  prior = null;
  records = [];
});

test("iTIP stores the original UID and revision before queueing its private appointment", async () => {
  const result = await create(blob());
  assert.equal(isClientInvitation(result.item), true);
  assert.match(result.item, /UID:source@example.org/);
  assert.match(result.item, /SEQUENCE:1/);
  assert.match(records[0].source["item:source@example.org"].blob, /SEQUENCE:1/);
});

test("ordinary ICS import, opt-out and unrelated attendee do not acquire client policy", async () => {
  for (const [value, options, enabled] of [
    [blob(), {}, true],
    [blob(), { invitation: true }, false],
    [blob().replace("me@example.org", "other@example.org"), { invitation: true }, true],
  ]) {
    stored["eas.imapInvitations.calendar"] = enabled ? { user: "me@example.org" } : null;
    const result = await create(value, options);
    assert.equal(isClientInvitation(result.item?.item || result.item), false);
    assert.equal(stored["eas.clientInvitations.calendar"], undefined);
  }
});

test("a server-native meeting stays on EAS scheduling", async () => {
  prior = blob(["X-EAS-SERVERID:server-native"]);
  const result = await create(blob());
  assert.equal(isClientInvitation(result.item), false);
  assert.match(result.item, /X-EAS-SERVERID:server-native/);
});

test("a later iTIP update retains the marker and saves the source's newer sequence", async () => {
  prior = (await create(blob())).item;
  const result = await listeners.onItemUpdated({ id: "calendar" }, item(blob([], 3)), item(prior), { invitation: true });
  assert.equal(isClientInvitation(result.item), true);
  assert.match(result.item, /SEQUENCE:3/);
  assert.match(records.at(-1).source["item:source@example.org"].blob, /SEQUENCE:3/);
});

test("an unsupported invitation series fails before any source or queue write", async () => {
  await assert.rejects(create(blob(["RRULE:FREQ=DAILY;COUNT=3", "RRULE:FREQ=WEEKLY;COUNT=3"])), /cykl/);
  assert.equal(records.length, 0);
  assert.equal(stored["eas.clientInvitations.calendar"], undefined);
});

test("iTIP with no object id retains its source UID", async () => {
  const result = await listeners.onItemCreated({ id: "calendar" }, { ...item(blob()), id: null }, { invitation: true });
  assert.equal(result.id, "source@example.org");
  assert.match(result.item, /UID:source@example.org/);
});

test("a lone occurrence is merged with its full source and keeps the other overrides", async () => {
  const original = new ICAL.Component(ICAL.parse(blob([
    "X-EAS-CLIENT-INVITATION:1", "RRULE:FREQ=DAILY;COUNT=3",
  ])));
  const oldException = new ICAL.Component(ICAL.parse(blob([
    "RECURRENCE-ID:20261023T100000Z", "LOCATION:old exception",
  ]))).getFirstSubcomponent("vevent");
  original.addSubcomponent(oldException);
  prior = original.toString();
  const oldOccurrence = blob(["RECURRENCE-ID:20261024T100000Z"]);
  const newOccurrence = blob(["RECURRENCE-ID:20261024T100000Z", "LOCATION:new exception"], 4);
  const result = await listeners.onItemUpdated({ id: "calendar" }, item(newOccurrence), item(oldOccurrence), { invitation: true });
  const events = new ICAL.Component(ICAL.parse(result.item)).getAllSubcomponents("vevent");
  assert.equal(events.length, 3);
  assert.ok(events.every(event => event.getFirstPropertyValue("uid") === "source@example.org"));
  assert.match(result.item, /LOCATION:old exception/);
  assert.match(result.item, /LOCATION:new exception/);
  assert.match(records[0].source["item:source@example.org"].blob, /LOCATION:new exception/);
  assert.match(records[0].source["item:source@example.org"].blob, /SEQUENCE:4/);
});

test("a client occurrence without its full source or with a different UID is refused", async () => {
  const occurrence = blob(["RECURRENCE-ID:20261024T100000Z"]);
  await assert.rejects(listeners.onItemUpdated({ id: "calendar" }, item(occurrence), item(occurrence), { invitation: true }), /Brak pełnego zaproszenia/);
  prior = blob(["X-EAS-CLIENT-INVITATION:1", "RRULE:FREQ=DAILY;COUNT=3"]).replace("UID:source@example.org", "UID:other@example.org");
  await assert.rejects(listeners.onItemUpdated({ id: "calendar" }, item(occurrence), item(occurrence), { invitation: true }), /Nie można połączyć/);
  assert.equal(records.length, 0);
  assert.equal(stored["eas.clientInvitations.calendar"], undefined);
});

test("deletion waits for the same invitation ACK lock before recording its pending operation", async () => {
  let releaseAck;
  let enteredAck;
  const ackEntered = new Promise(resolve => { enteredAck = resolve; });
  const ack = withClientInvitationLock("calendar", "source@example.org", async () => {
    enteredAck();
    await new Promise(resolve => { releaseAck = resolve; });
    records.push({ op: "ack-complete" });
  });
  await ackEntered;
  const removal = listeners.onItemRemoved({ id: "calendar" }, item(blob()), {});
  await Promise.resolve();
  assert.equal(records.length, 0, "the delete cannot be queued while an older ACK can clear it");
  releaseAck();
  await Promise.all([ack, removal]);
  assert.deepEqual(records.map(entry => entry.op), ["ack-complete", "deleted"]);
});

test("an accepted mail update for a single event preserves its existing Exchange identity", async () => {
  prior = blob(["X-EAS-CLIENT-INVITATION:1", "X-EAS-SERVERID:existing-ms-event"]);
  const update = blob([], 5).replace("SUMMARY:meeting", "SUMMARY:updated meeting")
    .replace("DTSTART:20261022T100000Z", "DTSTART:20261022T120000Z");
  const result = await create(update);
  assert.match(result.item, /UID:source@example.org/);
  assert.match(result.item, /X-EAS-SERVERID:existing-ms-event/);
  assert.match(result.item, /SEQUENCE:5/);
  assert.match(result.item, /DTSTART:20261022T120000Z/);
  assert.equal(records[0].op, "updated", "same UID plus prior server identity must queue Change, not Add");
  assert.equal(records[0].source["item:source@example.org"].serverID, "existing-ms-event");
});
