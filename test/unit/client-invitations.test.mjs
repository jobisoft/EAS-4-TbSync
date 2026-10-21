import { test } from "node:test";
import assert from "node:assert/strict";
import ICAL from "../../src/vendor/ical.min.js";
import { installWebextEnv } from "./support/webext-env.mjs";
import {
  isClientInvitation, prepareClientInvitation, toClientAppointment,
  mergeClientInvitationOccurrence,
  withClientInvitationLock,
  restoreClientInvitation, rememberClientInvitation, clientInvitationFor,
  clientInvitationByServerId, forgetClientInvitation,
} from "../../src/modules/eas/client-invitations.mjs";

installWebextEnv();
const USER = "person@example.invalid";
const ORGANIZER = "mailto:organizer@example.invalid";
function event(extra = [], uid = "original-uid") {
  return ["BEGIN:VEVENT", `UID:${uid}`, "DTSTART:20261101T100000Z",
    "DTEND:20261101T110000Z", "SUMMARY:Original invitation", "SEQUENCE:4",
    "DTSTAMP:20261021T080000Z", "X-MOZ-RECEIVED-SEQUENCE:4",
    "X-MOZ-RECEIVED-DTSTAMP:20261021T080000Z", `ORGANIZER;CN=External:${ORGANIZER}`,
    `ATTENDEE;PARTSTAT=ACCEPTED:mailto:${USER}`, "ATTENDEE;CN=Guest:mailto:guest@example.invalid",
    `X-MOZ-INVITED-ATTENDEE:mailto:${USER}`, ...extra, "END:VEVENT"];
}
function calendar(...events) {
  return ["BEGIN:VCALENDAR", "VERSION:2.0", ...events.flat(), "END:VCALENDAR", ""].join("\r\n");
}
const SOURCE = calendar(event());
const marked = (blob = SOURCE) => prepareClientInvitation({ blob, enabled: true, userEmail: USER });
const components = (blob) => new ICAL.Component(ICAL.parse(blob)).getAllSubcomponents("vevent");

test("only a new received invitation is marked; prior identity survives edits", () => {
  assert.equal(isClientInvitation(SOURCE), false);
  const prepared = marked();
  assert.equal(isClientInvitation(prepared), true);
  assert.equal(typeof prepared, "string");
  assert.equal(prepareClientInvitation({ blob: SOURCE, enabled: false, userEmail: USER }), SOURCE);
  const own = SOURCE.replace(ORGANIZER, `mailto:${USER}`);
  assert.equal(marked(own), own);
  const tracked = calendar(event(["X-EAS-SERVERID:server-one"]));
  assert.equal(marked(tracked), tracked);
  assert.equal(isClientInvitation(prepareClientInvitation({ blob: tracked, priorBlob: prepared, enabled: false })), true);
  const unrelated = tracked.replace("UID:original-uid", "UID:other");
  assert.equal(isClientInvitation(prepareClientInvitation({ blob: unrelated, priorBlob: prepared, enabled: false })), false);
  assert.equal(SOURCE.includes("X-EAS-CLIENT-INVITATION"), false);
});

test("marker is on the master; an override alone cannot opt in", () => {
  const blob = calendar(event(), event(["RECURRENCE-ID:20261108T100000Z", "X-EAS-CLIENT-INVITATION:1"]));
  assert.equal(isClientInvitation(blob), false);
  assert.equal(toClientAppointment(blob), blob);
  assert.equal(isClientInvitation("broken"), false);
  assert.equal(prepareClientInvitation({ blob: "broken", enabled: true }), "broken");
});

test("outbound appointment projection strips participants on every VEVENT only", () => {
  const source = marked(calendar(event(), event(["RECURRENCE-ID:20261108T100000Z"])));
  const original = source;
  const projected = toClientAppointment(source);
  for (const ev of components(projected)) {
    assert.equal(ev.hasProperty("organizer"), false);
    assert.equal(ev.hasProperty("attendee"), false);
    assert.equal(ev.hasProperty("x-moz-invited-attendee"), false);
    assert.equal(ev.getFirstPropertyValue("uid"), "original-uid");
    assert.equal(ev.getFirstPropertyValue("summary"), "Original invitation");
  }
  assert.equal(source, original);
  assert.equal(components(source)[0].getAllProperties("attendee").length, 2);
  assert.equal(toClientAppointment(SOURCE), SOURCE);
});

test("restore retains iTIP identity and revisions while server content and stamps win", () => {
  const source = marked();
  const built = calendar(["BEGIN:VEVENT", "UID:server-uid", "SUMMARY:Changed in Teams",
    "DTSTART:20261101T130000Z", "DTEND:20261101T140000Z", "SEQUENCE:99",
    "DTSTAMP:20261022T120000Z", "X-EAS-SERVERID:server-one", "X-EAS-MEETINGSTATUS:0",
    "ORGANIZER:mailto:server-owner@example.invalid", "END:VEVENT"]);
  const restored = restoreClientInvitation({ builtBlob: built, sourceBlob: source });
  const ev = components(restored)[0];
  const original = components(source)[0];
  for (const name of ["uid", "organizer", "attendee", "sequence", "dtstamp", "x-moz-received-sequence", "x-moz-received-dtstamp", "x-moz-invited-attendee", "x-eas-client-invitation"]) {
    assert.deepEqual(ev.getAllProperties(name).map((p) => p.toJSON()), original.getAllProperties(name).map((p) => p.toJSON()));
  }
  assert.equal(ev.getFirstPropertyValue("summary"), "Changed in Teams");
  assert.equal(ev.getFirstPropertyValue("dtstart").toString(), "2026-11-01T13:00:00Z");
  assert.equal(ev.getFirstPropertyValue("x-eas-serverid"), "server-one");
  assert.equal(ev.getFirstPropertyValue("x-eas-meetingstatus"), "0");
  assert.equal(components(built)[0].getFirstPropertyValue("uid"), "server-uid");
  assert.equal(restoreClientInvitation({ builtBlob: built, sourceBlob: SOURCE }), built);
});

test("recurrence restoration matches occurrence identity and never resurrects server deletions", () => {
  const source = marked(calendar(event(), event(["RECURRENCE-ID:20261108T100000Z"]), event(["RECURRENCE-ID:20261115T100000Z"])));
  const built = calendar(["BEGIN:VEVENT", "UID:server-uid", "SUMMARY:Server series", "EXDATE:20261115T100000Z", "END:VEVENT"],
    ["BEGIN:VEVENT", "UID:server-uid", "RECURRENCE-ID:20261108T100000Z", "SUMMARY:Server override", "DTSTART:20261108T120000Z", "END:VEVENT"]);
  const restored = components(restoreClientInvitation({ builtBlob: built, sourceBlob: source }));
  assert.equal(restored.length, 2);
  assert.equal(restored[0].hasProperty("exdate"), true);
  assert.equal(restored[1].getFirstPropertyValue("uid"), "original-uid");
  assert.equal(restored[1].getFirstPropertyValue("organizer"), ORGANIZER);
  assert.equal(restored[1].getFirstPropertyValue("summary"), "Server override");
  assert.equal(restored[1].getFirstPropertyValue("dtstart").toString(), "2026-11-08T12:00:00Z");
  assert.equal(components(source).length, 3);
});

test("an override added on the server inherits invitation identity and revision from the source master", () => {
  const built = calendar(["BEGIN:VEVENT", "UID:server-uid", "END:VEVENT"],
    ["BEGIN:VEVENT", "UID:server-uid", "RECURRENCE-ID:20261122T100000Z", "SUMMARY:New server exception",
      "ORGANIZER:mailto:server-owner@example.invalid", "SEQUENCE:99", "DTSTART:20261122T120000Z", "END:VEVENT"]);
  const restored = components(restoreClientInvitation({ builtBlob: built, sourceBlob: marked() }));
  assert.equal(restored.length, 2);
  assert.equal(restored[1].getFirstPropertyValue("uid"), "original-uid");
  assert.equal(restored[1].getFirstPropertyValue("summary"), "New server exception");
  assert.equal(restored[1].getFirstPropertyValue("organizer"), ORGANIZER);
  const self = restored[1].getAllProperties("attendee").find((p) => p.getFirstValue() === `mailto:${USER}`);
  assert.equal(self.getParameter("partstat"), "ACCEPTED");
  assert.equal(restored[1].getAllProperties("attendee").length, 2);
  assert.equal(restored[1].getFirstPropertyValue("sequence"), 4);
  assert.equal(restored[1].getFirstPropertyValue("x-moz-received-sequence"), "4");
  assert.equal(restored[1].getFirstPropertyValue("dtstart").toString(), "2026-11-22T12:00:00Z");
});

test("a lone occurrence update keeps its marked parent and unrelated overrides", () => {
  const parent = marked(calendar(event(["RRULE:FREQ=WEEKLY;COUNT=5"]),
    event(["RECURRENCE-ID:20261108T100000Z"]), event(["RECURRENCE-ID:20261115T100000Z"])));
  const incoming = calendar(event(["RECURRENCE-ID:20261108T100000Z"])).replace("SUMMARY:Original invitation", "SUMMARY:Changed occurrence");
  const merged = mergeClientInvitationOccurrence({ blob: incoming, masterBlob: parent });
  const events = components(merged);
  assert.equal(events.length, 3);
  assert.equal(isClientInvitation(merged), true);
  assert.deepEqual(events.find((ev) => !ev.hasProperty("recurrence-id")).toJSON(), components(parent)[0].toJSON());
  assert.equal(events.find((ev) => ev.getFirstPropertyValue("recurrence-id")?.toString() === "2026-11-08T10:00:00Z").getFirstPropertyValue("summary"), "Changed occurrence");
  assert.deepEqual(events.find((ev) => ev.getFirstPropertyValue("recurrence-id")?.toString() === "2026-11-15T10:00:00Z").toJSON(), components(parent)[2].toJSON());
  assert.equal(components(incoming).length, 1);
  assert.equal(components(parent)[1].getFirstPropertyValue("summary"), "Original invitation");
});

test("a new lone occurrence is added without replacing the parent", () => {
  const parent = marked(calendar(event(["RRULE:FREQ=WEEKLY;COUNT=5"])));
  const incoming = calendar(event(["RECURRENCE-ID:20261122T100000Z"]));
  const merged = mergeClientInvitationOccurrence({ blob: incoming, masterBlob: parent });
  assert.equal(components(merged).length, 2);
  assert.equal(isClientInvitation(merged), true);
  assert.ok(components(merged).every((ev) => ev.getFirstPropertyValue("uid") === "original-uid"));
  assert.equal(components(merged)[0].getFirstPropertyValue("rrule").toString(), "FREQ=WEEKLY;COUNT=5");
});

test("occurrence merge declines unrelated, unavailable or unmarked parent identity", () => {
  const incoming = calendar(event(["RECURRENCE-ID:20261122T100000Z"]));
  for (const masterBlob of [undefined, "broken", SOURCE, marked(calendar(event([], "different-uid")))]) {
    assert.equal(mergeClientInvitationOccurrence({ blob: incoming, masterBlob }), incoming);
  }
  assert.equal(mergeClientInvitationOccurrence({ blob: SOURCE, masterBlob: marked() }), SOURCE);
});

function storage() {
  const data = {};
  browser.storage = { local: {
    async get(defaults) { return structuredClone({ ...defaults, ...data }); },
    async set(values) { Object.assign(data, structuredClone(values)); },
    async remove(key) { delete data[key]; },
  } };
  return data;
}

test("invitation lock serializes same-item callbacks and passes through their results", async () => {
  let release;
  let started;
  const gate = new Promise((resolve) => { release = resolve; });
  const firstStarted = new Promise((resolve) => { started = resolve; });
  const order = [];
  const first = withClientInvitationLock("calendar", "uid", async () => {
    order.push("first-start");
    started();
    await gate;
    order.push("first-end");
    return 1;
  });
  await firstStarted;
  const second = withClientInvitationLock("calendar", "uid", async () => {
    order.push("second");
    return 2;
  });
  await Promise.resolve();
  assert.deepEqual(order, ["first-start"]);
  release();
  assert.deepEqual(await Promise.all([first, second]), [1, 2]);
  assert.deepEqual(order, ["first-start", "first-end", "second"]);
});

test("invitation lock permits other items and calendars while one callback waits", async () => {
  let release;
  let started;
  const gate = new Promise((resolve) => { release = resolve; });
  const firstStarted = new Promise((resolve) => { started = resolve; });
  const first = withClientInvitationLock("calendar", "uid", async () => { started(); await gate; });
  await firstStarted;
  assert.equal(await withClientInvitationLock("calendar", "another", () => "other-item"), "other-item");
  assert.equal(await withClientInvitationLock("another-calendar", "uid", () => "other-calendar"), "other-calendar");
  release();
  await first;
});

test("failed invitation callback releases the lock and later callers still run", async () => {
  const error = new Error("callback failed");
  const failed = withClientInvitationLock("calendar", "uid", () => { throw error; });
  const next = withClientInvitationLock("calendar", "uid", () => "recovered");
  await assert.rejects(failed, (cause) => cause === error);
  assert.equal(await next, "recovered");
  assert.equal(await withClientInvitationLock("calendar", "uid", () => "fresh"), "fresh");
});

test("same-key callbacks queued inside an operation are not reentrant", async () => {
  const order = [];
  let nested;
  await withClientInvitationLock("calendar", "uid", () => {
    order.push("outer-start");
    nested = withClientInvitationLock("calendar", "uid", () => { order.push("nested"); });
    order.push("outer-end");
  });
  await nested;
  assert.deepEqual(order, ["outer-start", "outer-end", "nested"]);
});

test("durable sources survive module restart and missing new stamps; delete reclaims records", async () => {
  const data = storage();
  const source = marked();
  await rememberClientInvitation("calendar", "original-uid", source, "server-one");
  assert.deepEqual(await clientInvitationFor("calendar", "original-uid"), { itemId: "original-uid", blob: source, serverID: "server-one" });
  await rememberClientInvitation("calendar", "original-uid", source.replace("SUMMARY:Original invitation", "SUMMARY:New mail"));
  const restarted = await import(`../../src/modules/eas/client-invitations.mjs?restart=${Date.now()}`);
  assert.equal((await restarted.clientInvitationByServerId("calendar", "server-one")).itemId, "original-uid");
  assert.match((await restarted.clientInvitationFor("calendar", "original-uid")).blob, /SUMMARY:New mail/);
  assert.equal(await clientInvitationFor("another-calendar", "original-uid"), null);
  assert.equal(await clientInvitationByServerId("calendar", "missing"), null);
  await forgetClientInvitation("calendar", "original-uid");
  assert.equal(await clientInvitationFor("calendar", "original-uid"), null);
  assert.deepEqual(data, {});
});

test("concurrent saves retain both sources; removing a binding clears only that calendar", async () => {
  storage();
  const source = marked();
  await Promise.all([rememberClientInvitation("one", "first", source, "a"), rememberClientInvitation("one", "second", source, "b"), rememberClientInvitation("two", "first", source, "a")]);
  assert.equal((await clientInvitationFor("one", "first")).serverID, "a");
  assert.equal((await clientInvitationFor("one", "second")).serverID, "b");
  await forgetClientInvitation("one");
  assert.equal(await clientInvitationFor("one", "first"), null);
  assert.ok(await clientInvitationFor("two", "first"));
  await rememberClientInvitation("two", "unmarked", SOURCE);
  assert.equal(await clientInvitationFor("two", "unmarked"), null);
});
