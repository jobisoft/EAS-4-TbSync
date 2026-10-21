/** Invitation responses use both sources of identity and never destroy an
 *  invitation merely because its ServerId has not become available.
 *  All transport is mocked; these tests cannot contact a server. */
import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

const NETWORK = new URL("../../src/modules/network.mjs", import.meta.url);
const realNetwork = await import(NETWORK.href);
mock.module(NETWORK.href, {
  namedExports: {
    ...realNetwork,
    easRequest: async () => assert.fail("no Sync/Add or other network request is permitted"),
  },
});

const calls = [];
let response;
mock.module(new URL("../../src/modules/eas/meeting-response.mjs", import.meta.url).href, {
  namedExports: {
    sendMeetingResponse: async (request) => {
      calls.push(request);
      return response;
    },
  },
});

const { invitationPhase } = await import("../../src/modules/eas/sync-runner.mjs");
const { readEasServerIdFromIcal, isReceivedMeeting } = await import(
  "../../src/modules/eas/calendar-codec.mjs"
);
const ME = "user@example.invalid";

function fixture({ stamp = null, indexed = null, ageMs = 0 } = {}) {
  calls.length = 0;
  response = { status: "1", calendarId: null };
  const entry = {
    parentId: "calendar-1", itemId: "invitation-1", kind: "event",
    status: "added_by_user", timestamp: Date.now() - ageMs,
  };
  const blob = [
    "BEGIN:VCALENDAR", "VERSION:2.0", "BEGIN:VEVENT", "UID:invitation-1",
    "DTSTAMP:20261021T080000Z", "DTSTART:20261022T100000Z",
    "DTEND:20261022T110000Z", "SUMMARY:Interview",
    "ORGANIZER:mailto:organizer@example.invalid",
    `ATTENDEE;PARTSTAT=ACCEPTED:mailto:${ME}`,
    `X-MOZ-INVITED-ATTENDEE:mailto:${ME}`,
    ...(stamp ? [`X-EAS-SERVERID:${stamp}`] : []),
    "END:VEVENT", "END:VCALENDAR",
  ].join("\r\n");
  const item = { blob };
  const pending = [entry];
  const removed = [];
  const events = [];
  const ctx = {
    account: { custom: { user: ME, allowedEasCommands: ["MeetingResponse"] } },
    accountId: "account-1", folderId: "folder-1", collectionId: "collection-1",
    folder: { custom: { type: "8" } },
    targetID: entry.parentId, asVersion: "16.1",
    itemKind: { changelogKind: "event", codec: { readEasServerIdFromBlob: readEasServerIdFromIcal } },
    indexMap: { serverIdFor: (id) => { assert.equal(id, entry.itemId); return indexed; } },
    store: {
      get: async (id) => { assert.equal(id, entry.itemId); return item; },
      update: async () => assert.fail("the invitation must not be rewritten"),
      create: async () => assert.fail("the invitation must not become a new event"),
    },
    queue: {
      remove: async (key) => { removed.push(key); pending.splice(0); },
      markServerWrite: async () => assert.fail("no server write is expected"),
    },
    provider: { reportEventLog: (event) => events.push(event) },
    eventLog: (level, message) => events.push({ level, message }),
  };
  return { ctx, entry, item, blob, pending, removed, events };
}

test("a missing stamp falls back to the index and ACK settles the answer", async () => {
  const f = fixture({ indexed: "index-server-id" });
  await invitationPhase(f.ctx, [f.entry]);
  assert.deepEqual(calls, [{
    account: f.ctx.account, asVersion: "16.1", collectionId: "collection-1",
    serverID: "index-server-id", userResponse: 1, instanceId: null,
  }]);
  assert.deepEqual(f.removed, [{ parentId: "calendar-1", itemId: "invitation-1", kind: "event" }]);
  assert.equal(f.pending.length, 0);
  assert.equal(f.item.blob, f.blob);
});

test("the item's stamp takes precedence over the index", async () => {
  const f = fixture({ stamp: "stamp-server-id", indexed: "stale-index-id" });
  await invitationPhase(f.ctx, [f.entry]);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].serverID, "stamp-server-id");
  assert.equal(f.pending.length, 0);
});

for (const failure of [null, { status: "3" }]) {
  test(`an unsuccessful response (${failure?.status ?? "no response"}) keeps the answer`, async () => {
    const f = fixture({ indexed: "index-server-id" });
    response = failure;
    await invitationPhase(f.ctx, [f.entry]);
    assert.equal(calls.length, 1);
    assert.deepEqual(f.pending, [f.entry]);
    assert.deepEqual(f.removed, []);
    assert.equal(f.item.blob, f.blob);
  });
}

for (const [ageMs, level] of [[60_000, "info"], [3 * 60 * 60 * 1000, "warning"]]) {
  test(`no identity after ${ageMs / 60_000} minutes preserves the invitation and answer`, async () => {
    const f = fixture({ ageMs });
    const originalEntry = { ...f.entry };
    // Repeated syncs cannot release this item into the ordinary Add path.
    await invitationPhase(f.ctx, [f.entry]);
    await invitationPhase(f.ctx, [f.entry]);
    assert.deepEqual(calls, []);
    assert.equal(f.item.blob, f.blob);
    assert.equal(isReceivedMeeting(f.item.blob, ME), true);
    assert.deepEqual(f.pending, [originalEntry]);
    assert.deepEqual(f.removed, []);
    assert.equal(f.events.length, 2);
    assert.ok(f.events.every((e) => e.level === level && /no ServerId/.test(e.message)));
    assert.ok(f.events.every((e) => !/plain event|holds no copy|guest list dropped/.test(e.message)));
  });
}
