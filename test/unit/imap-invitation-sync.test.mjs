import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
installWebextEnv();
const storage = {};
browser.storage = { local: {
  get: async (defaults) => Object.fromEntries(Object.entries(defaults).map(([key, value]) => [key, structuredClone(storage[key] ?? value)])),
  set: async (values) => Object.assign(storage, structuredClone(values)),
  remove: async (key) => { delete storage[key]; },
} };
const { applyResponses, applyServerCommands, clientInvitationCodec } =
  await import("../../src/modules/eas/sync-runner.mjs");
const codec = await import("../../src/modules/eas/calendar-codec.mjs");
const { rememberClientInvitation, clientInvitationFor } =
  await import("../../src/modules/eas/client-invitations.mjs");
const { createServerIdIndex } = await import("../../src/modules/eas/server-id-index.mjs");
const { createWBXML, decodeWBXML } = await import("../../src/modules/wbxml.mjs");
await (await import("../../src/modules/eas/timezone-mapping.mjs")).ensureLoaded();

const UID = "external-original@example.invalid";
const source = (sequence = 1, start = "20261022T100000Z") => [
  "BEGIN:VCALENDAR", "VERSION:2.0", "BEGIN:VEVENT", `UID:${UID}`,
  "X-EAS-CLIENT-INVITATION:1", "SUMMARY:Interview", "DTSTAMP:20261021T080000Z",
  `SEQUENCE:${sequence}`, `X-MOZ-RECEIVED-SEQUENCE:${sequence}`, `DTSTART:${start}`,
  "DTEND:20261022T120000Z", "ORGANIZER:mailto:organizer@example.invalid",
  "ATTENDEE;PARTSTAT=ACCEPTED:mailto:user@example.invalid",
  "X-MOZ-INVITED-ATTENDEE:mailto:user@example.invalid", "END:VEVENT", "END:VCALENDAR",
].join("\r\n");
const baseCodec = {
  readEasServerIdFromBlob: codec.readEasServerIdFromIcal,
  stampEasServerId: codec.stampEasServerId,
  applyInstanceChange: codec.applyInstanceChange,
  applyInstanceDelete: codec.applyInstanceDelete,
  appendApplicationDataFromBlob: ({ blob, ...options }) =>
    codec.appendApplicationDataFromIcal({ ...options, ical: blob }),
  applicationDataToBlob: ({ existingBlob, ...options }) =>
    codec.applicationDataToIcal({ ...options, existingIcal: existingBlob }),
};
let serial = 0;
function harness() {
  const items = new Map([[UID, source()]]);
  const entry = { parentId: `imap-sync-${++serial}`, itemId: UID, kind: "event", status: "added_by_user" };
  const pending = [entry];
  const ctx = {
    targetID: entry.parentId, folderId: "f-7", accountId: "1", asVersion: "16.1",
    account: { custom: { user: "user@example.invalid" } },
    itemKind: { changelogKind: "event", className: "Calendar", codec: clientInvitationCodec(baseCodec) },
    defaultTimezone: "UTC", syncRecurrence: true, separator: ",", msTodoCompat: false,
    indexMap: createServerIdIndex([]), uidClaims: new Map(), pendingAtStart: [],
    eventLog: () => {}, provider: { reportEventLog: () => {}, updateAccount: async () => {} },
    store: {
      get: async (id) => items.has(id) ? { id, blob: items.get(id) } : null,
      list: async () => [...items].map(([id, blob]) => ({ id, blob })),
      update: async (id, blob) => { assert.ok(items.has(id)); items.set(id, blob); },
      create: async (id, blob) => { assert.ok(!items.has(id)); items.set(id, blob); return id; },
      delete: async (id) => items.delete(id),
    },
    queue: {
      pending: async () => [...pending], markServerWrite: async () => {},
      remove: async ({ itemId }) => { const i = pending.findIndex((e) => e.itemId === itemId); if (i >= 0) pending.splice(i, 1); },
      record: async ({ op, ...key }) => {
        pending.splice(0); pending.push({ ...key, status: `${op === "updated" ? "modified" : op === "deleted" ? "deleted" : "added"}_by_user` });
      },
    },
  };
  const sent = () => ({ adds: [{ clientId: "c1", item: { id: UID, blob: source() }, entry }], mods: [], dels: [] });
  const ack = parseAdNode("<Add><ClientId>c1</ClientId><ServerId>7:42</ServerId><Status>1</Status></Add>");
  return { ctx, items, entry, pending, sent, ack };
}
const commands = (kind, xml) => ({ adds: [], changes: [], deletes: [], softDeletes: [], [kind]: [parseAdNode(xml)] });
const serverAdd = "<Add><ServerId>7:42</ServerId><ApplicationData><UID>server-generated-uid</UID><Subject>Server title</Subject><StartTime>20261022T110000Z</StartTime><MeetingStatus>0</MeetingStatus><OrganizerEmail>user%40example.invalid</OrganizerEmail><Attendees/></ApplicationData></Add>";

test("external invitation wire never re-invites attendees; ClientUid is creation-only", () => {
  for (const op of ["add", "change"]) {
    const builder = createWBXML(); builder.switchpage("AirSync"); builder.otag("ApplicationData");
    clientInvitationCodec(baseCodec).appendApplicationDataFromBlob({ builder, blob: source(), op, asVersion: "16.1", defaultTimezone: "UTC", syncRecurrence: true });
    builder.ctag(); const xml = decodeWBXML(builder.getBytes());
    assert.match(xml, /<MeetingStatus(?:\s[^>]*)?>0<\/MeetingStatus>/);
    assert.doesNotMatch(xml, /<Attendee(?:\s|>)|<OrganizerEmail(?:\s|>)|<OrganizerName(?:\s|>)/);
    assert.equal(xml.includes("<ClientUid"), op === "add");
  }
});

test("Add ACK, server update and Teams deletion retain original identity without resurrection", async () => {
  const f = harness();
  await applyResponses(f.ctx, { adds: [f.ack], changes: [], deletes: [] }, f.sent(), new Set());
  assert.equal(f.ctx.indexMap.serverIdFor(UID), "7:42");
  await applyServerCommands(f.ctx, commands("adds", serverAdd));
  const blob = f.items.get(UID);
  assert.match(blob, /UID:external-original@example.invalid/);
  assert.match(blob, /ORGANIZER:mailto:organizer@example.invalid/);
  assert.match(blob, /SEQUENCE:1/);
  assert.match(blob, /DTSTART:20261022T110000Z/);
  assert.equal(f.items.size, 1);
  await applyServerCommands(f.ctx, commands("deletes", "<Delete><ServerId>7:42</ServerId></Delete>"));
  assert.equal(f.items.size, 0); assert.equal(f.pending.length, 0);
  assert.equal(await clientInvitationFor(f.ctx.targetID, UID), null);
});

test("cache reset recovers original UID only when server offers the bound item", async () => {
  const f = harness();
  await rememberClientInvitation(f.ctx.targetID, UID, source(), "7:42");
  f.items.clear(); f.pending.splice(0); f.ctx.indexMap = createServerIdIndex([]);
  assert.equal(f.items.size, 0);
  await applyServerCommands(f.ctx, commands("adds", serverAdd));
  assert.ok(f.items.has(UID)); assert.equal(f.items.size, 1);
  assert.match(f.items.get(UID), /ORGANIZER:mailto:organizer@example.invalid/);
});

test("accepted newer REQUEST during Add and following stale pull is not lost", async () => {
  const f = harness(); f.items.set(UID, source(2, "20261022T113000Z"));
  await applyResponses(f.ctx, { adds: [f.ack], changes: [], deletes: [] }, f.sent(), new Set());
  assert.equal(f.pending[0].status, "modified_by_user");
  await applyServerCommands(f.ctx, commands("adds", serverAdd));
  assert.match(f.items.get(UID), /SEQUENCE:2/);
  assert.match(f.items.get(UID), /DTSTART:20261022T113000Z/);
});

test("a source Change refused as missing never becomes Add after a Teams deletion", async () => {
  const f = harness(); const stamped = codec.stampEasServerId(source(2), "7:42");
  f.items.set(UID, stamped); f.ctx.indexMap.set(UID, "7:42");
  await rememberClientInvitation(f.ctx.targetID, UID, stamped, "7:42");
  const sent = { adds: [], mods: [{ entry: { ...f.entry, status: "modified_by_user" }, serverID: "7:42", item: { id: UID, blob: stamped } }], dels: [] };
  await applyResponses(f.ctx, { adds: [], changes: [parseAdNode("<Change><ServerId>7:42</ServerId><Status>8</Status></Change>")], deletes: [] }, sent, new Set());
  assert.equal(f.items.size, 0); assert.equal(f.pending.length, 0);
  assert.equal(f.ctx.indexMap.serverIdFor(UID), null);
});

test("server Change cannot recreate a locally deleted source appointment", async () => {
  const f = harness(); await rememberClientInvitation(f.ctx.targetID, UID, source(), "7:42");
  f.items.clear(); f.pending[0].status = "deleted_by_user";
  await applyServerCommands(f.ctx, commands("changes", serverAdd.replaceAll("Add", "Change")));
  assert.equal(f.items.size, 0); assert.equal(f.pending[0].status, "deleted_by_user");
});

test("a higher-sequence mail is pushed as Change to the original server item", async () => {
  const f = harness();
  await applyResponses(f.ctx, { adds: [f.ack], changes: [], deletes: [] }, f.sent(), new Set());
  const updated = codec.stampEasServerId(source(2, "20261022T113000Z"), "7:42");
  f.items.set(UID, updated); f.pending.push({ ...f.entry, status: "modified_by_user" });
  await rememberClientInvitation(f.ctx.targetID, UID, updated, "7:42");
  const sent = { adds: [], mods: [{ entry: f.pending[0], serverID: f.ctx.indexMap.serverIdFor(UID), item: { id: UID, blob: updated } }], dels: [] };
  await applyResponses(f.ctx, { adds: [], changes: [], deletes: [] }, sent, new Set());
  assert.equal(f.pending.length, 0); assert.equal(f.items.size, 1);
  assert.equal(f.ctx.indexMap.serverIdFor(UID), "7:42");
  assert.match(f.items.get(UID), /SEQUENCE:2/);
  assert.match(f.items.get(UID), /DTSTART:20261022T113000Z/);
});

test("16.1 instance update preserves invitation identity and updates durable source", async () => {
  const f = harness(); f.pending.splice(0);
  const series = source().replace("SUMMARY:Interview", "RRULE:FREQ=DAILY;COUNT=3\r\nSUMMARY:Interview");
  f.items.set(UID, codec.stampEasServerId(series, "7:42")); f.ctx.indexMap.set(UID, "7:42");
  await applyServerCommands(f.ctx, commands("changes", "<Change><ServerId>7:42</ServerId><ApplicationData><InstanceId>20261023T100000Z</InstanceId><StartTime>20261023T110000Z</StartTime><EndTime>20261023T120000Z</EndTime><OrganizerEmail>user%40example.invalid</OrganizerEmail><Attendees/><DtStamp>20261021T090000Z</DtStamp></ApplicationData></Change>"));
  const blob = f.items.get(UID);
  assert.equal((blob.match(/ORGANIZER:mailto:organizer@example.invalid/g) ?? []).length, 2);
  assert.equal((blob.match(/PARTSTAT=ACCEPTED/g) ?? []).length, 2);
  assert.match(blob, /DTSTART:20261023T110000Z/);
  assert.equal((await clientInvitationFor(f.ctx.targetID, UID)).blob, blob);
});

test("16.1 instance update cannot overwrite an accepted mail still pending push", async () => {
  const f = harness(); const series = source(2).replace("SUMMARY:Interview", "RRULE:FREQ=DAILY;COUNT=3\r\nSUMMARY:Interview");
  const stamped = codec.stampEasServerId(series, "7:42");
  f.items.set(UID, stamped); f.ctx.indexMap.set(UID, "7:42");
  await applyServerCommands(f.ctx, commands("changes", "<Change><ServerId>7:42</ServerId><ApplicationData><InstanceId>20261023T100000Z</InstanceId><StartTime>20261023T080000Z</StartTime></ApplicationData></Change>"));
  assert.equal(f.items.get(UID), stamped);
});

test("server deletion removes durable source even after the cache was reset", async () => {
  const f = harness(); await rememberClientInvitation(f.ctx.targetID, UID, source(), "7:42");
  f.items.clear(); f.pending.splice(0);
  await applyServerCommands(f.ctx, commands("deletes", "<Delete><ServerId>7:42</ServerId></Delete>"));
  assert.equal(await clientInvitationFor(f.ctx.targetID, UID), null);
  assert.equal(f.items.size, 0);
});

test("SoftDelete preserves source identity for a later return into the sync window", async () => {
  const f = harness();
  await applyResponses(f.ctx, { adds: [f.ack], changes: [], deletes: [] }, f.sent(), new Set());
  await applyServerCommands(f.ctx, commands("softDeletes", "<SoftDelete><ServerId>7:42</ServerId></SoftDelete>"));
  assert.equal(f.items.size, 0);
  assert.ok(await clientInvitationFor(f.ctx.targetID, UID));
  await applyServerCommands(f.ctx, commands("adds", serverAdd));
  assert.equal(f.items.size, 1);
  assert.match(f.items.get(UID), /UID:external-original@example.invalid/);
});

test("SoftDelete cannot acknowledge a pending user deletion or newer accepted mail", async () => {
  for (const status of ["deleted_by_user", "modified_by_user"]) {
    const f = harness();
    await rememberClientInvitation(f.ctx.targetID, UID, source(2), "7:42");
    f.ctx.indexMap.set(UID, "7:42"); f.pending[0].status = status;
    if (status === "deleted_by_user") f.items.clear();
    await applyServerCommands(f.ctx, commands("softDeletes", "<SoftDelete><ServerId>7:42</ServerId></SoftDelete>"));
    assert.equal(f.pending[0].status, status);
    assert.ok(await clientInvitationFor(f.ctx.targetID, UID));
    assert.equal(f.ctx.indexMap.serverIdFor(UID), "7:42");
  }
});

test("invalid instance id on a marked meeting finishes without reacquiring its lock", async () => {
  const f = harness(); f.pending.splice(0);
  f.ctx.indexMap.set(UID, "7:42");
  await applyServerCommands(f.ctx, commands("changes", "<Change><ServerId>7:42</ServerId><ApplicationData><InstanceId>invalid</InstanceId><Subject>Updated title</Subject></ApplicationData></Change>"));
  assert.match(f.items.get(UID), /SUMMARY:Updated title/);
});
