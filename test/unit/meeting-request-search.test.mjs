/** Exercise Search -> exact UID matching -> real MeetingResponse encoding.
 * Only the network boundary is replaced; no server or user profile is used. */
import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
installWebextEnv();

function doc(xml) {
  const root = parseAdNode(xml);
  const nodes = [];
  (function walk(n) { nodes.push(n); for (const child of n.children) walk(child); })(root);
  return { documentElement: root, getElementsByTagName: (tag) => nodes.filter((n) => n.tagName === tag) };
}

const calls = [];
let searchXml;
let meetingXml;
let failCommand;
const NETWORK = new URL("../../src/modules/network.mjs", import.meta.url);
const realNetwork = await import(NETWORK.href);
mock.module(NETWORK.href, {
  namedExports: {
    ...realNetwork,
    easRequest: async (request) => {
      calls.push(request);
      assert.ok(["Search", "MeetingResponse"].includes(request.command), "no Sync/Add or SendMail allowed");
      if (request.command === failCommand) throw new Error("offline transport failure");
      return { doc: doc(request.command === "Search" ? searchXml : meetingXml) };
    },
  },
});

const { invitationPhase } = await import("../../src/modules/eas/sync-runner.mjs");
const { buildMeetingRequestSearchBody, readMeetingRequestLongId, findMeetingRequestLongId } =
  await import("../../src/modules/eas/meeting-request-search.mjs");
const { readEasServerIdFromIcal, stampEasServerId } = await import("../../src/modules/eas/calendar-codec.mjs");
const { createServerIdIndex } = await import("../../src/modules/eas/server-id-index.mjs");
const { decodeWBXML } = await import("../../src/modules/wbxml.mjs");

const ME = "user@example.invalid";
const UID = "#meeting@u@external-calendar";
const LONG_ID = "mail-id:opaque/one%2F";
const BLOB = [
  "BEGIN:VCALENDAR", "VERSION:2.0", "BEGIN:VEVENT", `UID:${UID}`,
  "DTSTAMP:20261021T080000Z", "DTSTART:20261022T100000Z", "DTEND:20261022T110000Z",
  "SUMMARY:Interview", "ORGANIZER:mailto:organizer@example.invalid",
  `ATTENDEE;PARTSTAT=ACCEPTED:mailto:${ME}`, `X-MOZ-INVITED-ATTENDEE:mailto:${ME}`,
  "END:VEVENT", "END:VCALENDAR",
].join("\r\n");

function row({ uid = UID, longId = LONG_ID, organizer = "organizer@example.invalid",
  start = "2026-10-22T10:00:00.000Z", messageClass = "IPM.Schedule.Meeting.Request",
  type = "1" } = {}) {
  return `<Result><Class>Email</Class><LongId>${encodeURIComponent(longId)}</LongId><Properties>` +
    `<MessageClass>${messageClass}</MessageClass><MeetingRequest>` +
    `<UID>${encodeURIComponent(uid)}</UID><MeetingMessageType>${type}</MeetingMessageType>` +
    (organizer ? `<Organizer>${encodeURIComponent(organizer)}</Organizer>` : "") +
    (start ? `<StartTime>${encodeURIComponent(start)}</StartTime>` : "") +
    `</MeetingRequest></Properties></Result>`;
}

function reply(rows = [row()], { total = rows.length, status = "1", storeStatus = "1" } = {}) {
  return `<Search><Status>${status}</Status><Response><Store><Status>${storeStatus}</Status>` +
    rows.join("") + (total === null ? "" : `<Total>${total}</Total>`) + `</Store></Response></Search>`;
}

function fixture() {
  calls.length = 0;
  failCommand = null;
  searchXml = reply();
  meetingXml = "<MeetingResponse><Result><Status>1</Status></Result></MeetingResponse>";
  const entry = { parentId: "calendar", itemId: UID, kind: "event", status: "added_by_user", timestamp: Date.now() - 10_800_000 };
  const pending = [entry];
  const item = { blob: BLOB };
  const writes = [];
  const events = [];
  const ctx = {
    account: { custom: { user: ME, allowedEasCommands: ["Search", "MeetingResponse"] } },
    asVersion: "16.1", accountId: "account", folderId: "folder", collectionId: "calendar-collection", targetID: "calendar",
    folder: { custom: { type: "8" } },
    itemKind: { changelogKind: "event", codec: { readEasServerIdFromBlob: readEasServerIdFromIcal, stampEasServerId } },
    indexMap: createServerIdIndex(),
    store: {
      get: async () => item,
      update: async (id, blob) => { assert.equal(id, UID); writes.push(blob); item.blob = blob; },
      create: async () => assert.fail("must not create a plain event"),
    },
    queue: {
      remove: async () => pending.splice(0),
      markServerWrite: async (tag) => assert.equal(tag.itemId, UID),
    },
    provider: { reportEventLog: (event) => events.push(event) },
    eventLog: (level, message) => events.push({ level, message }),
  };
  return { ctx, entry, pending, item, writes, events };
}

test("the bounded mailbox query searches Email by subject across folders", () => {
  const body = buildMeetingRequestSearchBody("Interview");
  const xml = decodeWBXML(body);
  assert.match(xml, /<Name>Mailbox<\/Name>/);
  assert.match(xml, /<Query><And><Class xmlns='AirSync'>Email<\/Class><FreeText>Interview<\/FreeText><\/And><\/Query>/);
  assert.match(xml, /<DeepTraversal\s*\/>/);
  assert.match(xml, /<Range>0-99<\/Range>/);
  assert.ok(Buffer.from(body).includes(Buffer.from([0, 0, 0x50])), "Class uses AirSync page 0");
});

test("Search finds the source email; its LongId is sent once and never stored as calendar identity", async () => {
  const f = fixture();
  await invitationPhase(f.ctx, [f.entry]);
  assert.deepEqual(calls.map((c) => c.command), ["Search", "MeetingResponse"]);
  const xml = decodeWBXML(calls[1].body);
  assert.match(xml, new RegExp(`<LongId xmlns='Search'>${encodeURIComponent(LONG_ID)}<\\/LongId>`));
  assert.doesNotMatch(xml, /CollectionId|RequestId/);
  assert.match(xml, /<SendResponse\s*\/>/);
  assert.equal(f.pending.length, 0);
  assert.equal(f.item.blob, BLOB);
  assert.deepEqual(f.writes, []);
  assert.equal(f.ctx.indexMap.size, 0);
});

test("only an acknowledged CalendarId is stamped and indexed", async () => {
  const f = fixture();
  meetingXml = "<MeetingResponse><Result><Status>1</Status><CalendarId>calendar-server-id</CalendarId></Result></MeetingResponse>";
  await invitationPhase(f.ctx, [f.entry]);
  assert.equal(readEasServerIdFromIcal(f.item.blob), "calendar-server-id");
  assert.equal(f.ctx.indexMap.serverIdFor(UID), "calendar-server-id");
  assert.equal(f.pending.length, 0);
  assert.equal(f.writes.length, 1);
  assert.ok(!f.item.blob.includes(LONG_ID));
});

for (const [label, rows, options] of [
  ["no results", []],
  ["different UID", [row({ uid: "another-uid" })]],
  ["multiple matching emails", [row(), row({ longId: "another-mail" })]],
  ["same UID on a stale invitation and its update", [row(), row({ longId: "older-mail", start: "2026-10-20T10:00:00.000Z" })]],
  ["cancellation", [row({ messageClass: "IPM.Schedule.Meeting.Canceled" })]],
  ["reply", [row({ messageClass: "IPM.Schedule.Meeting.Resp.Pos" })]],
  ["non-owner notification", [row({ type: "6" })]],
  ["missing message type", [row({ type: "" })]],
  ["different organizer", [row({ organizer: "other@example.invalid" })]],
  ["different start", [row({ start: "2026-10-23T10:00:00.000Z" })]],
  ["missing LongId", [row({ longId: "" })]],
  ["oversized LongId", [row({ longId: "a".repeat(257) })]],
  ["truncated results", [row()], { total: 101 }],
  ["unknown completeness", [row()], { total: null }],
  ["failed search", [row()], { status: "3" }],
  ["failed store", [row()], { storeStatus: "2" }],
]) {
  test(`${label}: no response and no invitation or queue changes`, async () => {
    const f = fixture();
    searchXml = reply(rows, options);
    assert.equal(readMeetingRequestLongId(doc(searchXml), BLOB), null);
    await invitationPhase(f.ctx, [f.entry]);
    assert.deepEqual(calls.map((c) => c.command), ["Search"]);
    assert.deepEqual(f.pending, [f.entry]);
    assert.equal(f.item.blob, BLOB);
    assert.deepEqual(f.writes, []);
    assert.equal(f.ctx.indexMap.size, 0);
  });
}

test("optional absent organizer/start does not block an exact, unique UID", () => {
  assert.equal(readMeetingRequestLongId(doc(reply([row({ organizer: null, start: null })])), BLOB), LONG_ID);
});

for (const command of ["Search", "MeetingResponse"]) {
  test(`${command} transport failure keeps the answer for retry`, async () => {
    const f = fixture();
    failCommand = command;
    await invitationPhase(f.ctx, [f.entry]);
    assert.deepEqual(f.pending, [f.entry]);
    assert.equal(f.item.blob, BLOB);
    assert.deepEqual(f.writes, []);
    assert.equal(f.ctx.indexMap.size, 0);
  });
}

test("a rejected source request keeps its queued answer", async () => {
  const f = fixture();
  meetingXml = "<MeetingResponse><Result><Status>2</Status><CalendarId>untrusted</CalendarId></Result></MeetingResponse>";
  await invitationPhase(f.ctx, [f.entry]);
  assert.deepEqual(f.pending, [f.entry]);
  assert.equal(f.item.blob, BLOB);
  assert.equal(f.ctx.indexMap.size, 0);
});

for (const commands of [undefined, [], ["MeetingResponse"]]) {
  test(`Search must be explicitly advertised (${JSON.stringify(commands)})`, async () => {
    const f = fixture();
    f.ctx.account.custom.allowedEasCommands = commands;
    await invitationPhase(f.ctx, [f.entry]);
    assert.deepEqual(calls, []);
    assert.deepEqual(f.pending, [f.entry]);
  });
}

test("UID search is not attempted on pre-16 protocols", async () => {
  const f = fixture();
  assert.equal(await findMeetingRequestLongId({ account: f.ctx.account, asVersion: "14.1", blob: BLOB }), null);
  assert.deepEqual(calls, []);
});

for (const type of ["13", undefined, "shared"]) {
  test(`source mail is not used for a non-default or unknown Calendar (${type})`, async () => {
    const f = fixture();
    f.ctx.folder.custom.type = type;
    await invitationPhase(f.ctx, [f.entry]);
    assert.deepEqual(calls, []);
    assert.deepEqual(f.pending, [f.entry]);
    assert.equal(f.item.blob, BLOB);
    assert.deepEqual(f.writes, []);
    assert.equal(f.ctx.indexMap.size, 0);
  });
}
