/** Plain events must use ordinary Sync even when IMAP invitations are enabled.
 * Exercises the real dispatcher, codec, queue and response handler offline. */
import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
installWebextEnv();
const storage = {};
browser.storage = { local: {
  get: async (defaults) => Object.fromEntries(Object.entries(defaults).map(([k, v]) => [k, structuredClone(storage[k] ?? v)])),
  set: async (values) => Object.assign(storage, structuredClone(values)),
  remove: async (key) => { delete storage[key]; },
} };
function doc(xml) {
  const root = parseAdNode(xml);
  const nodes = [];
  (function walk(n) { nodes.push(n); for (const c of n.children) walk(c); })(root);
  return { documentElement: root, getElementsByTagName: (tag) => nodes.filter((n) => n.tagName === tag) };
}
const requests = [];
let addStatus = "1";
const NETWORK = new URL("../../src/modules/network.mjs", import.meta.url);
const realNetwork = await import(NETWORK.href);
const { decodeWBXML } = await import("../../src/modules/wbxml.mjs");
mock.module(NETWORK.href, { namedExports: {
  ...realNetwork,
  easRequest: async ({ command, body }) => {
    assert.equal(command, "Sync", "a plain event never needs Search, MeetingResponse or mail");
    const xml = decodeWBXML(body);
    requests.push(xml);
    const clientId = /<ClientId>([^<]+)<\/ClientId>/.exec(xml)?.[1];
    const response = clientId ? `<Responses><Add><ClientId>${clientId}</ClientId><ServerId>7:plain</ServerId><Status>${addStatus}</Status></Add></Responses>` : "";
    return { doc: doc(`<Sync><Collections><Collection><Status>1</Status><SyncKey>${requests.length + 10}</SyncKey>${response}</Collection></Collections></Sync>`) };
  },
} });
const { runItemSync } = await import("../../src/modules/eas/sync-runner.mjs");
const codec = await import("../../src/modules/eas/calendar-codec.mjs");
const { localQueue } = await import("../../src/vendor/tbsync/change-queue.mjs");
const { rememberBindings } = await import("../../src/vendor/tbsync/change-queue.mjs");
mock.module(new URL("../../src/modules/sync-after-change.mjs", import.meta.url).href, {
  namedExports: { armSyncAfterChange() {}, quietSecondsFor: async () => 0 },
});
const hooks = {};
const listener = (name) => ({ addListener(callback) { hooks[name] = callback; } });
messenger.calendar.provider = Object.fromEntries([
  "onItemCreated", "onItemUpdated", "onItemRemoved", "onSync", "onResetSync",
].map(name => [name, listener(name)]));
messenger.calendar.calendars = { onUpdated: listener("calendarUpdated"), onRemoved: listener("calendarRemoved") };
let hookPrior = null;
messenger.calendar.items = { get: async () => hookPrior ? { item: hookPrior } : null };
const { registerCalendarProvider } = await import("../../src/modules/calendar-provider.mjs");
registerCalendarProvider();
await (await import("../../src/modules/eas/timezone-mapping.mjs")).ensureLoaded();

const UID = "plain-event@example.invalid";
const PLAIN = ["BEGIN:VCALENDAR", "VERSION:2.0", "BEGIN:VEVENT", `UID:${UID}`,
  "DTSTAMP:20261021T080000Z", "DTSTART:20261022T100000Z", "DTEND:20261022T110000Z",
  "SUMMARY:Plain event", "END:VEVENT", "END:VCALENDAR"].join("\r\n");
let serial = 0;
async function fixture({ op = "created", enabled = false, blob = PLAIN } = {}) {
  requests.length = 0; addStatus = "1";
  const sessionId = `plain-session-${++serial}`;
  const folder = { targetID: `plain-calendar-${serial}`, targetType: "calendar", sessionId,
    custom: { serverID: "calendar-collection", synckey: "10", type: "8", imapInvitations: enabled } };
  storage[`eas.imapInvitations.${folder.targetID}`] = enabled ? { user: "user@example.invalid" } : null;
  const queue = localQueue({ accountId: "1", folderId: "f-7", sessionId, observed: false });
  await queue.record({ parentId: folder.targetID, itemId: UID, kind: "event", op });
  let current = blob;
  const events = [];
  const patches = [];
  const options = {
    accountId: "1", folderId: "f-7", folder, asVersion: "16.1", defaultTimezone: "UTC",
    account: { custom: { user: "user@example.invalid", imapInvitations: enabled, syncrecurrence: true, allowedEasCommands: ["Sync"] } },
    provider: { reportProgress: () => {}, reportEventLog: (e) => events.push(e),
      updateFolder: async ({ patch }) => patches.push(patch) },
    itemKind: { changelogKind: "event", className: "Calendar", filterType: "0",
      codec: {
        readEasServerIdFromBlob: codec.readEasServerIdFromIcal,
        stampEasServerId: codec.stampEasServerId,
        clientRejectReason: codec.clientRejectReason,
        appendApplicationDataFromBlob: ({ blob, ...rest }) => codec.appendApplicationDataFromIcal({ ...rest, ical: blob }),
      },
      storeFactory: () => ({
        get: async () => ({ id: UID, blob: current }),
        list: async () => [{ id: UID, blob: current }],
        update: async (id, next) => { assert.equal(id, UID); current = next; },
      }),
    },
  };
  return { options, queue, events, patches, blob: () => current };
}

for (const enabled of [false, true]) {
  test(`plain create uses Sync Add and is acknowledged (IMAP option ${enabled})`, async () => {
    const f = await fixture({ enabled });
    assert.equal(codec.isReceivedMeeting(PLAIN, "user@example.invalid"), false);
    await runItemSync(f.options);
    assert.equal(requests.length, 2, "one push and one pull");
    assert.match(requests[0], /<Add>/);
    assert.match(requests[0], /<Subject(?:\s[^>]*)?>Plain%20event<\/Subject>/);
    assert.doesNotMatch(requests[0], /<Attendee(?:\s|>)/);
    assert.equal(codec.readEasServerIdFromIcal(f.blob()), "7:plain");
    assert.equal((await f.queue.pending()).length, 0);
    assert.ok(!f.events.some((e) => /holding the answer/.test(e.message)));
    assert.ok(f.patches.some((p) => p.custom?.indexMap?.some((row) => row.uid === UID && row.serverId === "7:plain")));
  });
}

test("explicit opt-in resumes an already accepted unsent invitation without replying from sync", async () => {
  for (const op of ["created", "updated"]) {
    const invitation = PLAIN.replace("SUMMARY:Plain event", [
      "SUMMARY:Plain event", "ORGANIZER:mailto:external@example.invalid",
      "ATTENDEE;PARTSTAT=ACCEPTED:mailto:user@example.invalid",
      "X-MOZ-INVITED-ATTENDEE:mailto:user@example.invalid", "SEQUENCE:4",
    ].join("\r\n"));
    const f = await fixture({ op, enabled: true, blob: invitation });
    await runItemSync(f.options);
    assert.match(requests[0], /<Add>/);
    assert.doesNotMatch(requests[0], /<Attendee(?:\s|>)|<OrganizerEmail(?:\s|>)/);
    assert.match(f.blob(), /UID:plain-event@example.invalid/);
    assert.match(f.blob(), /ORGANIZER:mailto:external@example.invalid/);
    assert.match(f.blob(), /SEQUENCE:4/);
    assert.match(f.blob(), /X-EAS-CLIENT-INVITATION:1/);
    assert.equal(codec.readEasServerIdFromIcal(f.blob()), "7:plain");
    assert.equal((await f.queue.pending()).length, 0);
  }
});

test("a plain change without stamp or index remains queued for recovery", async () => {
  const f = await fixture({ op: "updated" });
  await runItemSync(f.options);
  assert.equal(requests.length, 1, "only the pull is sent: Change has no address");
  assert.doesNotMatch(requests[0], /<Add>|<Change>/);
  assert.equal((await f.queue.pending()).length, 1);
  assert.equal(f.blob(), PLAIN);
  assert.ok(f.events.some((e) => e.level === "warning" && /no ServerId.*keeping the local change/.test(e.message)));
  assert.ok(!f.events.some((e) => /holding the answer/.test(e.message)));
});

test("opt-in cannot shorten an old accepted invitation with multiple recurrence rules", async () => {
  const blob = PLAIN.replace("SUMMARY:Plain event", [
    "SUMMARY:Plain event", "ORGANIZER:mailto:external@example.invalid",
    "ATTENDEE;PARTSTAT=ACCEPTED:mailto:user@example.invalid",
    "X-MOZ-INVITED-ATTENDEE:mailto:user@example.invalid",
    "RRULE:FREQ=WEEKLY;COUNT=3", "RRULE:FREQ=MONTHLY;COUNT=3",
  ].join("\r\n"));
  const f = await fixture({ enabled: true, blob });
  await runItemSync(f.options);
  assert.equal(requests.length, 1);
  assert.doesNotMatch(requests[0], /<Add>|<Change>/);
  assert.equal(f.blob(), blob);
  assert.equal((await f.queue.pending()).length, 1);
  assert.ok(f.events.some(event => /recurrence cannot be stored/.test(event.message)));
});

test("the real dialog create hook queues an id-less plain event as added, and an edit before push stays added", async () => {
  const f = await fixture({ enabled: true });
  await f.queue.remove({ parentId: f.options.folder.targetID, itemId: UID, kind: "event" });
  await rememberBindings([{ targetID: f.options.folder.targetID, accountId: "1", folderId: "f-7",
    sessionId: f.options.folder.sessionId, targetType: "calendars" }]);
  const calendar = { id: f.options.folder.targetID };
  const input = { id: null, type: "event", format: "ical", item: PLAIN };
  hookPrior = null;
  const created = await hooks.onItemCreated(calendar, input, {});
  assert.ok(created.id);
  assert.equal(codec.isReceivedMeeting(created.item, "user@example.invalid"), false);
  let pending = await f.queue.pending();
  assert.equal(pending.length, 1);
  assert.equal(pending[0].itemId, created.id);
  assert.equal(pending[0].status, "added_by_user");
  hookPrior = created.item;
  await hooks.onItemUpdated(calendar, { ...created, item: created.item.replace("SUMMARY:Plain event", "SUMMARY:Edited plain event") }, created, {});
  pending = await f.queue.pending();
  assert.equal(pending.length, 1);
  assert.equal(pending[0].status, "added_by_user");
  hookPrior = null;
});

test("a server-rejected plain Add retains its queued edit", async () => {
  const f = await fixture(); addStatus = "6";
  await runItemSync(f.options);
  assert.equal(requests.length, 2);
  assert.match(requests[0], /<Add>/);
  assert.equal((await f.queue.pending()).length, 1);
  assert.equal(codec.readEasServerIdFromIcal(f.blob()), null);
  assert.ok(f.events.some((e) => /Status 6/.test(e.message)));
});
