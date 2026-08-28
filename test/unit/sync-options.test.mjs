/**
 * Unit tests for the `<Options>` block of a Sync request.
 *
 * The rule under test: [MS-ASCMD] keeps the last Options block per
 * collection and a new one REPLACES it. So a request that states one
 * option retires every other one, and the server carries that loss until
 * something states them again.
 *
 * This is not hypothetical. The push batch grew an Options block naming
 * only `Conflict`, which dropped `FilterType` - so every push reset the
 * calendar window to unfiltered and the next filtered pull deleted what it
 * had just added. Measured on Exchange Online 16.1: a push-triggered sync
 * pulled 45 items as adds with a six-month window, and none once the
 * window was set to "all". The commit that first diagnosed the mechanism
 * restored `BodyPreference` and still left `FilterType` out.
 *
 * Hence these tests assert the *whole* block on every request that writes
 * one, rather than the element that happened to be forgotten.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

import { buildSyncBody } from "../../src/modules/eas/sync-body.mjs";
import { decodeWBXML } from "../../src/modules/wbxml.mjs";

/** `appendCommands` writes nothing for an empty batch, which is what these
 *  tests want - the Options block is the subject, not the commands. */
const EMPTY_COMMANDS = { adds: [], mods: [], dels: [] };

const CALENDAR = {
  collectionId: "42",
  className: "Calendar",
  filterType: "7",
  conflict: "1",
  synckey: "abc",
};

/** The decoded request as text, which is enough to assert on presence and
 *  order without walking the tree - these are flat, short elements. */
function xml(overrides = {}) {
  return decodeWBXML(buildSyncBody({ ...CALENDAR, ...overrides }));
}

/** Every `<Options>` … `</Options>` run in the request. */
function optionBlocks(text) {
  return [...text.matchAll(/<Options>([\s\S]*?)<\/Options>/g)].map((m) => m[1]);
}

test("the pull states the whole set", () => {
  const blocks = optionBlocks(xml({ withChanges: true }));
  assert.equal(blocks.length, 1);
  const [o] = blocks;
  assert.match(o, /<FilterType>7<\/FilterType>/);
  assert.match(o, /<Class>Calendar<\/Class>/);
  assert.match(o, /<Conflict>1<\/Conflict>/);
  assert.match(o, /<BodyPreference[^>]*>[\s\S]*<Type[^>]*>1</);
});

test("the push states the whole set too - not just the conflict policy", () => {
  // The regression: this block used to carry Conflict and BodyPreference
  // and nothing else, which retired the window for every request after it.
  const blocks = optionBlocks(
    xml({ withChanges: false, withCommands: EMPTY_COMMANDS }),
  );
  assert.equal(blocks.length, 1, "a push that carries commands states options");
  const [o] = blocks;
  assert.match(
    o,
    /<FilterType>7<\/FilterType>/,
    "the window must survive a push",
  );
  assert.match(o, /<Class>Calendar<\/Class>/);
  assert.match(o, /<Conflict>1<\/Conflict>/);
  assert.match(o, /<BodyPreference[^>]*>[\s\S]*<Type[^>]*>1</);
});

test("device-wins is stated, on the pull and the push alike", () => {
  // The account setting reaches the wire here and nowhere else, and it is
  // the only part of the conflict policy that is ours: what the server does
  // when told it loses is the server's business. Asserted on both requests
  // because the block is sticky per collection - stating it on one and not
  // the other would leave the policy depending on which request came last.
  const pull = optionBlocks(xml({ conflict: "0", withChanges: true }))[0];
  const push = optionBlocks(
    xml({ conflict: "0", withChanges: false, withCommands: EMPTY_COMMANDS }),
  )[0];
  assert.match(pull, /<Conflict>0<\/Conflict>/);
  assert.equal(push, pull);
});

test("pull and push state the same options", () => {
  // The property that matters is not what the block contains but that the
  // two agree - a difference between them is a window that flaps.
  const pull = optionBlocks(xml({ withChanges: true }))[0];
  const push = optionBlocks(
    xml({ withChanges: false, withCommands: EMPTY_COMMANDS }),
  )[0];
  assert.equal(push, pull);
});

test("a push says GetChanges 0 - silence would mean 1", () => {
  // [MS-ASCMD] 2.2.3.84: absent + non-zero SyncKey is handled as if set to
  // 1, and a client that does not want changes MUST send 0. A push that
  // says nothing therefore asks the server for a snapshot taken while our
  // own commands are still in flight.
  const push = xml({ withChanges: false, withCommands: EMPTY_COMMANDS });
  assert.match(push, /<GetChanges>0<\/GetChanges>/);

  const instance = xml({
    withChanges: false,
    withInstanceCommand: { emit: (w) => w.atag("ServerId", "1") },
  });
  assert.match(instance, /<GetChanges>0<\/GetChanges>/);

  // The pull still asks. An empty element is TRUE, per the same section.
  const pull = xml({ withChanges: true });
  assert.doesNotMatch(pull, /<GetChanges>0<\/GetChanges>/);
  assert.match(pull, /<GetChanges\s*\/>|<GetChanges><\/GetChanges>/);
});

test("GetChanges precedes Options and Commands", () => {
  // The Collection schema fixes the order: SyncKey, CollectionId,
  // Supported, DeletesAsMoves, GetChanges, WindowSize, ConversationMode,
  // Options, Commands. Out of order is a Status 4.
  const push = xml({ withChanges: false, withCommands: EMPTY_COMMANDS });
  const g = push.indexOf("<GetChanges>");
  const o = push.indexOf("<Options>");
  assert.ok(g > -1 && o > -1, "both must be present on a push");
  assert.ok(g < o, `GetChanges must precede Options: ${push}`);
});

test("a request with neither changes nor commands states nothing", () => {
  // The bootstrap. Writing no block leaves the server's options alone,
  // which is the only safe way to say nothing.
  assert.deepEqual(optionBlocks(xml({ withChanges: false })), []);
});

test("an instance command states nothing", () => {
  // Deliberate, and load-bearing twice over: an explicit <Conflict> makes
  // Exchange discard the exception with Status 7, and writing no block
  // leaves the collection's options as the request before it set them.
  const text = xml({
    withChanges: false,
    withInstanceCommand: { emit: (w) => w.atag("ServerId", "1") },
  });
  assert.deepEqual(optionBlocks(text), []);
});

test("contacts get no FilterType, and still get the rest", () => {
  // Only Calendar has a time axis. The other elements are unaffected.
  const [o] = optionBlocks(xml({ className: "Contacts", withChanges: true }));
  assert.doesNotMatch(o, /FilterType/);
  assert.match(o, /<Class>Contacts<\/Class>/);
  assert.match(o, /<BodyPreference[^>]*>[\s\S]*<Type[^>]*>1</);
});

test("2.5 pulls a window and pushes no options at all", () => {
  const pull = optionBlocks(xml({ asVersion: "2.5", withChanges: true }));
  assert.equal(pull.length, 1);
  assert.match(pull[0], /<FilterType>7<\/FilterType>/);
  assert.doesNotMatch(pull[0], /Class|Conflict|BodyPreference/);

  // 2.5 never stated options on a push, and nothing there can retire the
  // one the pull set. Left alone on a version we cannot test.
  const push = optionBlocks(
    xml({
      asVersion: "2.5",
      withChanges: false,
      withCommands: EMPTY_COMMANDS,
    }),
  );
  assert.deepEqual(push, []);
});

/* ── Where a calendar's FilterType comes from ───────────────────────── */

const { calendarFilterType } =
  await import("../../src/modules/eas/calendar-sync.mjs");

test("a calendar's FilterType is the account's window, not the kind's", () => {
  // `calendarItemKind.filterType` is "0" - unfiltered - and is the answer
  // for tasks and contacts, which have no window to set. Reading it for a
  // calendar states an Options block that widens the window to everything,
  // and because the block is sticky that loss outlives the request that
  // caused it. Any request that states options for a calendar has to ask
  // here.
  assert.equal(calendarFilterType({ custom: { synclimit: "4" } }), "4");
  assert.equal(
    calendarFilterType({ custom: {} }),
    "7",
    "six months by default",
  );
  assert.equal(calendarFilterType(undefined), "7");
});
