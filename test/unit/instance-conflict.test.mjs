/**
 * Unit tests for what an instance command does with the server's verdict.
 *
 * Every push declares `<Conflict>1</Conflict>` - the server's copy wins - so
 * a Status 7 is that policy working, not a fault. [MS-ASCMD] 2.2.3.177.17
 * gives it Item scope and resolves it with "inform the user that the change
 * they made to the item has been overwritten by a server change". Accept it,
 * say so, and stop: re-sending asks the server to reverse a verdict it is
 * entitled to give.
 *
 * A Status 5 or 16 is the opposite - a fault with no verdict attached, whose
 * documented resolutions are "retry the synchronization" and "resend the
 * request". Both cases are asserted here so the distinction cannot be lost
 * by editing one table.
 *
 * The live suite cannot cover this: a conflict arrives when the server
 * happens to produce one, which is luck rather than a test. Hence a
 * synthetic reply.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
installWebextEnv();

const NETWORK = new URL("../../src/modules/network.mjs", import.meta.url);

/** One `<Change>` response for our command, carrying `status`. */
function replyWith(status) {
  return `<?xml version="1.0" encoding="utf-8"?>
    <Sync xmlns="AirSync">
      <Collections><Collection>
        <SyncKey>2</SyncKey><CollectionId>7</CollectionId><Status>1</Status>
        <Responses><Change>
          <ServerId>SRV-1</ServerId>
          <InstanceId>20260909T130000Z</InstanceId>
          <Status>${status}</Status>
        </Change></Responses>
      </Collection></Collections>
    </Sync>`;
}

/** The two things reply parsing asks a Document for: `documentElement`,
 *  which `readPath` walks from, and `getElementsByTagName`, which finds the
 *  collection. Built over `parseAdNode`'s node shape so a reply pasted from
 *  the Event Log stays usable verbatim. */
function fakeDoc(xml) {
  const root = parseAdNode(xml);
  const all = [];
  (function walk(n) {
    all.push(n);
    for (const c of n.children ?? []) walk(c);
  })(root);
  return {
    documentElement: root,
    getElementsByTagName: (tag) => all.filter((n) => n.tagName === tag),
  };
}

/** A context with only what `sendInstanceCommand` reaches for. */
function fakeCtx(events) {
  return {
    account: { serverURL: "https://example.invalid", user: "u" },
    accountId: "1",
    folderId: "f",
    asVersion: "16.1",
    synckey: "1",
    collectionId: "7",
    itemKind: { changelogKind: "event" },
    provider: { reportEventLog: (e) => events.push(e) },
  };
}

const COMMAND = {
  kind: "Change",
  instanceId: "20260909T130000Z",
  serverID: "SRV-1",
  emit: (w) => w.atag("ServerId", "SRV-1"),
};

/* The stub is installed once, before the runner is imported, and stays.
 * ESM caches the import, so a per-test mock would bind the runner to the
 * first test's stub and leave every later one talking to a function that
 * has been reset - which reads as "no request was ever sent". */
const calls = [];
let replies = [];

// Spread the real module: mocking replaces the whole namespace, and other
// modules in this graph import `EasHttpError` from here. Only the one
// function that reaches the network is swapped.
const real = await import(NETWORK.href);
mock.module(NETWORK.href, {
  namedExports: {
    ...real,
    easRequest: async ({ body }) => {
      calls.push(body);
      // The last reply repeats, so a test says only as much as it means:
      // "a 7" is one entry, and any resend meets the same answer.
      return { doc: fakeDoc(replies[Math.min(calls.length - 1, replies.length - 1)]) };
    },
  },
});
const { sendInstanceCommand } = await import(
  "../../src/modules/eas/sync-runner.mjs"
);

/** Arm the stub for one test. */
function expecting(...bodies) {
  calls.length = 0;
  replies = bodies;
  return calls;
}

test("a Status 7 is accepted, not re-sent", async () => {
  const events = [];
  const calls = expecting(replyWith("7"));
  const r = await sendInstanceCommand(fakeCtx(events), COMMAND, "BLOB");

  // The whole point: one request. Asking again is asking the server to
  // change a verdict it is entitled to give, and the third attempt is what
  // put the two sides out of step on 16.1.
  assert.equal(calls.length, 1, "the command must not be re-sent");

  // Not a failure either. The sync did what the declared policy told it to,
  // so nothing here may sink the folder or re-stage the changelog entry -
  // the server's copy has already been applied by the caller.
  assert.equal(r.failed, false);

  // Said out loud at warning level: the spec resolves a 7 with "inform
  // the user", and a user's edit has just been undone.
  assert.ok(
    events.some((e) => e.level === "warning" && /Status 7/.test(e.message)),
    `the conflict must be reported: ${JSON.stringify(events)}`,
  );
});

test("a transient fault is still re-sent", async () => {
  const events = [];
  // Status 5 first, then success: the distinction the retry table encodes
  // is verdict versus fault, and dropping the conflict retry must not have
  // taken the fault retries with it.
  const calls = expecting(replyWith("5"), replyWith("1"));
  const r = await sendInstanceCommand(fakeCtx(events), COMMAND, "BLOB");
  assert.ok(calls.length > 1, `a Status 5 must be re-sent, sent ${calls.length}`);
  assert.equal(r.failed, false);
});
