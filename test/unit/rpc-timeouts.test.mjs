/**
 * Unit tests for which RPCs are allowed to take as long as they take.
 *
 * `NO_TIMEOUT_CMDS` is not transmitted or agreed between the two sides:
 * each reads its own copy to decide how long *it* waits. So a command
 * dropped from this set does not fail loudly anywhere - it starts
 * abandoning work that is still running and reporting a failure for it.
 * That is what happened to `requestSync` (#351): the host answers it only
 * once the sync it asked for has finished, the 30s default applied, and a
 * duplicate cleanup of several hundred copies removed every one of them
 * and then said it had timed out.
 *
 * These are pinned here because the file they live in is vendored, and a
 * re-sync from upstream that lost an entry would restore the bug in
 * silence.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";

const { NO_TIMEOUT_CMDS, DEFAULT_RPC_TIMEOUT_MS, HOST_CMD, PROVIDER_CMD } =
  await import("../../src/vendor/tbsync/protocol.mjs");

test("a sync is never timed out, whichever side is waiting for it", () => {
  // The host waits on these while a provider syncs.
  assert.ok(NO_TIMEOUT_CMDS.has(HOST_CMD.SYNC_ACCOUNT));
  assert.ok(NO_TIMEOUT_CMDS.has(HOST_CMD.SYNC_FOLDER));
  assert.ok(NO_TIMEOUT_CMDS.has(HOST_CMD.MAINTAIN));
  // And the provider waits on this one, which the host answers only when
  // the sync it asked for has finished - the same wait, seen from the
  // other end.
  assert.ok(
    NO_TIMEOUT_CMDS.has(PROVIDER_CMD.REQUEST_SYNC),
    "requestSync resolves when a sync does, so it cannot carry a deadline",
  );
});

test("a window someone is reading is never timed out", () => {
  for (const cmd of [
    HOST_CMD.OPEN_SETUP_POPUP,
    HOST_CMD.OPEN_CONFIG_POPUP,
    HOST_CMD.OPEN_SERVICES_POPUP,
    HOST_CMD.REAUTHENTICATE,
  ]) {
    assert.ok(NO_TIMEOUT_CMDS.has(cmd), `${cmd} waits on a person`);
  }
});

test("everything else answers promptly or is considered dead", () => {
  // The guard is only worth having if it still applies to the ordinary
  // calls - a set that swallowed everything would never catch a peer that
  // has stopped answering at all.
  for (const cmd of [
    HOST_CMD.GET_CHANGELOG,
    HOST_CMD.GET_SORTED_FOLDERS,
    PROVIDER_CMD.UPDATE_FOLDER,
    PROVIDER_CMD.GET_ACCOUNT,
  ]) {
    assert.ok(!NO_TIMEOUT_CMDS.has(cmd), `${cmd} must stay bounded`);
  }
  assert.equal(DEFAULT_RPC_TIMEOUT_MS, 30_000);
});
