/**
 * Unit tests for the two concurrency rules the OAuth token cache lives by.
 *
 * A Microsoft refresh token is single-use: redeeming it rotates it, and the
 * old one is dead the moment the new one is issued. Everything here follows
 * from that. Reported as #352 - GAL search raising a login + MFA prompt on
 * every keystroke while sync, which runs one at a time, never prompted.
 *
 * The token endpoint is the only thing stubbed: `getGlobalClientID` already
 * falls back when `browser.storage` is absent, which the test env is.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

const {
  primeAuth,
  forgetAuth,
  getAccessToken,
  invalidateAccessToken,
  setRotationSink,
} = await import("../../src/modules/eas/oauth.mjs");

/** A token endpoint that rotates the refresh token on every redemption,
 *  like Entra does, and refuses one that has already been redeemed. */
function installTokenEndpoint({ hold = false } = {}) {
  const state = { calls: [], live: "rt-0", issued: 0 };
  // With `hold`, a request parks after it has decided its answer until the
  // test releases it - by call index, so two overlapping refreshes can be
  // landed independently. Deterministic where a sleep is not: a test that
  // races a timer still passes when the timer wins, it just stops testing
  // anything.
  const arrivalWatchers = [];
  const parked = [];
  /** Resolves once `n` requests have reached the endpoint. */
  state.arrived = (n = 1) =>
    new Promise((resolve) => {
      if (state.calls.length >= n) return resolve();
      arrivalWatchers.push({ n, resolve });
    });
  /** Lets the request with this call index answer. A refused (400) request
   *  never parks, so releasing its index is a no-op. */
  state.release = (i = 0) => parked[i]?.();
  globalThis.fetch = async (url, init) => {
    const body = new URLSearchParams(init?.body ?? "");
    const presented = body.get("refresh_token");
    const idx = state.calls.push(presented) - 1;
    const still = [];
    for (const w of arrivalWatchers.splice(0)) {
      if (state.calls.length >= w.n) w.resolve();
      else still.push(w);
    }
    arrivalWatchers.push(...still);
    if (presented !== state.live) {
      // What a redeemed token earns: the loser of a race sees exactly this.
      return {
        ok: false,
        status: 400,
        text: async () => '{"error":"invalid_grant"}',
      };
    }
    state.issued += 1;
    const next = `rt-${state.issued}`;
    state.live = next;
    if (hold) {
      await new Promise((r) => (parked[idx] = r));
    }
    return {
      ok: true,
      status: 200,
      text: async () =>
        JSON.stringify({
          access_token: `at-${state.issued}`,
          expires_in: 3600,
          refresh_token: next,
        }),
    };
  };
  return state;
}

let accountSeq = 0;
const freshAccount = () => `acct-${(accountSeq += 1)}`;

beforeEach(() => setRotationSink(null));

test("concurrent callers share one refresh instead of racing for the token", async () => {
  // The #352 shape: several GAL searches in flight at once, no cached
  // access token. Without sharing, each redeems the same refresh token and
  // all but one is answered invalid_grant - an E:AUTH, and a login prompt.
  const accountId = freshAccount();
  const endpoint = installTokenEndpoint({ hold: true });
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });

  const calls = Array.from({ length: 5 }, () => getAccessToken(accountId));
  await endpoint.arrived(); // one request is provably at the endpoint
  endpoint.release(0);
  const tokens = await Promise.all(calls);

  assert.equal(endpoint.calls.length, 1, "the endpoint was asked once");
  assert.deepEqual(
    tokens,
    Array(5).fill("at-1"),
    "and every caller got that one answer",
  );
});

test("a stale prime cannot replace a token that has since rotated", async () => {
  // Each GAL callback reads the account at its own moment and primes what
  // it saw. A late arrival must not put back the token the server retired.
  const accountId = freshAccount();
  const endpoint = installTokenEndpoint();
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });
  await getAccessToken(accountId); // rotates rt-0 -> rt-1

  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });
  invalidateAccessToken(accountId);
  const token = await getAccessToken(accountId);

  assert.equal(token, "at-2", "the refresh succeeded");
  assert.deepEqual(
    endpoint.calls,
    ["rt-0", "rt-1"],
    "the second refresh presented the rotated token, not the stale one",
  );
});

test("forgetAuth is how a caller does replace it", async () => {
  // The re-authentication path: the user signed in again, so what storage
  // holds really is newer than anything in memory.
  const accountId = freshAccount();
  const endpoint = installTokenEndpoint();
  primeAuth(accountId, { refreshToken: "rt-9", servertype: "office365" });
  forgetAuth(accountId);
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });
  await getAccessToken(accountId);

  assert.deepEqual(endpoint.calls, ["rt-0"], "the newly primed token was used");
});

test("a rotation is announced once, so it can reach storage", async () => {
  // Nothing else knows a rotation happened: the caller asked for an access
  // token and got one. Before this, only the sync hooks looked, and a
  // rotation caused by a GAL search was lost.
  const accountId = freshAccount();
  installTokenEndpoint();
  const seen = [];
  setRotationSink((id, refreshToken) => seen.push([id, refreshToken]));
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });

  await getAccessToken(accountId);

  assert.deepEqual(seen, [[accountId, "rt-1"]]);
});

test("a sink that throws does not fail the request it rode in on", async () => {
  const accountId = freshAccount();
  installTokenEndpoint();
  setRotationSink(() => {
    throw new Error("storage is having a day");
  });
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });

  assert.equal(await getAccessToken(accountId), "at-1");
});

test("a cached access token is served without touching the endpoint", async () => {
  const accountId = freshAccount();
  const endpoint = installTokenEndpoint();
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });
  await getAccessToken(accountId);
  await getAccessToken(accountId);

  assert.equal(endpoint.calls.length, 1, "the second call used the cache");
});

test("a failed refresh is not remembered as an in-flight one", async () => {
  // The in-flight entry has to be cleared on the failure path too, or one
  // rejected refresh would be handed to every later caller for good.
  const accountId = freshAccount();
  const endpoint = installTokenEndpoint();
  primeAuth(accountId, { refreshToken: "wrong", servertype: "office365" });

  await assert.rejects(() => getAccessToken(accountId));
  await assert.rejects(() => getAccessToken(accountId));
  assert.equal(endpoint.calls.length, 2, "the second call really tried again");
});

test("an unprimed account is an auth error, not a silent refresh", async () => {
  installTokenEndpoint();
  await assert.rejects(
    () => getAccessToken(freshAccount()),
    /not primed/,
  );
});

test("a refresh landing after a re-authentication does not resurrect the old token", async () => {
  // The likeliest moment for this: the user is at a sign-in prompt because
  // a refresh failed, so a refresh is in the air exactly when they sign in.
  // Whatever that refresh returns belongs to a lineage the server has moved
  // on from, and must not overwrite what the sign-in just stored.
  const accountId = freshAccount();
  const endpoint = installTokenEndpoint({ hold: true });
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });
  const inFlight = getAccessToken(accountId);
  await endpoint.arrived(); // the refresh is now provably in the air

  // The interactive sign-in: forget, then seed what it issued.
  forgetAuth(accountId);
  endpoint.live = "rt-fresh";
  primeAuth(accountId, { refreshToken: "rt-fresh", servertype: "office365" });

  endpoint.release(0);
  await inFlight;
  invalidateAccessToken(accountId);
  const after = getAccessToken(accountId);
  await endpoint.arrived(2);
  endpoint.release(1);
  await after;

  assert.equal(
    endpoint.calls.at(-1),
    "rt-fresh",
    "the next refresh used the signed-in token, not the one that landed late",
  );
});

test("a sink write is skipped for a lineage that has been replaced", async () => {
  // Same race, one layer out: storing the late token would put storage on
  // the dead lineage, and the next start would prime from it.
  const accountId = freshAccount();
  const endpoint = installTokenEndpoint({ hold: true });
  const seen = [];
  setRotationSink((id, refreshToken) => seen.push(refreshToken));
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });
  const inFlight = getAccessToken(accountId);
  await endpoint.arrived();

  forgetAuth(accountId);
  endpoint.live = "rt-fresh";
  primeAuth(accountId, { refreshToken: "rt-fresh", servertype: "office365" });

  endpoint.release(0);
  await inFlight;

  assert.deepEqual(seen, [], "nothing was written for the abandoned lineage");
});

test("a refresh finishing late does not evict its successor from the slot", async () => {
  // Refresh A is in the air when the user re-authenticates and refresh B
  // starts on the new lineage. A's cleanup must vacate only its own slot -
  // deleting blindly would evict B, and a caller arriving then would start
  // refresh C with the token B is currently redeeming: the
  // parallel-redemption race, reintroduced by the cleanup of its fix.
  const accountId = freshAccount();
  const endpoint = installTokenEndpoint({ hold: true });
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });
  const a = getAccessToken(accountId);
  await endpoint.arrived(1);

  forgetAuth(accountId);
  endpoint.live = "rt-fresh";
  primeAuth(accountId, { refreshToken: "rt-fresh", servertype: "office365" });
  const b = getAccessToken(accountId);
  await endpoint.arrived(2);

  endpoint.release(0); // A lands - with B still in the air
  await a;
  const c = getAccessToken(accountId); // must join B, not start its own
  endpoint.release(1);
  const [tokenB, tokenC] = await Promise.all([b, c]);

  assert.equal(endpoint.calls.length, 2, "no third redemption was started");
  assert.equal(tokenB, tokenC, "the late caller joined the refresh in flight");
});

test("the 401 retry joins the refresh already in flight", async () => {
  // network.mjs answers a 401/403 by invalidating the access token and
  // asking again. If a refresh is already in the air, that retry must join
  // it - the refresh is fetching a new token, which is what the retry
  // wants - not start a second redemption of the same refresh token.
  const accountId = freshAccount();
  const endpoint = installTokenEndpoint({ hold: true });
  primeAuth(accountId, { refreshToken: "rt-0", servertype: "office365" });
  const first = getAccessToken(accountId);
  await endpoint.arrived(1);

  invalidateAccessToken(accountId); // the 401 handler's move
  const retry = getAccessToken(accountId);

  endpoint.release(0);
  const [a, b] = await Promise.all([first, retry]);
  assert.equal(endpoint.calls.length, 1, "one redemption served both");
  assert.equal(a, b, "and both callers hold the same token");
});
