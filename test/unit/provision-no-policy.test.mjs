/**
 * Unit tests for what `Provision.Policies.Policy.Status = 2` means.
 *
 * [MS-ASPROV] reads it as "there is no policy for this client" - the
 * complete and correct answer of a server that enforces nothing on this
 * device. There is no key to acquire and no ACK to send, because an ACK
 * acknowledges a policy and none was offered.
 *
 * The reply below is tomaskovacik's, pasted verbatim from EAS #353. His
 * tenant answers a first-ever Provision this way, and treating it as a
 * failure aborted a connection that had nothing wrong with it - while the
 * `DeviceInformation` in the same response had already been accepted.
 *
 * No live server of ours gives this answer, which is why it is captured
 * here rather than measured: it is the server's verdict, not a state we
 * can put a server into.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
installWebextEnv();

const NETWORK = new URL("../../src/modules/network.mjs", import.meta.url);

/** EAS #353, the tenant's answer to a first Provision. */
const NO_POLICY = `<?xml version="1.0" encoding="utf-8"?>
  <Provision xmlns="Provision">
    <DeviceInformation><Status>1</Status></DeviceInformation>
    <Status>1</Status>
    <Policies><Policy>
      <PolicyType>MS-EAS-Provisioning-WBXML</PolicyType>
      <Status>2</Status>
    </Policy></Policies>
  </Provision>`;

/** A server that does enforce a policy: temp key, then the final one. */
const withPolicy = (policyKey) => `<?xml version="1.0" encoding="utf-8"?>
  <Provision xmlns="Provision">
    <Status>1</Status>
    <Policies><Policy>
      <PolicyType>MS-EAS-Provisioning-WBXML</PolicyType>
      <Status>1</Status>
      <PolicyKey>${policyKey}</PolicyKey>
    </Policy></Policies>
  </Provision>`;

/** A policy verdict we do not recognise. */
const POLICY_STATUS_3 = `<?xml version="1.0" encoding="utf-8"?>
  <Provision xmlns="Provision">
    <Status>1</Status>
    <Policies><Policy>
      <PolicyType>MS-EAS-Provisioning-WBXML</PolicyType>
      <Status>3</Status>
    </Policy></Policies>
  </Provision>`;

function fakeDoc(xml) {
  return { documentElement: parseAdNode(xml) };
}

/* Installed once, before the module under test is imported: ESM caches the
 * import, so a per-test mock would bind it to the first test's stub. Only
 * the network reach is swapped - `getUserAgent` / `getDeviceOs` live here
 * too and would go looking for browser storage, which the unit env keeps
 * deliberately absent. */
const calls = [];
let replies = [];
const real = await import(NETWORK.href);
mock.module(NETWORK.href, {
  namedExports: {
    ...real,
    getUserAgent: async () => "Thunderbird ActiveSync",
    getDeviceOs: async () => "Linux",
    easRequest: async ({ account }) => {
      calls.push({ policyKeyHeader: account?.custom?.policykey });
      return {
        doc: fakeDoc(replies[Math.min(calls.length - 1, replies.length - 1)]),
      };
    },
  },
});
const { acquirePolicyKey, NO_POLICY_FOR_DEVICE } = await import(
  "../../src/modules/eas/provision.mjs"
);

/** Arm the stub for one test, and hand back a fresh account. */
function expecting(...bodies) {
  calls.length = 0;
  replies = bodies;
  return { serverURL: "https://example.invalid", user: "u", custom: { deviceId: "tbsyncABCDEF" } };
}

test("no policy for this device is reported, not thrown", async () => {
  const account = expecting(NO_POLICY);
  const result = await acquirePolicyKey({ account, asVersion: "16.1" });
  assert.equal(result, NO_POLICY_FOR_DEVICE);
});

test("no policy means no ACK - there is nothing to acknowledge", async () => {
  const account = expecting(NO_POLICY);
  await acquirePolicyKey({ account, asVersion: "16.1" });
  assert.equal(calls.length, 1, "the ACK request must not be sent");
});

test("a real policy is acknowledged and its final key returned", async () => {
  const account = expecting(withPolicy("TEMP-1"), withPolicy("FINAL-2"));
  const result = await acquirePolicyKey({ account, asVersion: "16.1" });
  assert.equal(result, "FINAL-2");
  assert.equal(calls.length, 2);
  // The ACK carries the temp key as X-MS-PolicyKey; the bootstrap "0" is
  // what the first request sends.
  assert.deepEqual(
    calls.map((c) => c.policyKeyHeader),
    ["0", "TEMP-1"],
  );
});

test("a policy verdict we do not know still fails loudly", async () => {
  const account = expecting(POLICY_STATUS_3);
  await assert.rejects(
    () => acquirePolicyKey({ account, asVersion: "16.1" }),
    /PolicyStatus=3/,
  );
});
