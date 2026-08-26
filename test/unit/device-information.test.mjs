/**
 * Unit tests for introducing the device to the server.
 *
 * The partnership this creates is durable server-side state, so the
 * announcement is a one-time act - [MS-ASPROV] puts it in the *initial*
 * Provision request "but not on subsequent requests". The only proof it
 * landed is the server's acknowledgement, so it goes out on every sync
 * until one arrives and never afterwards.
 *
 * Both statuses in the reply have to say 1. The outer one only says the
 * command was understood; the inner one says the device information was
 * taken, and that is the thing being asked about.
 *
 * Why this matters: a server that does not know the device is not
 * obliged to say so. Exchange parks it in DeviceDiscovery and answers
 * every folder with "nothing here" - see EAS #353 and TbSync #805, where
 * the accounts that broke were the newly created ones, because those are
 * the ones that land on 16.1.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
import { parseAdNode } from "./support/ad-node.mjs";
installWebextEnv();

const NETWORK = new URL("../../src/modules/network.mjs", import.meta.url);

const reply = (status, deviceStatus) => `<?xml version="1.0" encoding="utf-8"?>
  <Settings xmlns="Settings">
    <Status>${status}</Status>
    ${deviceStatus === undefined ? "" : `<DeviceInformation><Status>${deviceStatus}</Status></DeviceInformation>`}
  </Settings>`;

const calls = [];
let replies = [];
const real = await import(NETWORK.href);
mock.module(NETWORK.href, {
  namedExports: {
    ...real,
    getUserAgent: async () => "Thunderbird ActiveSync",
    getDeviceOs: async () => "Linux",
    easRequest: async ({ command, asVersion }) => {
      calls.push({ command, asVersion });
      return {
        doc: {
          documentElement: parseAdNode(
            replies[Math.min(calls.length - 1, replies.length - 1)],
          ),
        },
      };
    },
  },
});
const { NET_ERR } = real;
const { sendDeviceInformation, shouldSendDeviceInformation } = await import(
  "../../src/modules/eas/settings.mjs"
);

/** An account the server has never confirmed hearing about. */
const fresh = (extra = {}) => ({
  serverURL: "https://example.invalid",
  user: "u",
  custom: {
    deviceId: "tbsyncABCDEF",
    allowedEasCommands: ["Sync", "FolderSync", "Provision", "Settings"],
    ...extra,
  },
});

function expecting(...bodies) {
  calls.length = 0;
  replies = bodies;
}

/* ── who gets asked ─────────────────────────────────────────────────── */

test("a fresh 16.1 account owes the server an introduction", () => {
  assert.equal(shouldSendDeviceInformation(fresh(), "16.1"), true);
});

test("14.1 and 14.0 owe one too - every version that can carry it", () => {
  for (const v of ["14.1", "14.0"]) {
    assert.equal(shouldSendDeviceInformation(fresh(), v), true, v);
  }
});

test("once acknowledged, it is never sent again", () => {
  const account = fresh({ deviceInfoAcked: true });
  for (const v of ["16.1", "14.1", "14.0"]) {
    assert.equal(shouldSendDeviceInformation(account, v), false, v);
  }
});

test("an account that provisions is not asked by this route on 14.1/16.x", () => {
  // The Provision body carries the details there, and `provision: true`
  // never stands alone: it arrives either with a cleared policy key, so a
  // Provision runs now, or after one has completed. Announcing the same
  // device twice in one connection is what this prevents - and the
  // acknowledgement is deliberately not consulted, because it says
  // nothing about a route that is not being used.
  for (const v of ["16.1", "16.0", "14.1"]) {
    assert.equal(shouldSendDeviceInformation(fresh({ provision: true }), v), false, v);
    assert.equal(
      shouldSendDeviceInformation(fresh({ provision: true, deviceInfoAcked: false }), v),
      false,
      `${v} with an explicit false`,
    );
  }
});

test("on 14.0 it is asked anyway - that Provision body carries nothing", () => {
  assert.equal(shouldSendDeviceInformation(fresh({ provision: true }), "14.0"), true);
});

test("provisioning off leaves the acknowledgement in charge", () => {
  assert.equal(shouldSendDeviceInformation(fresh({ provision: false }), "16.1"), true);
  assert.equal(
    shouldSendDeviceInformation(fresh({ provision: false, deviceInfoAcked: true }), "16.1"),
    false,
  );
});

test("2.5 cannot convey device information at all", () => {
  assert.equal(shouldSendDeviceInformation(fresh(), "2.5"), false);
});

test("a server that never advertised Settings is not asked", () => {
  const account = fresh({ allowedEasCommands: ["Sync", "FolderSync"] });
  assert.equal(shouldSendDeviceInformation(account, "16.1"), false);
});

/* ── what counts as an acknowledgement ──────────────────────────────── */

test("both statuses saying 1 is an acknowledgement", async () => {
  expecting(reply("1", "1"));
  assert.equal(
    await sendDeviceInformation({ account: fresh(), asVersion: "16.1" }),
    true,
  );
  assert.deepEqual(calls, [{ command: "Settings", asVersion: "16.1" }]);
});

test("a command that parsed but declined the operation is not one", async () => {
  expecting(reply("1", "2"));
  await assert.rejects(
    () => sendDeviceInformation({ account: fresh(), asVersion: "16.1" }),
    /DeviceInformation\.Status=2/,
  );
});

test("a reply that says nothing about the device is not one either", async () => {
  expecting(reply("1", undefined));
  await assert.rejects(
    () => sendDeviceInformation({ account: fresh(), asVersion: "16.1" }),
    /DeviceInformation\.Status=missing/,
  );
});

test("provision-required keeps its own shape, for the recovery pass", async () => {
  for (const status of ["141", "142", "143", "144"]) {
    expecting(reply(status, undefined));
    const err = await sendDeviceInformation({
      account: fresh(),
      asVersion: "16.1",
    }).then(
      () => null,
      (e) => e,
    );
    assert.equal(err?.code, NET_ERR.PROVISION_REQUIRED, status);
  }
});
