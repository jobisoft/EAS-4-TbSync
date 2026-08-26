/**
 * EAS Provision command. Servers that enforce policy require a two-step
 * dance to acquire a long-lived `PolicyKey`:
 *
 *   1. Send `<Provision><Policies><Policy><PolicyType>…</PolicyType></Policy>
 *      </Policies></Provision>` with no policy key. The server replies with
 *      a temporary `PolicyKey` plus the policy `Data` it wants enforced.
 *   2. ACK with `<Policy><PolicyType/><PolicyKey>$temp</PolicyKey><Status>1
 *      </Status></Policy>` (using the temp key as the `X-MS-PolicyKey`
 *      header on the request). The server replies with the final
 *      `PolicyKey` to use for all subsequent commands.
 *
 * We accept whatever policy the server sends without inspecting `Data` -
 * the legacy add-on did the same. If the server only stamps a single
 * round (some servers), we still complete after the second response.
 *
 * `PolicyType` differs by AS version: legacy `MS-WAP-Provisioning-XML`
 * for 2.5, `MS-EAS-Provisioning-WBXML` for everything else.
 *
 * Status fields are read with a path-anchored walk because both
 * `Provision.Status` and `Provision.Policies.Policy.Status` use the same
 * tag name; a flat `getElementsByTagName("Status")[0]` would be brittle.
 */

import { ERR, withCode } from "../../vendor/tbsync/provider.mjs";
import { createWBXML } from "../wbxml.mjs";
import { easRequest } from "../network.mjs";
import { readPath } from "./wbxml-helpers.mjs";
import {
  appendDeviceInformationSet,
  PROVISION_EMBEDS_DEVICE_INFO,
} from "./settings.mjs";

function policyTypeFor(asVersion) {
  return asVersion === "2.5"
    ? "MS-WAP-Provisioning-XML"
    : "MS-EAS-Provisioning-WBXML";
}

async function buildInitialBody(asVersion, account) {
  const w = createWBXML();
  w.switchpage("Provision");
  w.otag("Provision");
  if (PROVISION_EMBEDS_DEVICE_INFO.has(asVersion)) {
    await appendDeviceInformationSet(w, account);
    w.switchpage("Provision");
  }
  w.otag("Policies");
  w.otag("Policy");
  w.atag("PolicyType", policyTypeFor(asVersion));
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

function buildAckBody(asVersion, policyKey) {
  const w = createWBXML();
  w.switchpage("Provision");
  w.otag("Provision");
  w.otag("Policies");
  w.otag("Policy");
  w.atag("PolicyType", policyTypeFor(asVersion));
  w.atag("PolicyKey", policyKey);
  w.atag("Status", "1");
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** Sentinel return value from `acquirePolicyKey` when the server reports
 *  `Provision.Policies.Policy.Status = 2` - [MS-ASPROV] "there is no
 *  policy for this client".
 *
 *  A complete answer to what we asked, not a failure: a server that
 *  enforces nothing says so this way. There is no key to acquire and no
 *  ACK to send, because an ACK acknowledges a policy and none was
 *  offered - so this returns before the second request.
 *
 *  The caller records it as `provision: false` and carries on with the
 *  connection. */
export const NO_POLICY_FOR_DEVICE = Symbol("NO_POLICY_FOR_DEVICE");

/** Runs the Provision exchange and reports both of its outcomes:
 *
 *    `policy`         the post-ACK policy key, or `NO_POLICY_FOR_DEVICE`
 *    `deviceInfoAcked` whether the server confirmed the device details
 *
 *  The second is not a side errand. On 14.1/16.x the initial request
 *  carries `settings:DeviceInformation` because [MS-ASPROV] §2.2.2.53
 *  requires it there, and the reply answers it in the same document. A
 *  caller that ignored that would have to ask again through the Settings
 *  command for something it has already been told.
 *
 *  Read before the policy is judged, because the two are independent: a
 *  server with no policy to apply still accepts the device details, and
 *  that is the case this was measured on (#353).
 *
 *  Mutates `account.custom.policykey` and `account.custom.provision`
 *  in-memory between the two requests so the second POST sends the temp
 *  key as `X-MS-PolicyKey` (network.mjs gates that header on
 *  `provision === true`). The caller persists the returned final key
 *  (and the `provision: true` flip) onto the host row. */
export async function acquirePolicyKey({ account, asVersion }) {
  // Bootstrap state for iter 0: provision flag on, policykey "0". The
  // network layer will then send `X-MS-PolicyKey: 0` on the first request,
  // matching legacy behaviour.
  account.custom = {
    ...(account.custom ?? {}),
    provision: true,
    policykey: "0",
  };

  const initial = await easRequest({
    account,
    command: "Provision",
    body: await buildInitialBody(asVersion, account),
    asVersion,
  });
  if (!initial.doc)
    throw withCode(new Error("Empty Provision response"), ERR.UNKNOWN_COMMAND);

  // A missing top-level Status is tolerated, a present non-1 is not.
  // [MS-ASPROV] requires the element, but Tencent Exmail omits it
  // entirely and puts the verdict only in Policies/Policy/Status - which
  // is validated in full right below, including the no-policy case. v4
  // never read the top-level element, which is why it worked against
  // that server for years (#337; response captured there, fix verified
  // against the live server by the reporter).
  const provisionStatus = readPath(initial.doc, ["Status"]);
  if (provisionStatus !== null && provisionStatus !== "1") {
    throw withCode(
      new Error(`Provision rejected (Status=${provisionStatus})`),
      ERR.UNKNOWN_COMMAND,
    );
  }
  // Only ever present when we put the element in the request, so this
  // needs no version test of its own: a server does not answer for
  // something it was not asked.
  const deviceInfoAcked =
    readPath(initial.doc, ["DeviceInformation", "Status"]) === "1";

  const policyStatus = readPath(initial.doc, ["Policies", "Policy", "Status"]);
  if (policyStatus === "2") {
    // Server has no policy for this device. Surface to caller; do not
    // attempt the ACK request.
    return { policy: NO_POLICY_FOR_DEVICE, deviceInfoAcked };
  }
  if (policyStatus !== "1") {
    throw withCode(
      new Error(
        `Provision policy rejected (PolicyStatus=${policyStatus ?? "missing"})`,
      ),
      ERR.UNKNOWN_COMMAND,
    );
  }
  const tempKey = readPath(initial.doc, ["Policies", "Policy", "PolicyKey"]);
  if (!tempKey) {
    throw withCode(
      new Error("Provision response missing PolicyKey"),
      ERR.UNKNOWN_COMMAND,
    );
  }

  // Iter 1: temp key becomes the X-MS-PolicyKey header on the ACK request.
  account.custom.policykey = tempKey;

  const ack = await easRequest({
    account,
    command: "Provision",
    body: buildAckBody(asVersion, tempKey),
    asVersion,
  });
  if (!ack.doc)
    throw withCode(
      new Error("Empty Provision ACK response"),
      ERR.UNKNOWN_COMMAND,
    );

  // Same tolerance as the initial response: the ACK's PolicyKey below is
  // the part that cannot be absent.
  const ackStatus = readPath(ack.doc, ["Status"]);
  if (ackStatus !== null && ackStatus !== "1") {
    throw withCode(
      new Error(`Provision ACK rejected (Status=${ackStatus})`),
      ERR.UNKNOWN_COMMAND,
    );
  }
  const finalKey = readPath(ack.doc, ["Policies", "Policy", "PolicyKey"]);
  if (!finalKey) {
    throw withCode(
      new Error("Provision ACK response missing PolicyKey"),
      ERR.UNKNOWN_COMMAND,
    );
  }
  account.custom.policykey = finalKey;
  return { policy: finalKey, deviceInfoAcked };
}
