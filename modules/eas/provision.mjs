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

function policyTypeFor(asVersion) {
  return asVersion === "2.5"
    ? "MS-WAP-Provisioning-XML"
    : "MS-EAS-Provisioning-WBXML";
}

function buildInitialBody(asVersion) {
  const w = createWBXML();
  w.switchpage("Provision");
  w.otag("Provision");
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
 *  `Provision.Policies.Policy.Status = 2` ("server has no policy for
 *  this device"). The caller flips `provision: false`, clears the policy
 *  key, and aborts the current connect - legacy treats this as a
 *  contradictory state (the server demanded Provision but has no policy
 *  to apply) and lets the user re-try after the flag flips. */
export const NO_POLICY_FOR_DEVICE = Symbol("NO_POLICY_FOR_DEVICE");

/** Returns the post-ACK policy key, or `NO_POLICY_FOR_DEVICE` if the
 *  server says it has no policy for this device. Mutates
 *  `account.custom.policykey` and `account.custom.provision` in-memory
 *  between the two requests so the second POST sends the temp key as
 *  `X-MS-PolicyKey` (network.mjs gates that header on
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
    body: buildInitialBody(asVersion),
    asVersion,
  });
  if (!initial.doc)
    throw withCode(new Error("Empty Provision response"), ERR.UNKNOWN_COMMAND);

  const provisionStatus = readPath(initial.doc, ["Status"]);
  if (provisionStatus !== "1") {
    throw withCode(
      new Error(`Provision rejected (Status=${provisionStatus ?? "missing"})`),
      ERR.UNKNOWN_COMMAND,
    );
  }
  const policyStatus = readPath(initial.doc, ["Policies", "Policy", "Status"]);
  if (policyStatus === "2") {
    // Server has no policy for this device. Surface to caller; do not
    // attempt the ACK request.
    return NO_POLICY_FOR_DEVICE;
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

  const ackStatus = readPath(ack.doc, ["Status"]);
  if (ackStatus !== "1") {
    throw withCode(
      new Error(`Provision ACK rejected (Status=${ackStatus ?? "missing"})`),
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
  return finalKey;
}
