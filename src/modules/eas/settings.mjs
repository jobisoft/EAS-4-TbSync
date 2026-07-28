/**
 * EAS Settings command - specifically the DeviceInformation/Set sub-
 * operation. Legacy sends this every account sync (gated on AS != 2.5
 * AND the server having advertised the Settings command in OPTIONS).
 * Some servers reject FolderSync from devices that haven't introduced
 * themselves; sending DeviceInformation up front keeps everyone happy.
 *
 *   <Settings>
 *     <DeviceInformation>
 *       <Set>
 *         <Model>…</Model>
 *         <FriendlyName>…</FriendlyName>
 *         <OS>…</OS>
 *         <UserAgent>…</UserAgent>
 *       </Set>
 *     </DeviceInformation>
 *   </Settings>
 */

import { ERR, withCode } from "../../vendor/tbsync/provider.mjs";
import { createWBXML } from "../wbxml.mjs";
import {
  EasHttpError,
  NET_ERR,
  easRequest,
  getDeviceOs,
  getUserAgent,
} from "../network.mjs";
import { readPath } from "./wbxml-helpers.mjs";

const PROVISION_REQUIRED_STATUSES = new Set(["141", "142", "143", "144"]);

// Matches legacy network.js:832-833 verbatim. `Model` is the device-class
// label Exchange surfaces in its mobile-device list; `FriendlyName` is the
// per-account label - legacy strips the 4-char generator prefix off the
// deviceId, we preserve the same call shape so multi-account installations
// stay distinguishable in the Exchange admin UI.
const MODEL = "Computer";

/** Append `<DeviceInformation><Set>…</Set></DeviceInformation>` under
 *  the Settings codepage. Leaves the writer's codepage state at
 *  "Settings"; caller switches back to its own codepage if needed.
 *  Used by both `buildBody` (Settings command, AS 12.x/14.0) and
 *  `provision.buildInitialBody` (initial Provision, AS 14.1/16.x). */
export async function appendDeviceInformationSet(w, account) {
  const [userAgent, deviceOs] = await Promise.all([
    getUserAgent(),
    getDeviceOs(),
  ]);
  const deviceId = account?.custom?.deviceId;
  if (!deviceId) {
    throw new Error("appendDeviceInformationSet: deviceId is required");
  }
  w.switchpage("Settings");
  w.otag("DeviceInformation");
  w.otag("Set");
  w.atag("Model", MODEL);
  w.atag("FriendlyName", `TbSync on Device ${deviceId.slice(4)}`);
  w.atag("OS", deviceOs);
  w.atag("UserAgent", userAgent);
  w.ctag();
  w.ctag();
}

async function buildBody(account) {
  const w = createWBXML();
  w.switchpage("Settings");
  w.otag("Settings");
  await appendDeviceInformationSet(w, account);
  w.ctag();
  return w.getBytes();
}

/** Send DeviceInformation/Set. Throws on a non-1 Settings.Status; the
 *  caller is expected to invoke this only when `allowedEasCommands` includes
 *  "Settings" (the OPTIONS-probed command list) and `asVersion != "2.5"`.
 *  Returns null on success. */
export async function sendDeviceInformation({ account, asVersion }) {
  const { doc } = await easRequest({
    account,
    command: "Settings",
    body: await buildBody(account),
    asVersion,
  });
  if (!doc) {
    throw withCode(new Error("Empty Settings response"), ERR.UNKNOWN_COMMAND);
  }
  const status = readPath(doc, ["Status"]);
  if (status === "1") return null;
  if (PROVISION_REQUIRED_STATUSES.has(status)) {
    // Server demands re-Provision before accepting DeviceInformation.
    // Same shape HTTP 449 throws (network.mjs), so the upstream
    // recovery loop on PROVISION_REQUIRED handles both signals.
    throw new EasHttpError(NET_ERR.PROVISION_REQUIRED, 0, {
      message: `Settings rejected (Status=${status}); server demands re-Provision`,
    });
  }
  throw withCode(
    new Error(`Settings rejected (Status=${status ?? "missing"})`),
    ERR.UNKNOWN_COMMAND,
  );
}
