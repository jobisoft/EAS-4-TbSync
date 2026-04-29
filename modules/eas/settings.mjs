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
import { easRequest, getUserAgent } from "../network.mjs";
import { readPath } from "./wbxml-helpers.mjs";

// Matches legacy network.js:832-833 verbatim. `Model` is the device-class
// label Exchange surfaces in its mobile-device list; `FriendlyName` is the
// per-account label - legacy strips the 4-char generator prefix off the
// deviceId, we preserve the same call shape so multi-account installations
// stay distinguishable in the Exchange admin UI.
const MODEL = "Computer";

function buildBody({ deviceId, userAgent }) {
  if (!deviceId) {
    throw new Error("settings.buildBody: deviceId is required");
  }
  const w = createWBXML();
  w.switchpage("Settings");
  w.otag("Settings");
  w.otag("DeviceInformation");
  w.otag("Set");
  w.atag("Model", MODEL);
  w.atag("FriendlyName", `TbSync on Device ${deviceId.slice(4)}`);
  w.atag(
    "OS",
    typeof navigator !== "undefined" && navigator.platform
      ? navigator.platform
      : "Unknown",
  );
  w.atag("UserAgent", userAgent);
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** Send DeviceInformation/Set. Throws on a non-1 Settings.Status; the
 *  caller is expected to invoke this only when `allowedEasCommands` includes
 *  "Settings" (the OPTIONS-probed command list) and `asVersion != "2.5"`.
 *  Returns null on success. */
export async function sendDeviceInformation({ account, asVersion }) {
  const userAgent = await getUserAgent();
  const { doc } = await easRequest({
    account,
    command: "Settings",
    body: buildBody({
      deviceId: account?.custom?.deviceId,
      userAgent,
    }),
    asVersion,
  });
  if (!doc) {
    throw withCode(new Error("Empty Settings response"), ERR.UNKNOWN_COMMAND);
  }
  const status = readPath(doc, ["Status"]);
  if (status !== "1") {
    throw withCode(
      new Error(`Settings rejected (Status=${status ?? "missing"})`),
      ERR.UNKNOWN_COMMAND,
    );
  }
  return null;
}
