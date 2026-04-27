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
import { easRequest } from "../network.mjs";
import { readPath } from "./wbxml-helpers.mjs";

const MODEL = "TbSyncEAS";
const FRIENDLY_NAME = "Thunderbird";
const USER_AGENT = "TbSync-EAS/1.0";

function buildBody() {
  const w = createWBXML();
  w.switchpage("Settings");
  w.otag("Settings");
    w.otag("DeviceInformation");
      w.otag("Set");
        w.atag("Model", MODEL);
        w.atag("FriendlyName", FRIENDLY_NAME);
        w.atag("OS", typeof navigator !== "undefined" && navigator.platform ? navigator.platform : "Unknown");
        w.atag("UserAgent", USER_AGENT);
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
  const { doc } = await easRequest({
    account,
    command: "Settings",
    body: buildBody(),
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
