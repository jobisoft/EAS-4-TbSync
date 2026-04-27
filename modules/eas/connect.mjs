/**
 * Negotiate an EAS protocol version with the server. The OPTIONS probe
 * returns the server's supported versions and commands; we pick the
 * highest version present in both the server's list and our supported
 * set.
 *
 * Preference order is 14.1 → 16.1 → 14.0 → 2.5. 14.1 leads because the
 * contact / calendar codecs that we'll port from legacy are most reliably
 * tested against 14.1; 16.1 is a fallback for servers that no longer
 * advertise 14.1; the others are last-resort.
 */

import { ERR, withCode } from "../../vendor/tbsync/provider.mjs";
import { easOptions } from "../network.mjs";

const SUPPORTED = ["14.1", "16.1", "14.0", "2.5"];

export async function negotiateAsVersion({ account }) {
  const probe = await easOptions({ account });
  const serverVersions = probe.versions ?? [];
  if (serverVersions.length === 0) {
    throw withCode(
      new Error("Server did not advertise any MS-ASProtocolVersions"),
      ERR.UNKNOWN_COMMAND,
    );
  }
  const asVersion = SUPPORTED.find(v => serverVersions.includes(v));
  if (!asVersion) {
    throw withCode(
      new Error(`No mutually-supported EAS version (server: ${serverVersions.join(",")})`),
      ERR.UNKNOWN_COMMAND,
    );
  }
  return {
    asVersion,
    allowedAsVersions: serverVersions,
    allowedCommands: probe.commands ?? [],
  };
}
