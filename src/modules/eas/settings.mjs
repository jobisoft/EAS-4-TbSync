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
import {
  childByTag,
  readChildTexts,
  readPath,
  readPathFrom,
} from "./wbxml-helpers.mjs";

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

function buildUserInformationBody() {
  const w = createWBXML();
  w.switchpage("Settings");
  w.otag("Settings");
  w.otag("UserInformation");
  w.atag("Get");
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** The address this account *is*, as learned from the server and stored
 *  by `#maybeLearnUserAddress`. Falls back to the login, which is right
 *  whenever the login happens to be an address and no worse than the
 *  old behaviour when it is not. Lives here, beside the request that
 *  produces the value, so the storage key has one reader and one
 *  writer. */
export function accountUserAddress(account) {
  return account?.custom?.userSmtpAddress || account?.custom?.user;
}

/** Pull the mailbox's own SMTP address out of a `UserInformation/Get`
 *  reply. Two shapes: 14.1 and later wrap the addresses in
 *  `Accounts/Account` and name the primary one; 12.1/14.0 list bare
 *  `SMTPAddress` elements with no primary marker, so the first is the
 *  best available answer. `UserDisplayName` accompanies the 14.1 shape.
 *  Returns null when the reply carries neither - callers treat that as
 *  "not learned", never as an error. */
export function readUserInformation(doc) {
  const root = doc?.documentElement;
  if (!root) return null;
  const get = childByTag(childByTag(root, "UserInformation"), "Get");
  if (!get) return null;

  // `Accounts` may list more than one, and an entry carrying no address
  // is no use here - take the first that names one, and fall back to the
  // pre-14.1 shape, where the addresses hang off `Get` directly.
  // `Array.from`, because a parsed document's `children` is a live DOM
  // collection: iterable, but with none of Array's methods on it.
  const accounts = childByTag(get, "Accounts");
  const scopes = Array.from(accounts?.children ?? [])
    .filter((c) => c.tagName === "Account")
    .concat(get);
  for (const scope of scopes) {
    const emails = childByTag(scope, "EmailAddresses");
    const address =
      readPathFrom(emails, ["PrimarySmtpAddress"]) ||
      readChildTexts(emails, "SMTPAddress")[0];
    if (!address) continue;
    return {
      address,
      displayName: readPathFrom(scope, ["UserDisplayName"]) || null,
    };
  }
  return null;
}

/** Ask the server which mailbox this account is. Same gate as
 *  `sendDeviceInformation` - `Settings` advertised in the OPTIONS probe
 *  and `asVersion != "2.5"`, where the command does not exist.
 *
 *  Sent as its own request rather than beside DeviceInformation: that
 *  one heals accounts a server would otherwise refuse to talk to, and a
 *  server that dislikes a combined body would take it down with this.
 *
 *  Returns `{address, displayName}`, or null when the server answered
 *  but named no address - a settled "no", which the caller records so
 *  the question is not asked again. A transient condition throws
 *  instead, so the next sync retries: transport failures propagate, and
 *  a Status demanding Provision is raised in the shape the recovery
 *  loops already know. */
export async function fetchUserInformation({ account, asVersion }) {
  const { doc } = await easRequest({
    account,
    command: "Settings",
    body: buildUserInformationBody(),
    asVersion,
  });
  if (!doc) return null;
  const status = readPath(doc, ["Status"]);
  if (PROVISION_REQUIRED_STATUSES.has(status)) {
    throw new EasHttpError(NET_ERR.PROVISION_REQUIRED, 0, {
      message: `Settings/UserInformation rejected (Status=${status}); server demands re-Provision`,
    });
  }
  if (status !== null && status !== "1") {
    // A refusal, not an answer. The caller records a null as "this server
    // names nobody" and stops asking for the lifetime of the account, so a
    // server having a bad minute would cost the mailbox address
    // permanently. Throw and let the next sync ask again.
    throw withCode(
      new Error(`Settings/UserInformation rejected (Status=${status})`),
      ERR.UNKNOWN_COMMAND,
    );
  }
  return readUserInformation(doc);
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
