/**
 * EAS Settings command - specifically the DeviceInformation/Set sub-
 * operation, by which the device introduces itself to the server.
 *
 * There are two carriers. On 14.1 and above the initial Provision body
 * embeds the element, because [MS-ASPROV] §2.2.2.53 requires it there;
 * this command carries it everywhere else. An account that provisions
 * therefore never uses this route on those versions - the announcement
 * has gone, or is about to, in a request of its own.
 *
 * Where this route is the one in use, it goes out on every sync until the
 * server acknowledges it and never again. Some servers reject FolderSync
 * from a device that has not introduced itself; Exchange is quieter about
 * it and simply reports every folder as empty while the device sits in
 * DeviceDiscovery, which is why silence cannot be read as success.
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
  easCommandAdvertised,
  easCommandLikelyAvailable,
} from "./allowed-commands.mjs";
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

/** AS versions whose *initial* Provision body must embed
 *  `settings:DeviceInformation`, per [MS-ASPROV] §3.1.4.1.1. The ACK
 *  ("subsequent") request goes out without it even on these.
 *
 *  Lives here rather than with the Provision command because it is a
 *  fact about where device information travels, and both carriers need
 *  to agree on it: it is what tells `shouldSendDeviceInformation` that
 *  an account which provisions has no use for this route. */
export const PROVISION_EMBEDS_DEVICE_INFO = Object.freeze(
  new Set(["14.1", "16.0", "16.1"]),
);

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

/** Whether this account still owes the server an introduction *by this
 *  route*.
 *
 *  `2.5` has no way to convey device information at all, and a server
 *  that did not advertise `Settings` cannot be asked through it.
 *
 *  An account that provisions does not need this route on the versions
 *  whose Provision body embeds the details. `provision: true` is never
 *  written alone: it arrives either with `policykey: "0"`, so a Provision
 *  runs on this connect, or after one has completed. Either way the
 *  details have gone or are about to, and sending them again here would
 *  be the same announcement twice in one connection. On 14.0 the
 *  Provision body carries nothing, so there this route is the only one
 *  and the flag means nothing to it.
 *
 *  Otherwise it is the acknowledgement that decides, and only that: the
 *  partnership is durable server-side state - [MS-ASPROV] puts the
 *  announcement in the *initial* Provision request "but not on
 *  subsequent requests" - so the server's confirmation is the one thing
 *  that can stop the asking. An account never confirmed looks exactly
 *  like one never asked, which is what carries existing accounts over. */
export function shouldSendDeviceInformation(account, asVersion) {
  if (asVersion === "2.5") return false;
  // Normally the server has to have advertised Settings: introducing the
  // device is not worth a request against a server that says it cannot
  // take one. But an account connected without a usable OPTIONS probe has
  // no command list at all, and reading that silence as "no Settings"
  // would leave it never introducing itself - which is #353's empty-sync
  // exactly, reappearing for the accounts least able to afford it. There,
  // ask anyway: a refusal costs one request and is already absorbed.
  const advertised = account?.custom?.easOptionsUnavailable
    ? easCommandLikelyAvailable(account, "Settings")
    : easCommandAdvertised(account, "Settings");
  if (!advertised) return false;
  if (
    account?.custom?.provision === true &&
    PROVISION_EMBEDS_DEVICE_INFO.has(asVersion)
  ) {
    return false;
  }
  return account?.custom?.deviceInfoAcked !== true;
}

/** Send DeviceInformation/Set and report whether the server confirmed
 *  it. Returns true only on an acknowledgement; throws otherwise, and
 *  the caller decides what a refusal is worth (see
 *  `#maybeSendDeviceInformation`).
 *
 *  Both statuses have to be 1. The outer one says the command was
 *  understood, the inner one says the device information was taken, and
 *  only the second is the thing being asked about - a server can parse
 *  the request and still decline the operation inside it. */
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
  if (PROVISION_REQUIRED_STATUSES.has(status)) {
    // Server demands re-Provision before accepting DeviceInformation.
    // Same shape HTTP 449 throws (network.mjs), so the upstream
    // recovery loop on PROVISION_REQUIRED handles both signals.
    throw new EasHttpError(NET_ERR.PROVISION_REQUIRED, 0, {
      message: `Settings rejected (Status=${status}); server demands re-Provision`,
    });
  }
  const deviceStatus = readPath(doc, ["DeviceInformation", "Status"]);
  if (status === "1" && deviceStatus === "1") return true;
  throw withCode(
    new Error(
      `Settings rejected (Status=${status ?? "missing"}, ` +
        `DeviceInformation.Status=${deviceStatus ?? "missing"})`,
    ),
    ERR.UNKNOWN_COMMAND,
  );
}
