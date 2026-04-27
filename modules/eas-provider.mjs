/**
 * EAS provider. Implements the TbSync provider contract for Exchange
 * ActiveSync servers using basic-auth + WBXML (no OAuth for now).
 *
 * Host owns all persistent state. Account `custom` carries:
 *   - server       - full EAS endpoint URL
 *   - user         - login, often full email
 *   - password     - basic-auth password (plaintext; future: move to secure storage)
 *   - deviceId     - stable per-account EAS device identifier
 *   - devicetype   - EAS device type (static "TbSyncEAS")
 *   - asversion    - negotiated AS version ("2.5" | "14.0" | "14.1" | "16.1")
 *   - policykey    - current provision key ("0" before first Provision)
 *   - foldersynckey - FolderSync key ("0" before first FolderSync)
 * Folder `custom` carries:
 *   - serverID        - EAS folder serverID (stable across syncs)
 *   - synckey         - per-folder Sync key ("0" before first Sync)
 *   - class           - EAS Class (e.g. "Contacts", "Calendar", "Tasks")
 *   - contactMap      - itemId → serverID (contacts only)
 *   - displayNameRaw  - server-supplied folder name; the visible
 *                       `displayName` is recomputed from this on every
 *                       push (with optional "Trash | " prefix)
 * The host owns `folder.changelog` (top-level field, not in custom);
 * the provider reads it from `getAccount` and clears entries via
 * `changelogRemove` / pre-tags writes via `changelogMarkServerWrite`.
 *
 * Contact sync is fully implemented (Stage 6, separate module). Calendar
 * and task folders are discovered and displayable, but `onSyncFolder`
 * returns `ok()` without touching server state.
 */

import {
  ERR, withCode, ok,
  TbSyncProviderImplementation,
} from "../vendor/tbsync/provider.mjs";
import * as addressBook from "./address-book.mjs";
import { DEBUG_STATUS_DELAY_MS } from "./debug.mjs";
import { primeAuth, isOAuthAccount } from "./eas/oauth.mjs";
import { negotiateAsVersion } from "./eas/connect.mjs";
import { acquirePolicyKey, NO_POLICY_FOR_DEVICE } from "./eas/provision.mjs";
import { runFolderSync } from "./eas/folder-sync.mjs";
import { sendDeviceInformation } from "./eas/settings.mjs";
import { NET_ERR } from "./network.mjs";

/** EAS FolderSync status codes that indicate the server wants us to run
 *  Provision (in-band equivalent of the HTTP-449 path). */
const PROVISION_REQUIRED_STATUSES = new Set(["141", "142", "143", "144"]);

/** Re-run OPTIONS once a day so we pick up server-side changes to the
 *  advertised version / command list (legacy used the same window -
 *  EAS-4-TbSync sync.js:87, 86_400_000 ms). */
const OPTIONS_REPROBE_MS = 24 * 60 * 60 * 1000;

// ── Config-popup allow-list values ───────────────────────────────────────
//
// Bounded enums for config-popup fields. The UI renders the same lists,
// but the server-side validation re-checks so a tampered runtime message
// can't write an arbitrary value into account.custom.

const ALLOWED_AS_VERSION_SELECTIONS = ["auto", "2.5", "14.0", "14.1", "16.1"];
/** ASCII char code for the multi-line address-field separator, sent
 *  through `String.fromCharCode` at sync time to split/join address
 *  lines. `"10"` = newline, `"44"` = comma. */
const ALLOWED_NAME_SEPARATORS = ["10", "44"];

/** EAS FilterType codes for calendar windowing, sent on the wire as
 *  `<FilterType>…</FilterType>` in the Sync request. */
const ALLOWED_CALENDAR_SYNC_LIMITS = ["0", "4", "5", "6", "7"];

/** Setup-type → fixed EAS host. Only the OAuth setup types appear here. */
const HOST_BY_SERVERTYPE = {
  "office365":   "outlook.office365.com",
  "personal-ms": "eas.outlook.com",
};

// ── EAS folder Type → TbSync folder type ─────────────────────────────────
//
// EAS FolderHierarchy Type values (MS-ASCMD §Type) map to TbSync folder
// types consumed by the manager's type-icon renderer.
const EAS_TYPE_TO_TBSYNC = {
  // 1: User-created email folder (ignored)
  // 2: Default inbox (ignored)
  // 3: Default drafts (ignored)
  // 4: Default deleted items (ignored)
  // 5: Default sent items (ignored)
  // 6: Default outbox (ignored)
  7: "tasks",      // Default Tasks
  8: "calendars",  // Default Calendar
  9: "contacts",   // Default Contacts
  // 10: Default Notes (ignored)
  // 11: Default Journal (ignored)
  // 12: User-created email (ignored)
  13: "calendars", // User-created Calendar
  14: "contacts",  // User-created Contacts
  15: "tasks",     // User-created Tasks
};

export function easTypeToFolderType(type) {
  return EAS_TYPE_TO_TBSYNC[Number(type)] ?? null;
}

/** Class string sent in Sync Collection. */
export function folderTypeToEasClass(folderType) {
  switch (folderType) {
    case "contacts": return "Contacts";
    case "calendars": return "Calendar";
    case "tasks": return "Tasks";
    default: return null;
  }
}

export class EasProvider extends TbSyncProviderImplementation {
  constructor() {
    super({
      name: "Exchange ActiveSync",
      shortName: "eas",
      setupPath: "dialogs/setup/setup.html",
      setupWidth: 560,
      setupHeight: 620,
      configPath: "dialogs/config/config.html",
      configWidth: 560,
      configHeight: 620,
      capabilities: {
        folderTypes: ["contacts", "calendars", "tasks"],
        supportsReadOnly: true,
        multipleAccounts: true,
        hasSetupPopup: true,
        hasConfigPopup: true,
      },
      maintainerEmail: "john.bieling@gmx.de",
      contributorsUrl: "https://github.com/jobisoft/EAS-4-TbSync",
      logPrefix: "[eas-4-tbsync]",
    });
  }

  // ── Base-class hooks ───────────────────────────────────────────────────

  async onConnectedToHost() { return null; }
  async onCancelSync(_args) { return null; }

  // ── Account lifecycle ──────────────────────────────────────────────────

  async onAccountEnabled({ accountId }) {
    // First connect after setup (or re-enable after disable): negotiate
    // EAS version, run Provision if required, run FolderSync to discover
    // the resource list, and push it to the host so the manager can
    // render contacts/calendars/tasks rows. Idempotent on re-enable.
    await this.#connectAndDiscoverFolders(accountId);
    return null;
  }

  async onAccountDisabled({ accountId }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) return null;
    // Drop local TB books. Leave account-level credentials and deviceId
    // intact so re-enable works without re-setup. The host wipes its
    // folder rows right after this returns, so per-folder custom.*
    // doesn't need clearing here.
    await this.#deleteAccountTargets(ctx.folders);
    // Force a fresh FolderSync on re-enable.
    await this.updateAccount({
      accountId,
      patch: { custom: { foldersynckey: "0" } },
    }).catch(() => { });
    return null;
  }

  async onAccountDeleted({ accountId, purgeTargets }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) return null;
    if (purgeTargets !== false) {
      await this.#deleteAccountTargets(ctx.folders);
    }
    return null;
  }

  // ── Folder lifecycle ──────────────────────────────────────────────────

  async onFolderEnabled() { return null; }

  async onFolderDisabled({ accountId, folderId }) {
    const folder = await this.#getFolder(accountId, folderId);
    if (!folder) return null;
    if (folder.targetID) {
      await safeDeleteBook(folder.targetID);
    }
    await this.updateFolder({
      accountId, folderId,
      patch: {
        targetID: null,
        targetName: null,
        custom: { synckey: "0", contactMap: {} },
      },
    }).catch(() => { });
    return null;
  }

  // ── Sync ──────────────────────────────────────────────────────────────

  async onSyncAccount({ accountId }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) throw withCode(new Error("unknown account"), ERR.UNKNOWN_ACCOUNT);
    this.reportSyncState({ accountId, syncState: "prepare" });
    // Refresh the folder list each sync so server-side additions surface.
    // Per-folder item sync (Stage 6) will be wired into onSyncFolder.
    await this.#connectAndDiscoverFolders(accountId);
    return ok();
  }

  /** OPTIONS probe (once) → pre-emptive Provision (if user-toggled) →
   *  Settings/DeviceInformation → FolderSync (with 449 → Provision +
   *  Settings retry) → push folders → persist synckey. Called from both
   *  `onAccountEnabled` (initial connect) and `onSyncAccount` (every
   *  refresh).
   *
   *  HTTP 451 (X-MS-Location host migration) is caught at the top and
   *  triggers a one-shot retry against the new host. */
  async #connectAndDiscoverFolders(accountId, redirectsRemaining = 1) {
    try {
      await this.#doConnectAndDiscover(accountId);
    } catch (err) {
      if (err.code !== NET_ERR.HOST_REDIRECT || !err.newLocation || redirectsRemaining <= 0) {
        throw err;
      }
      await this.updateAccount({
        accountId,
        patch: { custom: { server: err.newLocation } },
      });
      await this.#connectAndDiscoverFolders(accountId, redirectsRemaining - 1);
    }
  }

  async #doConnectAndDiscover(accountId) {
    let ctx = await this.#loadContext(accountId);
    if (!ctx) throw withCode(new Error("unknown account"), ERR.UNKNOWN_ACCOUNT);

    // Prime OAuth so the network layer can refresh access tokens
    // transparently across the OPTIONS / Provision / FolderSync calls.
    if (isOAuthAccount(ctx.account.custom)) this.#primeAuth(ctx);

    // 1) OPTIONS probe - on first connect, or once a day thereafter so
    // we pick up server-side changes to the advertised version / command
    // list. Even when the user has forced a version via asversionselected,
    // we still run the probe so the config popup can show the
    // server-advertised list and the negotiated default.
    const lastOptionsUpdate = Number(ctx.account.custom?.lastEasOptionsUpdate ?? 0);
    const needsOptionsProbe = !ctx.account.custom?.asversion ||
      (Date.now() - lastOptionsUpdate) > OPTIONS_REPROBE_MS;
    if (needsOptionsProbe) {
      const negotiated = await negotiateAsVersion({ account: ctx.account });
      await this.updateAccount({
        accountId,
        patch: { custom: {
          asversion: negotiated.asVersion,
          allowedEasVersions: negotiated.allowedAsVersions,
          allowedEasCommands: negotiated.allowedCommands,
          lastEasOptionsUpdate: Date.now(),
        }},
      });
      ctx = await this.#loadContext(accountId);
    }

    // The user can override negotiation via the config popup. "auto"
    // (the default) uses the negotiated value cached in account.custom.
    const selected = ctx.account.custom?.asversionselected || "auto";
    const asVersion = (selected === "auto") ? ctx.account.custom.asversion : selected;

    // 2) Pre-emptive Provision (legacy "Kerio" semantics). When the
    // user has flipped the toggle on - or a previous 449 stuck it on -
    // and we have no policy key cached, run Provision before any other
    // command. Servers that don't return 449 (e.g. Kerio) need this.
    if (ctx.account.custom?.provision === true && (ctx.account.custom?.policykey ?? "0") === "0") {
      const result = await acquirePolicyKey({ account: ctx.account, asVersion });
      if (result === NO_POLICY_FOR_DEVICE) {
        // Server demanded Provision but has no policy to apply.
        // Disable the flag and abort; user can retry, the server may
        // by then have a policy or have stopped demanding one.
        await this.updateAccount({
          accountId,
          patch: { custom: { provision: false, policykey: "0" } },
        });
        throw withCode(
          new Error("Server has no policy for this device"),
          ERR.UNKNOWN_COMMAND,
        );
      }
      await this.updateAccount({
        accountId,
        patch: { custom: { policykey: result, provision: true } },
      });
      ctx = await this.#loadContext(accountId);
    }

    // 3) Settings/DeviceInformation. Skip on AS 2.5 (the command
    // doesn't exist there) and on servers that didn't advertise it in
    // the OPTIONS probe.
    await this.#maybeSendDeviceInformation(ctx.account, asVersion);

    // 4) FolderSync, with provision/sync-key recovery loops.
    const priorFolderSyncKey = ctx.account.custom?.foldersynckey ?? "0";
    const { syncResult, ctx: ctxAfterSync } = await this.#runFolderSyncWithRecovery(accountId, ctx, asVersion);
    ctx = ctxAfterSync;

    // 5) Apply Add/Update/Delete to the host's folder list.
    //    - Initial sync (priorFolderSyncKey === "0"): server emits every
    //      folder as an Add; push the fresh list.
    //    - Incremental sync (priorFolderSyncKey !== "0"): merge the
    //      delta into the existing folder list and push the result.
    //      Skipping the push on no-op deltas avoids an unnecessary
    //      storage write + broadcast.
    if (priorFolderSyncKey === "0") {
      const initial = syncResult.adds
        .map(a => folderDescriptorFromAdd(a))
        .filter(Boolean);
      await this.pushFolderList({ accountId, folders: finalizeFolderListForPush(initial) });
    } else if (syncResult.adds.length || syncResult.updates.length || syncResult.deletes.length) {
      const merged = mergeFolderDeltas(ctx.folders, syncResult);
      await this.pushFolderList({ accountId, folders: merged });
    }

    // 6) Persist the new FolderSync continuation key.
    await this.updateAccount({
      accountId,
      patch: { custom: { foldersynckey: syncResult.synckey } },
    });
  }

  /** FolderSync with recovery for HTTP 449 / in-band Sync.Status
   *  141-144 (server demands Provision) and Status 9 (invalid sync key,
   *  reset to "0" and treat as initial). One retry per failure mode;
   *  if the same recovery is needed twice in a row we bail rather than
   *  loop. */
  async #runFolderSyncWithRecovery(accountId, ctx, asVersion) {
    let provisioned = false;
    let resetSyncKey = false;
    while (true) {
      try {
        const syncResult = await runFolderSync({ account: ctx.account, asVersion });
        return { syncResult, ctx };
      } catch (err) {
        const provisionRequired =
          err.code === NET_ERR.PROVISION_REQUIRED ||
          PROVISION_REQUIRED_STATUSES.has(err.folderSyncStatus);
        if (provisionRequired && !provisioned) {
          provisioned = true;
          ctx = await this.#provisionAndPersist(accountId, ctx, asVersion);
          continue;
        }
        if (err.folderSyncStatus === "9" && !resetSyncKey) {
          resetSyncKey = true;
          await this.updateAccount({
            accountId,
            patch: { custom: { foldersynckey: "0" } },
          });
          ctx = await this.#loadContext(accountId);
          continue;
        }
        throw err;
      }
    }
  }

  async #provisionAndPersist(accountId, ctx, asVersion) {
    const result = await acquirePolicyKey({ account: ctx.account, asVersion });
    if (result === NO_POLICY_FOR_DEVICE) {
      await this.updateAccount({
        accountId,
        patch: { custom: { provision: false, policykey: "0" } },
      });
      throw withCode(
        new Error("Server has no policy for this device"),
        ERR.UNKNOWN_COMMAND,
      );
    }
    await this.updateAccount({
      accountId,
      patch: { custom: { policykey: result, provision: true } },
    });
    const next = await this.#loadContext(accountId);
    // Re-send DeviceInformation now that the policy key has changed -
    // some servers tie device registration to the active policy.
    await this.#maybeSendDeviceInformation(next.account, asVersion);
    return next;
  }

  async #maybeSendDeviceInformation(account, asVersion) {
    if (asVersion === "2.5") return;
    const allowed = account.custom?.allowedEasCommands ?? [];
    if (!allowed.includes("Settings")) return;
    await sendDeviceInformation({ account, asVersion });
  }

  #primeAuth(ctx) {
    const c = ctx.account.custom ?? {};
    primeAuth(ctx.account.accountId, {
      refreshToken: c.refreshToken,
      servertype: c.servertype,
    });
  }

  async onGetSortedFolders({ accountId }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) return [];
    const sorted = ctx.folders
      .slice()
      .sort((a, b) => (a.orderIndex ?? 0) - (b.orderIndex ?? 0));
    return finalizeFolderListForPush(sorted);
  }

  async onSyncFolder({ accountId, folderId }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) throw withCode(new Error("unknown account"), ERR.UNKNOWN_ACCOUNT);
    const folder = ctx.folders.find(f => f.folderId === folderId);
    if (!folder) throw withCode(new Error("unknown folder"), ERR.UNKNOWN_FOLDER);

    // Per the user's scope: calendar and task folders are discoverable but
    // actual item sync is stubbed until the Calendar/Tasks codecs are
    // ported. Return ok() so the sync coordinator records a success
    // transition; no local state is changed.
    if (folder.targetType !== "contacts") {
      this.reportSyncState({ accountId, folderId, syncState: "sync" });
      await new Promise(r => setTimeout(r, DEBUG_STATUS_DELAY_MS));
      return ok();
    }

    // Stage 6: real contact sync. Until the contact-sync module lands, the
    // best we can do is return ok() with no work. The sync-coordinator
    // will mark the folder status "success" which is acceptable for a
    // skeleton build.
    return ok();
  }

  // ── Setup / config popup backings ─────────────────────────────────────

  /**
   * Create a new account from setup.html. Branches on `servertype`:
   *
   *   "office365" | "personal-ms": OAuth. Host is fixed by setup type
   *      (HOST_BY_SERVERTYPE), server URL is derived from host, refresh
   *      token comes from the consent popup.
   *   "custom": basic auth. User enters the full server URL + username +
   *      password. No host stored.
   *
   * Generates a stable device id and lets the host register a brand-new
   * account row. Initial folder list is empty; FolderSync populates it
   * on first sync.
   */
  async createAccountFromSetup(args) {
    const servertype = args.servertype;
    const trimmedLabel = String(args.label ?? "").trim() || "Exchange account";

    if (servertype === "office365" || servertype === "personal-ms") {
      const { refreshToken, authenticatedUserEmail } = args;
      if (!refreshToken) throw new Error("OAuth refresh token is required");
      const host = HOST_BY_SERVERTYPE[servertype];
      const server = `https://${host}/Microsoft-Server-ActiveSync`;
      const user = authenticatedUserEmail || args.loginHint || "";
      return {
        accountName: trimmedLabel,
        initialFolders: [],
        custom: {
          servertype,
          host,
          server,
          user,
          // Plain-old basic-auth fields stay empty for OAuth accounts.
          password: "",
          // OAuth-only fields:
          refreshToken,
          authenticatedUserEmail: authenticatedUserEmail ?? null,
          // Common EAS state:
          deviceId: generateDeviceId(),
          devicetype: "TbSyncEAS",
          asversion: "",
          policykey: "0",
          foldersynckey: "0",
          // Legacy semantic: off by default. Self-corrects to true on
          // the first 449. The user-visible toggle in the config popup
          // is a pre-emptive override for servers that need provisioning
          // but don't return 449 (e.g. Kerio).
          provision: false,
        },
      };
    }

    if (servertype !== "custom") {
      throw new Error(`Unknown servertype '${servertype}'`);
    }

    const trimmedServer = String(args.server ?? "").trim();
    const trimmedUser = String(args.user ?? "").trim();
    if (!trimmedServer) throw new Error("Server URL is required");
    if (!trimmedUser) throw new Error("Username is required");
    if (!args.password) throw new Error("Password is required");

    return {
      accountName: trimmedLabel,
      initialFolders: [],
      custom: {
        servertype: "custom",
        server: trimmedServer,
        user: trimmedUser,
        password: args.password,
        deviceId: generateDeviceId(),
        devicetype: "TbSyncEAS",
        asversion: "",
        policykey: "0",
        foldersynckey: "0",
        provision: false,
      },
    };
  }

  /** Sanitized view for config.html. Never includes password.
   *  Returns the full set of fields the popup can render, with sensible
   *  defaults so existing accounts don't show empty controls for newly-
   *  introduced settings. */
  async getAccountForConfig(accountId) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) throw withCode(new Error("Unknown account"), ERR.UNKNOWN_ACCOUNT);
    const c = ctx.account.custom ?? {};
    return {
      accountId,
      accountName: ctx.account.accountName,
      // Connection (basic auth - empty for OAuth, popup hides them anyway).
      server: c.server ?? "",
      user: c.user ?? "",
      // Account-type discriminator + identity (read-only display).
      servertype: c.servertype ?? "custom",
      authenticatedUserEmail: c.authenticatedUserEmail ?? null,
      host: c.host ?? "",
      // Protocol section.
      deviceId: c.deviceId ?? "",
      asVersion: c.asversion ?? "",
      allowedAsVersions: Array.isArray(c.allowedEasVersions) ? c.allowedEasVersions : [],
      asVersionSelected: c.asversionselected ?? "auto",
      // Legacy default is off. The 449 self-correction path will flip
      // it on automatically when the server demands provisioning.
      provision: !!c.provision,
      // Contacts section.
      contactsDisplayOverride: !!c.displayoverride,
      contactsNameSeparator: c.seperator || "10",
      // Calendar section.
      calendarSyncLimit: c.synclimit || "7",
    };
  }

  /** Write allow-listed fields from config.html via UPDATE_ACCOUNT. Any
   *  key not on this list is silently dropped. Validates the bounded
   *  enum fields so a tampered patch can't smuggle in a bogus value. */
  async saveAccountFromConfig({ accountId, patch }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) throw withCode(new Error("Unknown account"), ERR.UNKNOWN_ACCOUNT);
    const topLevelPatch = {};
    const customPatch = {};

    if ("accountName" in patch) {
      const trimmed = String(patch.accountName ?? "").trim();
      if (!trimmed) throw withCode(new Error("Account name is required"), ERR.UNKNOWN_ACCOUNT);
      topLevelPatch.accountName = trimmed;
    }
    for (const key of ["server", "user"]) {
      if (key in patch) customPatch[key] = String(patch[key] ?? "").trim();
    }
    if ("password" in patch && patch.password) {
      customPatch.password = patch.password;
    }

    if ("asVersionSelected" in patch) {
      const v = String(patch.asVersionSelected ?? "");
      if (!ALLOWED_AS_VERSION_SELECTIONS.includes(v)) {
        throw withCode(new Error("Invalid ActiveSync version selection"), ERR.UNKNOWN_COMMAND);
      }
      customPatch.asversionselected = v;
    }
    if ("provision" in patch) {
      customPatch.provision = !!patch.provision;
    }
    if ("contactsDisplayOverride" in patch) {
      customPatch.displayoverride = !!patch.contactsDisplayOverride;
    }
    if ("contactsNameSeparator" in patch) {
      const v = String(patch.contactsNameSeparator ?? "");
      if (!ALLOWED_NAME_SEPARATORS.includes(v)) {
        throw withCode(new Error("Invalid name-separator selection"), ERR.UNKNOWN_COMMAND);
      }
      customPatch.seperator = v;
    }
    if ("calendarSyncLimit" in patch) {
      const v = String(patch.calendarSyncLimit ?? "");
      if (!ALLOWED_CALENDAR_SYNC_LIMITS.includes(v)) {
        throw withCode(new Error("Invalid calendar sync limit"), ERR.UNKNOWN_COMMAND);
      }
      customPatch.synclimit = v;
    }

    const outgoing = { ...topLevelPatch };
    if (Object.keys(customPatch).length) outgoing.custom = customPatch;
    if (Object.keys(outgoing).length) {
      await this.updateAccount({ accountId, patch: outgoing });
    }
    return null;
  }

  // ── Internals ─────────────────────────────────────────────────────────

  async #loadContext(accountId) {
    const rv = await this.getAccount(accountId);
    if (!rv?.account) return null;
    return {
      account: rv.account,
      folders: rv.folders ?? [],
    };
  }

  async #getFolder(accountId, folderId) {
    const ctx = await this.#loadContext(accountId);
    return ctx?.folders.find(f => f.folderId === folderId) ?? null;
  }

  async #deleteAccountTargets(folderList) {
    for (const folder of folderList) {
      if (folder.targetID) {
        await safeDeleteBook(folder.targetID);
      }
    }
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────

async function safeDeleteBook(targetID) {
  try {
    await addressBook.deleteBook(targetID);
  } catch (err) {
    console.warn(`[eas-4-tbsync] delete book ${targetID} failed:`, err?.message ?? err);
  }
}

/** EAS requires a stable device identifier. 32 hex chars is the de-facto
 *  convention (some servers truncate to 32; we generate exactly that). */
function generateDeviceId() {
  const bytes = new Uint8Array(16);
  crypto.getRandomValues(bytes);
  return Array.from(bytes, b => b.toString(16).padStart(2, "0")).join("");
}

/** Map a FolderSync `<Add>` entry into the host folder-descriptor shape.
 *  Returns null for folder types we don't surface (mail, notes, journal).
 *  Type 4 (Deleted Items) returns a hidden marker descriptor, kept around
 *  so the trash-prefix logic can resolve children's `parentID`. */
function folderDescriptorFromAdd(add) {
  const isTrash = Number(add.type) === 4;
  const targetType = isTrash ? null : easTypeToFolderType(add.type);
  if (!isTrash && !targetType) return null;
  return {
    folderId: `f-${add.serverID}`,
    targetType,
    displayName: add.displayName,
    selected: false,
    custom: {
      serverID: add.serverID,
      parentID: add.parentID,
      type: add.type,
      class: folderTypeToEasClass(targetType),
      synckey: "0",
      contactMap: {},
      displayNameRaw: add.displayName,
    },
  };
}

/** Build the set of serverIDs identifying EAS trash folders (rows whose
 *  `custom.type === "4"`) within a list of folder records. */
export function buildTrashServerIDs(folders) {
  const ids = new Set();
  for (const f of folders) {
    if (f?.custom?.type === "4") {
      const sid = f.custom?.serverID;
      if (sid) ids.add(sid);
    }
  }
  return ids;
}

/** Recompute `hidden`, `targetType`, `displayName`, and
 *  `custom.displayNameRaw` from the current trash state. Idempotent:
 *  re-running on an already-processed folder yields the same output. */
export function applyTrashVisibility(folder, trashServerIDs) {
  const c = folder.custom ?? {};
  const isTrash = c.type === "4";
  const raw = c.displayNameRaw ?? folder.displayName ?? "";
  const inTrash = trashServerIDs.has(c.parentID);
  const displayName = inTrash
    ? `${browser.i18n.getMessage("folder.trashPrefix")} | ${raw}`
    : raw;
  return {
    ...folder,
    displayName,
    hidden: isTrash,
    targetType: isTrash ? null : folder.targetType,
    custom: { ...c, displayNameRaw: raw },
  };
}

/** Apply an incremental FolderSync delta to the existing host folder
 *  list and return the new full list to push. The host's
 *  pushFolderList is a full replace, so we have to send every folder
 *  that should remain - including the ones the delta didn't mention.
 *
 *  - Add: append a new descriptor (Type filtered as in the initial path).
 *    Legacy treats an Add for an already-known serverID as an Update.
 *  - Update: merge `displayName` / `parentID` / `type` into the
 *    existing entry's custom blob; folderType is recomputed from Type
 *    in case the server reclassified.
 *  - Delete: drop the entry entirely. */
function mergeFolderDeltas(existingFolders, delta) {
  // serverID → existing folder record. Folders that lack a serverID are
  // pre-Stage-2 ghosts; drop them silently.
  const byServerID = new Map();
  for (const f of existingFolders) {
    const sid = f.custom?.serverID;
    if (sid) byServerID.set(sid, f);
  }

  for (const upd of delta.updates) {
    const existing = byServerID.get(upd.serverID);
    if (!existing) continue;
    const targetType = easTypeToFolderType(upd.type) ?? existing.targetType;
    const rawName = upd.displayName || existing.custom?.displayNameRaw || existing.displayName;
    byServerID.set(upd.serverID, {
      ...existing,
      targetType,
      displayName: rawName,
      custom: {
        ...(existing.custom ?? {}),
        parentID: upd.parentID,
        type: upd.type,
        class: folderTypeToEasClass(targetType),
        displayNameRaw: rawName,
      },
    });
  }
  for (const del of delta.deletes) {
    byServerID.delete(del.serverID);
  }
  for (const add of delta.adds) {
    const existing = byServerID.get(add.serverID);
    if (existing) {
      // Add for a serverID we already track → treat as update (legacy).
      const targetType = easTypeToFolderType(add.type) ?? existing.targetType;
      const rawName = add.displayName || existing.custom?.displayNameRaw || existing.displayName;
      byServerID.set(add.serverID, {
        ...existing,
        targetType,
        displayName: rawName,
        custom: {
          ...(existing.custom ?? {}),
          parentID: add.parentID,
          type: add.type,
          class: folderTypeToEasClass(targetType),
          displayNameRaw: rawName,
        },
      });
      continue;
    }
    const desc = folderDescriptorFromAdd(add);
    if (desc) byServerID.set(add.serverID, desc);
  }

  return finalizeFolderListForPush([...byServerID.values()]);
}

/** Run the trash-visibility pass over a folder list and emit the
 *  canonical descriptor shape that pushFolderList accepts. Used by both
 *  the runtime sync path (initial + delta) and the migration upgrade. */
export function finalizeFolderListForPush(folders) {
  const trashServerIDs = buildTrashServerIDs(folders);
  return folders
    .map(f => applyTrashVisibility(f, trashServerIDs))
    .map(f => ({
      folderId: f.folderId,
      targetType: f.targetType,
      displayName: f.displayName,
      hidden: f.hidden,
      custom: {
        serverID: f.custom?.serverID,
        parentID: f.custom?.parentID,
        type: f.custom?.type,
        class: f.custom?.class,
        synckey: f.custom?.synckey ?? "0",
        contactMap: f.custom?.contactMap ?? {},
        displayNameRaw: f.custom?.displayNameRaw,
      },
    }));
}
