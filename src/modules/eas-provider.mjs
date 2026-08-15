/**
 * EAS provider. Implements the TbSync provider contract, speaking WBXML to
 * Exchange ActiveSync servers.
 *
 * Two authentication flavours, chosen by `custom.servertype`:
 *   - "office365" / "personal-ms" - OAuth. `refreshToken` and
 *     `authenticatedUserEmail` carry the identity; `server` is derived from
 *     the flavour rather than entered.
 *   - "auto" (Autodiscover) / "custom" - basic auth with `user` + `password`.
 * `eas/oauth.mjs::isOAuthAccount` is the check used throughout.
 *
 * Host owns all persistent state; this add-on has no storage of its own
 * beyond a schema marker (see upgrades.mjs). Account `custom` is the opaque
 * blob the host round-trips for us. `createAccountFromSetup` below writes
 * the authoritative shape - rather than restate it here and let the two
 * drift, only the fields you need in order to follow the code:
 *
 *   - server, user, password / servertype, refreshToken,
 *     authenticatedUserEmail - see the two flavours above
 *   - deviceId       - stable per-account EAS device identifier
 *   - asversion      - negotiated AS version ("2.5" | "14.0" | "14.1" | "16.1")
 *   - policykey      - current provision key ("0" before first Provision)
 *   - foldersynckey  - FolderSync key ("0" before first FolderSync)
 *
 * The rest fall into three groups: the OPTIONS probe cache
 * (`allowedEasVersions`, `allowedEasCommands`, `lastEasOptionsUpdate`), the
 * config-popup options (`asversionselected`, `provision`, `conflict`,
 * `syncrecurrence`, `synclimit`, `displayoverride`, `seperator` - spelled
 * that way on disk), and GAL state (`galenabled`, `galName`).
 *
 * Credentials sit in that account row in `storage.local`, unencrypted. A
 * WebExtension has no write path to Thunderbird's password manager - the
 * `LegacyLoginManager` experiment this add-on ships is read-only
 * (`getLoginInfo`), and exists only to import what TbSync 4 stored there.
 * A known limitation, not a pending task.
 *
 * Folder `custom` carries:
 *   - serverID        - EAS folder serverID (stable across syncs)
 *   - parentID        - EAS parent folder serverID, for hierarchy
 *   - synckey         - per-folder Sync key ("0" before first Sync)
 *   - class           - EAS Class (e.g. "Contacts", "Calendar", "Tasks")
 *   - indexMap        - array of {uid, serverId} (any kind). Which local
 *                       item each server id names, and the only answer to
 *                       that - a pulled command that cannot be matched
 *                       creates a second copy of the item. Kept across a
 *                       sync key reset, cleared only where the resource's
 *                       contents go too. See `eas/server-id-index.mjs`.
 *   - displayNameRaw  - server-supplied folder name; the visible
 *                       `displayName` is recomputed from this on every
 *                       push (with optional "Trash | " prefix)
 * Pending user edits are ours, for every resource. A calendar we supply
 * hands us the edit directly; an address book we watch (see the vendored
 * the vendored `address-book.mjs`), pre-tagging our own writes so the events they
 * produce are not mistaken for the user's. Both queues live in our own
 * storage, keyed by `folder.sessionId` - the binding the host names - so
 * they survive the host being absent and vanish when it ends the binding.
 *
 * Sync entry points: `syncContactFolder` for contacts (vCard codec, TB
 * address book), `syncCalendarFolder` / `syncTaskFolder` for calendars
 * and tasks (iCal codec, TB calendar via the vendored experiment). All
 * three share the framework in `eas/sync-runner.mjs`.
 */

import {
  ERR,
  withCode,
  ok,
  error,
  TbSyncProviderImplementation,
} from "../vendor/tbsync/provider.mjs";
import * as addressBook from "../vendor/tbsync/address-book.mjs";
import * as calendarStore from "../vendor/tbsync/calendar.mjs";
import { readOof, writeOof } from "./eas/oof.mjs";
import {
  primeAuth,
  primeAccessToken,
  forgetAuth,
  currentRefreshToken,
  startAuth,
  isOAuthAccount,
} from "./eas/oauth.mjs";
import { negotiateAsVersion } from "./eas/connect.mjs";
import { discoverEasServer } from "./eas/autodiscover.mjs";
import {
  acquirePolicyKey,
  NO_POLICY_FOR_DEVICE,
  PROVISION_EMBEDS_DEVICE_INFO,
} from "./eas/provision.mjs";
import { runFolderSync } from "./eas/folder-sync.mjs";
import { syncContactFolder } from "./eas/contact-sync.mjs";
import { syncCalendarFolder, syncTaskFolder } from "./eas/calendar-sync.mjs";
import {
  fetchUserInformation,
  sendDeviceInformation,
} from "./eas/settings.mjs";
import { runGalSearch as runGalSearchRequest } from "./eas/gal-search.mjs";
import { NET_ERR, setSyncSignalResolver } from "./network.mjs";
import {
  enableGal,
  disableGal,
  enableGalForAllAccounts,
  installRenameWatcher as installGalRenameWatcher,
} from "./gal.mjs";
import { forgetFreeBusyCache, refreshFreeBusyListener } from "./free-busy.mjs";
import { easCommandAdvertised } from "./eas/allowed-commands.mjs";
import { setEventLogSink } from "./eas-event-log.mjs";
// Cyclic: upgrades.mjs imports `easTypeToFolderType`,
// `finalizeFolderListForPush` and `iconForServerType` from this module.
// Safe because neither side touches an imported binding while the modules
// are evaluating - both only reach across inside function bodies, by which
// point both namespaces are complete.
import { runStartupMigrations } from "./upgrades.mjs";
import {
  localQueue,
  rememberBindings,
  sweep,
} from "../vendor/tbsync/change-queue.mjs";

/** This add-on's calendar type, and the scheme its calendars carry. A
 *  calendar's `type` is the id of the add-on supplying it, so this is the
 *  one part of calendar handling that cannot be shared - the vendored
 *  wrapper is told it rather than knowing it. */
const CALENDAR_TYPE = "ext-eas4tbsync@jobisoft.de";
const CALENDAR_URL = "moz-eas-calendar://";

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

/** `<Conflict>` values, sent in every Sync Options block: "1" = the
 *  server's copy wins a two-writer conflict (default), "0" = this device
 *  wins. */
const ALLOWED_CONFLICT_VALUES = ["0", "1"];

/** How long a calendar must stay quiet before an edit in it is synced, in
 *  seconds. "0" is off, and then an edit waits for the scheduled sync.
 *  Read by the calendar provider's item hooks - see `armSyncAfterChange`. */
const ALLOWED_SYNC_ON_CHANGE_VALUES = ["0", "5", "15", "30", "60"];
const DEFAULT_SYNC_ON_CHANGE = "15";

/** Setup-type → fixed EAS host. Only the OAuth setup types appear here. */
const HOST_BY_SERVERTYPE = {
  office365: "outlook.office365.com",
  "personal-ms": "eas.outlook.com",
};

/** Setup-type → per-account icon override. Sent to TbSync at register
 *  time so the manager's accounts list shows a flavour-specific icon
 *  per row (Microsoft brand for the OAuth flavours, generic EAS for
 *  Autodiscover and custom-server flavours). The setup and config
 *  popups inline the same per-type 16px URLs in their dropdown options
 *  to keep the dropdown trigger and options aligned with the manager
 *  row. */
const ICON_PATHS_BY_SERVERTYPE = {
  office365: { 16: "icons/365_16.png", 32: "icons/365_32.png" },
  "personal-ms": { 16: "icons/365_16.png", 32: "icons/365_32.png" },
  auto: { 16: "icons/eas16.png", 32: "icons/eas32.png" },
  custom: { 16: "icons/eas16.png", 32: "icons/eas32.png" },
};

/** Build the size-keyed relative-path map TbSync's REGISTER_ACCOUNT.icon
 *  expects, for the given account flavour. Returns null for unknown
 *  servertypes so the host falls back to the provider-wide icon set.
 *  TbSync resolves the paths against this extension's URL prefix at
 *  render time, so absolute URLs never reach persistent storage. */
export function iconForServerType(servertype) {
  return ICON_PATHS_BY_SERVERTYPE[servertype] ?? null;
}

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
  7: "tasks", // Default Tasks
  8: "calendars", // Default Calendar
  9: "contacts", // Default Contacts
  // 10: Default Notes (ignored)
  // 11: Default Journal (ignored)
  // 12: User-created email (ignored)
  13: "calendars", // User-created Calendar
  14: "contacts", // User-created Contacts
  15: "tasks", // User-created Tasks
};

export function easTypeToFolderType(type) {
  return EAS_TYPE_TO_TBSYNC[Number(type)] ?? null;
}

/** Class string sent in Sync Collection. */
export function folderTypeToEasClass(folderType) {
  switch (folderType) {
    case "contacts":
      return "Contacts";
    case "calendars":
      return "Calendar";
    case "tasks":
      return "Tasks";
    default:
      return null;
  }
}

export class EasProvider extends TbSyncProviderImplementation {
  /** Accounts whose config popup has saved since someone last asked.
   *  `onReauthenticate` clears an entry before opening the popup and reads
   *  it back afterwards, which turns "did the user save?" into a question
   *  about what they did rather than about which fields changed - the
   *  password field opens blank and is only submitted when typed, so a
   *  field comparison would misread "Save without retyping it". */
  #configSaves = new Set();

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
      servicesPath: "dialogs/services/services.html",
      servicesWidth: 560,
      servicesHeight: 640,
      capabilities: {
        folderTypes: ["contacts", "calendars", "tasks"],
        supportsReadOnly: true,
        multipleAccounts: true,
        hasSetupPopup: true,
        hasConfigPopup: true,
        // Opts this provider into the host's Services button. What is
        // behind it is ours; the button and its label are the host's, so
        // it reads the same for every provider that has any.
        hasServicesPopup: true,
      },
      maintainerEmail: "john.bieling@gmx.de",
      contributorsUrl: "https://github.com/jobisoft/EAS-4-TbSync",
      logPrefix: "[eas-4-tbsync]",
    });
    // Point the wire-level event-log sink at the host. Wire code
    // (`network.mjs`) emits debug entries for every WBXML send/receive;
    // before this binding the calls silently no-op.
    setEventLogSink((args) => this.reportEventLog(args));
    // Let the wire layer see the running sync's AbortSignal, so a cancel
    // drops the request in flight rather than after the server answers -
    // or, when it never does, after the connection timeout.
    setSyncSignalResolver((accountId) => this.syncSignal(accountId));
    // One-shot global listener that mirrors GAL directory renames back
    // into `account.custom.galName` so the rename survives across the
    // teardown/recreate cycle on next TB start.
    installGalRenameWatcher(this);
  }

  // ── Base-class hooks ───────────────────────────────────────────────────

  async onConnectedToHost() {
    // Bring storage and accounts up to date before anything below reads
    // account state - a legacy-imported account still holds it in the shape
    // the legacy add-on wrote. Cheap when there is nothing to do, and it
    // has to run on every port open rather than once per boot: the host
    // re-imports whenever its own storage has been cleared, and that
    // reaches us as a reconnect and nothing more.
    await runStartupMigrations(this);
    // Re-establish the per-account read-only GAL directories. The host
    // does not re-fire `onAccountEnabled` for already-enabled accounts
    // on extension boot, so the listener registrations would otherwise
    // be lost across restarts.
    await enableGalForAllAccounts(this);
    // Attendee availability. One listener for the add-on rather than one
    // per account - it joins a global service - so it is re-evaluated
    // rather than added.
    await refreshFreeBusyListener(this);
    // Now that the host is answering, reconcile our own per-folder state
    // with the bindings it names: learn where each calendar belongs, and
    // drop the queues of bindings that have ended.
    await this.#reconcileFolderSessions();
    // Watch our own address books. After the reconcile, so the bindings the
    // observer resolves against are current before the first event lands.
    addressBook.installContactsObserver({
      provider: this,
      report: (args) => this.reportEventLog(args),
    });
    return null;
  }

  /** Line our own storage up with the bindings the host currently names.
   *
   *  Two things at once, from one read of the folder rows:
   *
   *  Remembering, so an item hook can file a user edit against the right
   *  folder without a round trip - it is handed a calendar id and nothing
   *  else, and the host may be gone by the time the user hits save.
   *
   *  Sweeping, which is the entire teardown path for what we store. A
   *  folder deselected, an account disconnected or removed, a calendar
   *  deleted - the host ends each of those by minting a new session id and
   *  telling nobody. It cannot tell us: those flows have to work while this
   *  add-on is broken or uninstalled, so they cannot wait on it. We find
   *  out by looking.
   *
   *  Reads every account before deciding anything. A partial answer would
   *  sweep live queues away, so any failure abandons the whole pass - the
   *  garbage keeps until the next port open, which costs nothing: a queue
   *  whose session nothing names is never read. */
  async #reconcileFolderSessions() {
    const liveSessions = new Set();
    const liveTargets = new Set();
    const bindings = [];
    try {
      for (const { accountId } of await this.listAccounts()) {
        const { folders = [] } = (await this.getAccount(accountId)) ?? {};
        for (const f of folders) {
          if (!f?.sessionId) continue;
          liveSessions.add(f.sessionId);
          if (f.targetID) {
            liveTargets.add(f.targetID);
            bindings.push({
              targetID: f.targetID,
              accountId,
              folderId: f.folderId,
              sessionId: f.sessionId,
              targetType: f.targetType,
            });
          }
        }
      }
    } catch (err) {
      this.reportEventLog({
        level: "debug",
        message: `[queue] could not read folder sessions; reconcile skipped: ${err?.message ?? String(err)}`,
      });
      return;
    }
    try {
      await rememberBindings(bindings);
      const { queues, orphans } = await sweep({ liveSessions, liveTargets });
      for (const d of queues) {
        this.reportEventLog({
          level: d.entries > 0 ? "info" : "debug",
          accountId: d.accountId ?? undefined,
          folderId: d.folderId ?? undefined,
          message:
            `[queue] dropped the change queue of a binding that no longer ` +
            `exists (${d.entries} pending edit(s) went with it)`,
        });
      }
      // Whatever those bindings pointed at is no longer synced by anything.
      // The host marks the ones that still exist; the rest it skips.
      await this.reportOrphanedTargets(orphans);
    } catch (err) {
      this.reportEventLog({
        level: "warning",
        message: `[queue] reconcile failed: ${err?.message ?? String(err)}`,
      });
    }
  }

  /** The pending edits we hold for a folder. Read-only, and the only way
   *  anyone outside this add-on can see a provider-owned queue - the folder
   *  row's own changelog stays empty for these. */
  async onGetChangelog({ accountId, folderId }) {
    const { folders = [] } = (await this.getAccount(accountId)) ?? {};
    const folder = folders.find((f) => f.folderId === folderId);
    if (!folder) return null;
    if (!folder.sessionId) return null;
    return localQueue({
      accountId,
      folderId,
      sessionId: folder.sessionId,
    }).entries();
  }
  // ── Account lifecycle ──────────────────────────────────────────────────

  async onAccountEnabled({ accountId }) {
    // First connect after setup (or re-enable after disable): negotiate
    // EAS version, run Provision if required, run FolderSync to discover
    // the resource list, and push it to the host so the manager can
    // render contacts/calendars/tasks rows. Idempotent on re-enable.
    await this.#connectAndDiscoverFolders(accountId);
    // OPTIONS has populated allowedEasCommands by now, so we know
    // whether the server supports the GAL Search command.
    const rv = await this.getAccount(accountId);
    if (rv?.account) await enableGal({ provider: this, account: rv.account });
    await refreshFreeBusyListener(this);
    return null;
  }

  async onAccountDisabled({ accountId }) {
    // Ahead of the early return below, so the cache is dropped even when
    // the account row has already gone. Re-enable primes again from
    // `custom`, so nothing here is needed to come back.
    forgetAuth(accountId);
    await disableGal({ provider: this, accountId });
    forgetFreeBusyCache(accountId);
    await refreshFreeBusyListener(this);
    const ctx = await this.#loadContext(accountId);
    if (!ctx) return null;
    // No target deletion here: the host owns resource deletion in every
    // flow and deletes them right after this returns. This hook drops
    // provider-side state only - and that division is what makes the
    // disconnect a recovery path, because it works the same when this
    // provider is not around to run it. Account-level credentials and
    // deviceId stay so re-enable works without re-setup; the host wipes
    // its folder rows after this returns, so per-folder custom.* doesn't
    // need clearing here.
    // Force a fresh FolderSync on re-enable. Also invalidate the OPTIONS
    // probe cache so the next enable re-runs the probe - this is the
    // backfill path for users on 5.0.1 whose `allowedEasCommands` was
    // never persisted (or has otherwise diverged from the server's
    // current capability list). Existing-enabled accounts are left alone.
    // The mailbox address goes too: an account re-pointed at a different
    // login or server would otherwise keep naming the old mailbox as its
    // calendars' owner, and a wrong owner is worse than none.
    await this.updateAccount({
      accountId,
      patch: {
        custom: {
          foldersynckey: "0",
          lastEasOptionsUpdate: 0,
          userSmtpAddress: "",
          userDisplayName: "",
          noUserInformation: false,
        },
      },
    }).catch((err) =>
      console.debug(
        `[eas] updateAccount(disable-reset) for ${accountId} failed:`,
        err,
      ),
    );
    return null;
  }

  async onAccountDeleted({ accountId }) {
    // Same contract as onAccountDisabled: stop, drop provider state, never
    // touch resources - the host deletes them (or keeps them, its call)
    // and then forgets the account entirely.
    forgetAuth(accountId);
    await disableGal({ provider: this, accountId });
    forgetFreeBusyCache(accountId);
    await refreshFreeBusyListener(this);
    return null;
  }

  async onRegisterSuccessful({ accountId }) {
    // OPTIONS probe up-front so the manager UI / config popup have the
    // server-advertised EAS version and command list (notably Search/GAL)
    // before the user clicks Connect. The host broadcasts accounts-changed
    // before this hook resolves, so the new (still-disabled) row is already
    // visible - and holds this account "being prepared" for as long as we
    // are in here, so the user cannot race us by clicking Connect.
    try {
      const ctx = await this.#loadContext(accountId);
      if (!ctx) return null;
      if (isOAuthAccount(ctx.account.custom)) this.#primeAuth(ctx);
      const negotiated = await negotiateAsVersion({ account: ctx.account });
      await this.updateAccount({
        accountId,
        patch: {
          custom: {
            asversion: negotiated.asVersion,
            allowedEasVersions: negotiated.allowedAsVersions,
            allowedEasCommands: negotiated.allowedCommands,
            lastEasOptionsUpdate: Date.now(),
          },
        },
      });
    } catch (err) {
      this.reportEventLog({
        level: "warning",
        message: `Initial OPTIONS probe failed for account ${accountId}; will retry on first Connect`,
        details: err?.message ?? null,
      });
    }
    return null;
  }

  // ── Re-authentication ─────────────────────────────────────────────────

  /** Re-run the consent flow for an account the host has stamped E:AUTH.
   *
   *  Returns a StatusData rather than throwing: the host only clears the
   *  error and re-enables the account when this resolves with a success
   *  envelope, so a bare return would leave the account frozen even after
   *  a successful sign-in. By the time this runs the host has already sent
   *  ACCOUNT_DISABLED and dropped the folder list, so only `custom` is
   *  ours to touch - ACCOUNT_ENABLED rebuilds the rest. */
  async onReauthenticate({ accountId }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) return error("Unknown account", ERR.UNKNOWN_ACCOUNT);

    const c = ctx.account.custom ?? {};
    if (!isOAuthAccount(c)) {
      // The manager offers this action for any E:AUTH account, but a
      // username/password account has no consent flow to re-run - a
      // rejected password is what put it here. Open its Settings instead,
      // and report success once the user saves so the host clears the
      // error and re-enables the account.
      //
      // Reporting success is the whole point. The manager greys out
      // Connect while the error stands and nothing else clears it, so
      // correcting the password would otherwise leave the account stuck
      // for good. If the new password is wrong too, the next sync stamps
      // E:AUTH again and the button comes back.
      this.#configSaves.delete(accountId);
      await this.onOpenConfigPopup({ accountId });
      if (!(await this.#loadContext(accountId))) {
        return error("Unknown account", ERR.UNKNOWN_ACCOUNT);
      }
      if (!this.#configSaves.delete(accountId)) {
        // Closed without saving - treat it exactly like dismissing the
        // consent window, so the host logs nothing and the account keeps
        // its error rather than briefly looking healthy.
        return error("Settings closed without saving", ERR.CANCELLED);
      }
      return ok();
    }

    const knownEmail = c.authenticatedUserEmail || null;
    // Held so the finally can unregister the exact window it registered -
    // the base class removes by windowId, so that a second window can never
    // orphan the first.
    let consentWindowId = null;
    try {
      const { refreshToken, authenticatedUserEmail, accessToken, expiresIn } =
        await startAuth({
          loginHint: knownEmail || c.user || undefined,
          servertype: c.servertype,
          onWindowCreated: (windowId) => {
            consentWindowId = windowId;
            this.registerAccountWindow(accountId, windowId);
          },
        });

      // Signing in as somebody else would silently repoint the account at
      // a different mailbox while keeping its existing folders and targets.
      if (
        knownEmail &&
        authenticatedUserEmail &&
        knownEmail !== authenticatedUserEmail
      ) {
        return error(
          `Signed-in user (${authenticatedUserEmail}) does not match this account (${knownEmail}).`,
          ERR.AUTH,
        );
      }

      await this.updateAccount({
        accountId,
        patch: {
          custom: {
            refreshToken,
            authenticatedUserEmail: authenticatedUserEmail ?? knownEmail,
          },
        },
      });
      forgetAuth(accountId);
      primeAuth(accountId, { refreshToken, servertype: c.servertype });
      // startAuth already returned a usable access token; keeping it saves
      // the next sync a refresh round-trip.
      if (accessToken) primeAccessToken(accountId, accessToken, expiresIn);
      return ok();
    } catch (err) {
      // The code rides through in `details`, which is what the host reads
      // to stay quiet when the user simply closed the popup (ERR.CANCELLED).
      return error(
        err?.message ?? "Re-authentication failed",
        err?.code ?? ERR.AUTH,
      );
    } finally {
      this.unregisterAccountWindow(accountId, consentWindowId);
    }
  }

  // ── Folder lifecycle ──────────────────────────────────────────────────

  async onFolderEnabled() {
    return null;
  }

  async onFolderDisabled({ accountId, folderId }) {
    const folder = await this.#getFolder(accountId, folderId);
    if (!folder) return null;
    // The calendar or book itself is the host's to delete, right after this
    // returns - this hook only unhooks.
    // Clear the binding and the sync state, but keep `targetName` and
    // `targetColor`. The calendar or book is gone, so the id and the sync
    // position describe nothing and must not survive; the name and colour
    // describe what the user *chose*, and enabling the resource again should
    // give them back what they had rather than a freshly generated name and
    // an arbitrary colour. Nothing else remembers: no ActiveSync folder
    // element carries either, so clearing them here loses them for good.
    await this.updateFolder({
      accountId,
      folderId,
      patch: {
        targetID: null,
        custom: { synckey: "0", indexMap: [] },
      },
    }).catch((err) =>
      console.debug(
        `[eas] updateFolder(disable-reset) for ${accountId}/${folderId} failed:`,
        err,
      ),
    );
    return null;
  }

  // ── Sync ──────────────────────────────────────────────────────────────

  async onSyncAccount({ accountId }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) throw withCode(new Error("unknown account"), ERR.UNKNOWN_ACCOUNT);
    this.reportSyncState({ accountId, syncState: "prepare" });
    const oauth = isOAuthAccount(ctx.account.custom);
    const storedRefreshToken = ctx.account.custom?.refreshToken;
    try {
      // Refresh the folder list each sync so server-side additions surface.
      // Items themselves are synced per folder, by onSyncFolder.
      await this.#connectAndDiscoverFolders(accountId);
      return ok();
    } finally {
      // Also on the failure path: a sync can rotate the token and then fail
      // for an unrelated reason, and the rotated token is still the good one.
      if (oauth)
        await this.#persistRotatedRefreshToken(accountId, storedRefreshToken);
    }
  }

  /** OPTIONS probe (once) → pre-emptive Provision (if user-toggled) →
   *  Settings/DeviceInformation → FolderSync (with 449 → Provision +
   *  Settings retry) → push folders → persist synckey. Called from both
   *  `onAccountEnabled` (initial connect) and `onSyncAccount` (every
   *  refresh).
   *
   *  HTTP 451 (X-MS-Location host migration) is caught at the top and
   *  triggers a one-shot retry against the new host. A network-layer
   *  failure on a `servertype === "auto"` account triggers a one-shot
   *  Autodiscover re-run in case the cached MobileSync URL has rotated. */
  async #connectAndDiscoverFolders(
    accountId,
    redirectsRemaining = 1,
    rediscoversRemaining = 1,
  ) {
    try {
      await this.#doConnectAndDiscover(accountId);
    } catch (err) {
      if (
        err.code === NET_ERR.HOST_REDIRECT &&
        err.newLocation &&
        redirectsRemaining > 0
      ) {
        await this.updateAccount({
          accountId,
          patch: { custom: { server: err.newLocation } },
        });
        await this.#connectAndDiscoverFolders(
          accountId,
          redirectsRemaining - 1,
          rediscoversRemaining,
        );
        return;
      }
      if (err.code === NET_ERR.NETWORK && rediscoversRemaining > 0) {
        const rediscovered = await this.#rediscoverServerUrl(accountId);
        if (rediscovered) {
          await this.#connectAndDiscoverFolders(
            accountId,
            redirectsRemaining,
            rediscoversRemaining - 1,
          );
          return;
        }
      }
      throw err;
    }
  }

  /** Re-run Autodiscover for an `auto`-type account whose cached server
   *  URL just failed with E:NETWORK. Returns `true` when a different URL
   *  was found and persisted (caller should retry the connect once),
   *  `false` for any reason that should leave the original error to
   *  bubble up: account isn't `auto`, no cached password, autodiscover
   *  itself failed, or the rediscovered URL matches what we already had
   *  (a spurious-fault loop guard). */
  async #rediscoverServerUrl(accountId) {
    const ctx = await this.#loadContext(accountId);
    const c = ctx?.account?.custom;
    if (!c || c.servertype !== "auto" || !c.user || !c.password) return false;
    let result;
    try {
      result = await discoverEasServer({ email: c.user, password: c.password });
    } catch (err) {
      this.reportEventLog({
        level: "warning",
        accountId,
        message: `[autodiscover] runtime fallback failed: ${err?.message ?? String(err)}`,
      });
      return false;
    }
    const newUrl = result?.server;
    if (!newUrl || newUrl === c.server) return false;
    await this.updateAccount({
      accountId,
      patch: { custom: { server: newUrl } },
    });
    this.reportEventLog({
      level: "debug",
      accountId,
      message: `[autodiscover] rotated server URL: ${c.server} -> ${newUrl}`,
    });
    return true;
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
    const lastOptionsUpdate = Number(
      ctx.account.custom?.lastEasOptionsUpdate ?? 0,
    );
    const needsOptionsProbe =
      !ctx.account.custom?.asversion ||
      Date.now() - lastOptionsUpdate > OPTIONS_REPROBE_MS;
    if (needsOptionsProbe) {
      const negotiated = await negotiateAsVersion({ account: ctx.account });
      // Sticky: a connected account KEEPS the version it was negotiated
      // onto for as long as the server still advertises it. The re-probe
      // tracks the server's list; it does not migrate live accounts to
      // whatever the preference order currently picks - a protocol
      // switch changes item-identity semantics (UID handling differs
      // between 14.x and 16.1), so it happens only when the server
      // drops the stored version. The preference order thus reaches only
      // fresh negotiations.
      const current = ctx.account.custom?.asversion;
      const asversion =
        current && negotiated.allowedAsVersions.includes(current)
          ? current
          : negotiated.asVersion;
      await this.updateAccount({
        accountId,
        patch: {
          custom: {
            asversion,
            allowedEasVersions: negotiated.allowedAsVersions,
            allowedEasCommands: negotiated.allowedCommands,
            lastEasOptionsUpdate: Date.now(),
          },
        },
      });
      ctx = await this.#loadContext(accountId);
    }

    // The user can override negotiation via the config popup. "auto"
    // (the default) uses the negotiated value cached in account.custom.
    // When a specific version is forced, validate it against the server's
    // advertised list - matches legacy sync.js:107-108. We only enforce
    // when `allowedEasVersions` is non-empty so a corrupted-state account
    // isn't blocked from re-probing.
    const selected = ctx.account.custom?.asversionselected || "auto";
    const allowedVersions = ctx.account.custom?.allowedEasVersions ?? [];
    if (
      selected !== "auto" &&
      allowedVersions.length > 0 &&
      !allowedVersions.includes(selected)
    ) {
      throw withCode(
        new Error(
          `Selected EAS version ${selected} is not advertised by the server (server: ${allowedVersions.join(", ")})`,
        ),
        ERR.UNKNOWN_COMMAND,
      );
    }
    const asVersion =
      selected === "auto" ? ctx.account.custom.asversion : selected;

    // 2) Pre-emptive Provision (legacy "Kerio" semantics). When the
    // user has flipped the toggle on - or a previous 449 stuck it on -
    // and we have no policy key cached, run Provision before any other
    // command. Servers that don't return 449 (e.g. Kerio) need this.
    if (
      ctx.account.custom?.provision === true &&
      (ctx.account.custom?.policykey ?? "0") === "0"
    ) {
      const result = await acquirePolicyKey({
        account: ctx.account,
        asVersion,
      });
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
    // the OPTIONS probe. Wrapped in provision-then-retry so a server
    // that demands Provision before accepting Settings (Z-Push on
    // first connect, Kerio, some Exchange configs) self-heals via a
    // single re-Provision pass.
    ctx = await this.#withProvisionRecovery(
      accountId,
      ctx,
      asVersion,
      async (c) => {
        await this.#maybeSendDeviceInformation(c.account, asVersion);
        return c;
      },
    );

    // 4) FolderSync, with provision/sync-key recovery loops.
    //
    // A non-zero synckey with zero host folder rows is a contradiction: an
    // incremental FolderSync only makes sense when we hold the folders the
    // increments apply to. The state is real, not hypothetical - a
    // disconnect while this provider was dead tears down host-side only, so
    // onAccountDisabled's foldersynckey reset never runs, and the stale key
    // makes the next enable pull "no changes" and announce nothing. The
    // account then looks connected but has no resources. Trust the row
    // count over the key: start over.
    let priorFolderSyncKey = ctx.account.custom?.foldersynckey ?? "0";
    if (priorFolderSyncKey !== "0" && ctx.folders.length === 0) {
      this.reportEventLog({
        level: "info",
        accountId,
        message:
          "[eas] FolderSync key present but no folders on record - the " +
          "account was torn down while this provider could not hear it; " +
          "starting folder discovery over",
      });
      priorFolderSyncKey = "0";
      ctx.account.custom.foldersynckey = "0";
    }
    const { syncResult, ctx: ctxAfterSync } =
      await this.#runFolderSyncWithRecovery(accountId, ctx, asVersion);
    ctx = ctxAfterSync;

    // Between the wire work above and the writes below: the host may have
    // cancelled this sync while our FolderSync was in flight - a fetch that
    // had already resolved never feels the abort. Writing anyway is how a
    // cancelled sync corrupted a disconnect: its synckey persist landed on
    // top of the foldersynckey:"0" reset that onAccountDisabled had just
    // written, so the next enable ran an incremental FolderSync, saw no
    // changes, pushed no folders, and the account came back empty.
    this.throwIfCancelled(accountId);

    // 5) Apply Add/Update/Delete to the host's folder list.
    //    - Initial sync (priorFolderSyncKey === "0"): server emits every
    //      folder as an Add; push the fresh list.
    //    - Incremental sync (priorFolderSyncKey !== "0"): merge the
    //      delta into the existing folder list and push the result.
    //      Skipping the push on no-op deltas avoids an unnecessary
    //      storage write + broadcast.
    if (priorFolderSyncKey === "0") {
      const initial = syncResult.adds
        .map((a) => folderDescriptorFromAdd(a))
        .filter(Boolean);
      await this.pushFolderList({
        accountId,
        folders: await finalizeFolderListForPush(initial),
      });
    } else if (
      syncResult.adds.length ||
      syncResult.updates.length ||
      syncResult.deletes.length
    ) {
      const merged = await mergeFolderDeltas(ctx.folders, syncResult);
      // The host deletes the targets behind dropped folders itself, inside
      // its PUSH_FOLDER_LIST handler - it always did for address books, and
      // owns calendars too now.
      await this.pushFolderList({ accountId, folders: merged });
    }

    // 6) Persist the new FolderSync continuation key.
    this.throwIfCancelled(accountId);
    await this.updateAccount({
      accountId,
      patch: { custom: { foldersynckey: syncResult.synckey } },
    });

    // 7) Settings/UserInformation, once per account. After FolderSync,
    // not before: on 14.1 and later nothing above sends a Settings
    // request at all - DeviceInformation rides inside Provision there -
    // so this would otherwise be the first request a fresh account
    // makes, and a server that demands Provision first would refuse it.
    // FolderSync has the recovery loop that gets us provisioned, so by
    // here the question can actually be answered. It swallows its own
    // failures, which is why it sits outside that loop rather than in
    // it.
    await this.#maybeLearnUserAddress(accountId, ctx.account, asVersion);
  }

  /** Run `op(ctx)` with one provision-then-retry on PROVISION_REQUIRED.
   *  Recognizes both transport-level `NET_ERR.PROVISION_REQUIRED`
   *  (HTTP 449 out-of-band, Settings 141-144 in-band, Search 141-144
   *  in-band) and FolderSync in-band Status 141-144 carried on
   *  `err.folderSyncStatus`. After one failed retry, the thrown error
   *  escapes unchanged. The new ctx (post-Provision policy key) is
   *  threaded into the retry call. */
  async #withProvisionRecovery(accountId, ctx, asVersion, op) {
    let provisioned = false;
    while (true) {
      try {
        return await op(ctx);
      } catch (err) {
        const provisionRequired =
          err.code === NET_ERR.PROVISION_REQUIRED ||
          PROVISION_REQUIRED_STATUSES.has(err.folderSyncStatus);
        if (provisionRequired && !provisioned) {
          provisioned = true;
          ctx = await this.#provisionAndPersist(accountId, ctx, asVersion);
          continue;
        }
        throw err;
      }
    }
  }

  /** FolderSync with recovery for HTTP 449 / in-band Status 141-144
   *  (provision required, via `#withProvisionRecovery`) and Status 9
   *  (invalid sync key, reset to "0" and treat as initial). */
  async #runFolderSyncWithRecovery(accountId, ctx, asVersion) {
    return this.#withProvisionRecovery(accountId, ctx, asVersion, async (c) => {
      let resetSyncKey = false;
      while (true) {
        try {
          const syncResult = await runFolderSync({
            account: c.account,
            asVersion,
          });
          return { syncResult, ctx: c };
        } catch (err) {
          if (err.folderSyncStatus === "9" && !resetSyncKey) {
            resetSyncKey = true;
            await this.updateAccount({
              accountId,
              patch: { custom: { foldersynckey: "0" } },
            });
            c = await this.#loadContext(accountId);
            continue;
          }
          throw err;
        }
      }
    });
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
    // 14.1/16.0/16.1 carry DeviceInformation inside the initial
    // Provision body and forbid the separate Settings command for it.
    if (PROVISION_EMBEDS_DEVICE_INFO.has(asVersion)) return;
    if (!easCommandAdvertised(account, "Settings")) return;
    await sendDeviceInformation({ account, asVersion });
  }

  /** Learn which mailbox this account is, once, and remember it.
   *
   *  The address is what tells a calendar whose it is (see
   *  `calendarOwner`) and what decides organizer-vs-attendee on the
   *  wire; the EAS login cannot be trusted for either, since it may be
   *  `DOMAIN\user` or any other non-address. Only the version gate
   *  applies here - unlike DeviceInformation, `UserInformation` stays a
   *  Settings sub-command on every version that has the command.
   *
   *  Asked once. A server that answers with no address is recorded as
   *  such, because asking it again every sync would spend a request and
   *  a log line forever on an answer that will not change; a transient
   *  failure throws instead and is retried on the next sync. Both keys
   *  are cleared when the account is disabled, so re-pointing it at
   *  another mailbox asks again.
   *
   *  Nothing here may fail a sync: a calendar's owner is not worth an
   *  account going red over. */
  async #maybeLearnUserAddress(accountId, account, asVersion) {
    if (account.custom?.userSmtpAddress) return false;
    if (account.custom?.noUserInformation) return false;
    if (asVersion === "2.5") return false;
    if (!easCommandAdvertised(account, "Settings")) return false;

    let info;
    try {
      info = await fetchUserInformation({ account, asVersion });
    } catch (err) {
      this.reportEventLog({
        level: "debug",
        accountId,
        message: `[eas] could not read UserInformation, will retry: ${err?.message ?? String(err)}`,
      });
      return false;
    }

    const patch = info?.address
      ? {
          userSmtpAddress: info.address,
          userDisplayName: info.displayName ?? "",
        }
      : { noUserInformation: true };
    try {
      await this.updateAccount({ accountId, patch: { custom: patch } });
    } catch (err) {
      this.reportEventLog({
        level: "debug",
        accountId,
        message: `[eas] could not store the mailbox address: ${err?.message ?? String(err)}`,
      });
      return false;
    }

    this.reportEventLog({
      level: info?.address ? "info" : "debug",
      accountId,
      message: info?.address
        ? `[eas] this account syncs the mailbox ${info.address}`
        : "[eas] the server names no address for this account; its " +
          "calendars keep Thunderbird's default owner",
    });
    return !!info?.address;
  }

  /** GAL contact lookup with provision-then-retry recovery. Loaded by
   *  `gal.mjs` and invoked from the address-book search callback - that
   *  path runs outside any handshake, so we acquire ctx here and route
   *  through `#withProvisionRecovery` so a server that has expired the
   *  policy key gets one Provision attempt before the search returns
   *  empty. */
  async runGalSearch({ accountId, query, companyName }) {
    const ctx = await this.#loadContext(accountId);
    const asVersion = ctx.account.custom?.asversion;
    return await this.#withProvisionRecovery(
      accountId,
      ctx,
      asVersion,
      async (c) =>
        runGalSearchRequest({
          account: c.account,
          asVersion,
          query,
          companyName,
        }),
    );
  }

  #primeAuth(ctx) {
    const c = ctx.account.custom ?? {};
    primeAuth(ctx.account.accountId, {
      refreshToken: c.refreshToken,
      servertype: c.servertype,
    });
  }

  /** Write back a refresh token Microsoft rotated during this sync.
   *
   *  A refresh can hand us a new refresh token, which the OAuth layer keeps
   *  in memory. The background page is persistent, so that lasts the whole
   *  session - but the next start primes from `custom`, so without this the
   *  account comes back on a token the server may already have retired.
   *
   *  Done after the sync rather than at the moment of rotation, which keeps
   *  a storage write off the token-refresh path. Rotation is rare, so this
   *  is a comparison that almost always does nothing. */
  async #persistRotatedRefreshToken(accountId, storedRefreshToken) {
    const live = currentRefreshToken(accountId);
    if (!live || live === storedRefreshToken) return;
    try {
      await this.updateAccount({
        accountId,
        patch: { custom: { refreshToken: live } },
      });
      // Rotation cannot be provoked - the server decides when to hand back
      // a new token - so the only way to confirm this path works is to
      // catch it happening. Without a line here that means noting the
      // stored token, using the account for days, and diffing storage by
      // hand, with a *negative* as the pass condition: no re-authentication
      // prompt after a restart. This turns that into something readable,
      // and puts a marker just before any prompt that does appear.
      //
      // The token itself is never logged, in whole or in part: an event log
      // ends up in bug reports.
      this.reportEventLog({
        level: "debug",
        accountId,
        message:
          "[oauth] the server rotated the refresh token; stored the new one",
      });
    } catch (err) {
      // Losing the write costs a re-auth at worst; failing the sync that
      // otherwise succeeded would be the bigger harm.
      console.debug(
        `[eas] persisting rotated refresh token for ${accountId} failed:`,
        err,
      );
    }
  }

  async onGetSortedFolders({ accountId }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) return [];
    const sorted = ctx.folders
      .slice()
      .sort((a, b) => (a.orderIndex ?? 0) - (b.orderIndex ?? 0));
    return await finalizeFolderListForPush(sorted);
  }

  async onSyncFolder({ accountId, folderId }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) throw withCode(new Error("unknown account"), ERR.UNKNOWN_ACCOUNT);
    let folder = ctx.folders.find((f) => f.folderId === folderId);
    if (!folder)
      throw withCode(new Error("unknown folder"), ERR.UNKNOWN_FOLDER);

    const tt = folder.targetType;
    if (tt !== "contacts" && tt !== "calendars" && tt !== "tasks") {
      // Folder kind we don't sync yet (e.g. notes / journal): record a
      // success transition without touching server state.
      this.reportSyncState({ accountId, folderId, syncState: "sync" });
      return ok();
    }

    const oauth = isOAuthAccount(ctx.account.custom);
    const storedRefreshToken = ctx.account.custom?.refreshToken;
    if (oauth) this.#primeAuth(ctx);

    try {
      // Lazy-bind the local TB target on first sync (or after the user
      // removed it manually).
      if (tt === "contacts") {
        if (
          !folder.targetID ||
          !(await addressBook.bookExists(folder.targetID))
        ) {
          const name = localNameForFolder(folder, ctx);
          const targetID = await addressBook.createBook(name);
          await this.updateFolder({
            accountId,
            folderId,
            patch: { targetID, targetName: name },
          });
          folder = { ...folder, targetID, targetName: name };
        }
      } else {
        if (
          !folder.targetID ||
          !(await calendarStore.calendarExists(folder.targetID))
        ) {
          const name = localNameForFolder(folder, ctx);
          // The colour the user last gave this folder, or a fresh one from
          // the palette. Written back either way, so an assigned colour is
          // remembered from then on and the calendar keeps it through the
          // next disable/enable.
          const color =
            folder.targetColor || (await calendarStore.pickCalendarColor());
          const targetID = await calendarStore.createCalendar({
            name,
            kind: tt === "calendars" ? "events" : "tasks",
            color,
            type: CALENDAR_TYPE,
            url: CALENDAR_URL,
            ...calendarOwner(ctx.account),
          });
          await this.updateFolder({
            accountId,
            folderId,
            patch: { targetID, targetName: name, targetColor: color },
          });
          folder = {
            ...folder,
            targetID,
            targetName: name,
            targetColor: color,
          };
        }
        // Thunderbird's own refresh timer duplicates the schedule this
        // add-on already runs. New calendars are created with it off; this
        // reaches the ones that are not, including any where the user has
        // set an interval in the calendar's properties dialog since.
        await calendarStore.suppressCalendarRefresh(folder.targetID);

        // Exchange invites the attendees of a meeting we push, so Thunderbird
        // must not mail them as well. New calendars say so from the start;
        // this reaches the ones made before that.
        await calendarStore
          .deferSchedulingToServer(folder.targetID)
          .then((written) => {
            if (written) {
              this.reportEventLog({
                level: "info",
                accountId,
                folderId,
                message:
                  "[eas] invitations for this calendar are left to the server",
              });
            }
          })
          .catch((err) =>
            this.reportEventLog({
              level: "debug",
              accountId,
              folderId,
              message: `[eas] could not defer scheduling to the server: ${err?.message ?? String(err)}`,
            }),
          );

        // Calendars made before the account knew its own address carry no
        // owner; declaring it here reaches them without a recreate.
        const owner = calendarOwner(ctx.account);
        await calendarStore
          .setCalendarOwner(folder.targetID, owner)
          .then((written) => {
            // Only on a change - `setCalendarOwner` compares first, so a
            // line here means the calendar just learned something, not
            // that it is asked every sync.
            if (written) {
              this.reportEventLog({
                level: "info",
                accountId,
                folderId,
                message: `[eas] this calendar now belongs to ${owner.organizer}`,
              });
            }
          })
          .catch((err) =>
            this.reportEventLog({
              level: "debug",
              accountId,
              folderId,
              message: `[eas] could not declare the calendar owner: ${err?.message ?? String(err)}`,
            }),
          );
      }

      this.reportSyncState({ accountId, folderId, syncState: "sync" });

      const args = {
        provider: this,
        account: ctx.account,
        folder,
        accountId,
        folderId,
        asVersion: ctx.account.custom?.asversion ?? "14.1",
      };
      if (tt === "contacts") return await syncContactFolder(args);
      if (tt === "calendars") return await syncCalendarFolder(args);
      return await syncTaskFolder(args);
    } finally {
      // Item sync is where a long run is most likely to cross a token
      // refresh, so this is the folder hook's reason for a finally.
      if (oauth)
        await this.#persistRotatedRefreshToken(accountId, storedRefreshToken);
    }
  }

  // ── Setup / config popup backings ─────────────────────────────────────

  /**
   * Create a new account from setup.html. Branches on `servertype`:
   *
   *   "office365" | "personal-ms": OAuth. Host is fixed by setup type
   *      (HOST_BY_SERVERTYPE), server URL is derived from host, refresh
   *      token comes from the consent popup.
   *   "auto": basic auth. UI ran Autodiscover and supplies the resolved
   *      server URL alongside the email + password.
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
      const server = `https://${HOST_BY_SERVERTYPE[servertype]}/Microsoft-Server-ActiveSync`;
      const user = authenticatedUserEmail || args.loginHint || "";
      return {
        accountName: trimmedLabel,
        icon: iconForServerType(servertype),
        initialFolders: [],
        custom: {
          servertype,
          server,
          user,
          // Plain-old basic-auth fields stay empty for OAuth accounts.
          password: "",
          // OAuth-only fields:
          refreshToken,
          authenticatedUserEmail: authenticatedUserEmail ?? null,
          // Common EAS state:
          deviceId: generateDeviceId(),
          asversion: "",
          policykey: "0",
          foldersynckey: "0",
          // Legacy semantic: off by default. Self-corrects to true on
          // the first 449. The user-visible toggle in the config popup
          // is a pre-emptive override for servers that need provisioning
          // but don't return 449 (e.g. Kerio).
          provision: false,
          conflict: "1",
          syncrecurrence: true,
          syncOnChange: DEFAULT_SYNC_ON_CHANGE,
        },
      };
    }

    if (servertype === "auto") {
      const email = String(args.email ?? args.user ?? "").trim();
      const server = String(args.server ?? "").trim();
      if (!email) throw new Error("Email is required");
      if (!server) throw new Error("Server URL is required");
      if (!args.password) throw new Error("Password is required");
      const label = String(args.label ?? "").trim() || email;
      return {
        accountName: label,
        icon: iconForServerType("auto"),
        initialFolders: [],
        custom: {
          servertype: "auto",
          server,
          user: email,
          password: args.password,
          deviceId: generateDeviceId(),
          asversion: "",
          policykey: "0",
          foldersynckey: "0",
          provision: false,
          conflict: "1",
          syncrecurrence: true,
          syncOnChange: DEFAULT_SYNC_ON_CHANGE,
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
      icon: iconForServerType("custom"),
      initialFolders: [],
      custom: {
        servertype: "custom",
        server: trimmedServer,
        user: trimmedUser,
        password: args.password,
        deviceId: generateDeviceId(),
        asversion: "",
        policykey: "0",
        foldersynckey: "0",
        provision: false,
        conflict: "1",
        syncrecurrence: true,
        syncOnChange: DEFAULT_SYNC_ON_CHANGE,
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
      // Protocol section.
      deviceId: c.deviceId ?? "",
      asVersion: c.asversion ?? "",
      allowedAsVersions: Array.isArray(c.allowedEasVersions)
        ? c.allowedEasVersions
        : [],
      asVersionSelected: c.asversionselected ?? "auto",
      // Legacy default is off. The 449 self-correction path will flip
      // it on automatically when the server demands provisioning.
      provision: !!c.provision,
      // Absent (pre-option accounts) means server wins - the value the
      // sync also falls back to.
      conflict: c.conflict === "0" ? "0" : "1",
      // Contacts section.
      contactsDisplayOverride: !!c.displayoverride,
      contactsNameSeparator: c.seperator || "10",
      // Calendar section.
      calendarSyncLimit: c.synclimit || "7",
      syncRecurrence: !!c.syncrecurrence,
      // Anything unrecognised, including the absent value of an account
      // that predates the setting, reads as the default - the popup must
      // never render a select with no option selected.
      syncOnChange: ALLOWED_SYNC_ON_CHANGE_VALUES.includes(c.syncOnChange)
        ? c.syncOnChange
        : DEFAULT_SYNC_ON_CHANGE,
      // GAL section. `galEnabled` defaults to true so accounts that
      // pre-date the toggle keep their auto-on behavior; only an
      // explicit false (set via this dialog) disables it. `galSupported`
      // tells the dialog whether to enable the checkbox at all.
      galEnabled: c.galenabled !== false,
      galSupported: easCommandAdvertised(ctx.account, "Search"),
      // Attendee availability, same defaulting. Unsupported on 2.5,
      // which has no Availability element in ResolveRecipients.
      freeBusyEnabled: c.freebusyenabled !== false,
      freeBusySupported:
        c.asversion !== "2.5" &&
        easCommandAdvertised(ctx.account, "ResolveRecipients"),
    };
  }

  /**
   * Services: the mailbox's out-of-office reply, read from the server.
   *
   * No caching and no stored copy - a service is the server's state, and a
   * stale one shown in a dialog is worse than a slow one. The window asks
   * on open and again after every save.
   *
   * Returns null when the server answers without an Oof block at all,
   * which is how a mailbox that does not offer one reads.
   */
  async readOofForServices(accountId) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) throw withCode(new Error("Unknown account"), ERR.UNKNOWN_ACCOUNT);
    // Every request path has to do this, not just the sync ones: without
    // it the network layer refuses an OAuth account outright rather than
    // refreshing its token.
    if (isOAuthAccount(ctx.account.custom)) this.#primeAuth(ctx);
    return readOof({
      account: ctx.account,
      asVersion: ctx.account.custom?.asversion ?? "14.1",
    });
  }

  /** Services: write the out-of-office reply back. */
  async writeOofFromServices({ accountId, settings }) {
    const ctx = await this.#loadContext(accountId);
    if (!ctx) throw withCode(new Error("Unknown account"), ERR.UNKNOWN_ACCOUNT);
    if (isOAuthAccount(ctx.account.custom)) this.#primeAuth(ctx);
    await writeOof({
      account: ctx.account,
      asVersion: ctx.account.custom?.asversion ?? "14.1",
      ...settings,
    });
    return null;
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
      if (!trimmed)
        throw withCode(
          new Error("Account name is required"),
          ERR.UNKNOWN_ACCOUNT,
        );
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
        throw withCode(
          new Error("Invalid ActiveSync version selection"),
          ERR.UNKNOWN_COMMAND,
        );
      }
      customPatch.asversionselected = v;
    }
    if ("provision" in patch) {
      customPatch.provision = !!patch.provision;
    }
    if ("conflict" in patch) {
      const v = String(patch.conflict ?? "");
      if (!ALLOWED_CONFLICT_VALUES.includes(v)) {
        throw withCode(
          new Error("Invalid conflict-policy selection"),
          ERR.UNKNOWN_COMMAND,
        );
      }
      customPatch.conflict = v;
    }
    if ("contactsDisplayOverride" in patch) {
      customPatch.displayoverride = !!patch.contactsDisplayOverride;
    }
    if ("contactsNameSeparator" in patch) {
      const v = String(patch.contactsNameSeparator ?? "");
      if (!ALLOWED_NAME_SEPARATORS.includes(v)) {
        throw withCode(
          new Error("Invalid name-separator selection"),
          ERR.UNKNOWN_COMMAND,
        );
      }
      customPatch.seperator = v;
    }
    if ("calendarSyncLimit" in patch) {
      const v = String(patch.calendarSyncLimit ?? "");
      if (!ALLOWED_CALENDAR_SYNC_LIMITS.includes(v)) {
        throw withCode(
          new Error("Invalid calendar sync limit"),
          ERR.UNKNOWN_COMMAND,
        );
      }
      customPatch.synclimit = v;
    }
    if ("syncOnChange" in patch) {
      const v = String(patch.syncOnChange ?? "");
      if (!ALLOWED_SYNC_ON_CHANGE_VALUES.includes(v)) {
        throw withCode(
          new Error("Invalid sync-after-change selection"),
          ERR.UNKNOWN_COMMAND,
        );
      }
      customPatch.syncOnChange = v;
    }
    if ("syncRecurrence" in patch) {
      customPatch.syncrecurrence = !!patch.syncRecurrence;
    }
    if ("galEnabled" in patch) {
      customPatch.galenabled = !!patch.galEnabled;
    }
    if ("freeBusyEnabled" in patch) {
      customPatch.freebusyenabled = !!patch.freeBusyEnabled;
    }

    const outgoing = { ...topLevelPatch };
    if (Object.keys(customPatch).length) outgoing.custom = customPatch;
    if (Object.keys(outgoing).length) {
      await this.updateAccount({ accountId, patch: outgoing });
    }

    // Re-evaluate the GAL listener against the post-save state. Both
    // `enableGal` and `disableGal` are idempotent, so they're safe to
    // call regardless of whether the toggle actually changed.
    if ("galEnabled" in patch) {
      const fresh = await this.getAccount(accountId);
      if (patch.galEnabled && fresh?.account) {
        await enableGal({ provider: this, account: fresh.account });
      } else {
        await disableGal({ provider: this, accountId });
      }
    }

    // Answers gathered under the old setting must not survive it.
    if ("freeBusyEnabled" in patch) {
      forgetFreeBusyCache(accountId);
      await refreshFreeBusyListener(this);
    }

    // Recorded only once the save has actually gone through, so a failed
    // write is not mistaken for the user having fixed anything.
    this.#configSaves.add(accountId);
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
    return ctx?.folders.find((f) => f.folderId === folderId) ?? null;
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────

/** Who owns this account's calendars, in the `mailto:` form Thunderbird
 *  wants for `organizerId`. Empty while the mailbox address is unknown -
 *  on AS 2.5, or before the server has answered - and the calendar then
 *  keeps Thunderbird's own default. The login is deliberately not used
 *  as a stand-in: it may be `DOMAIN\user` or any other non-address, and
 *  a wrong owner is worse than none. */
function calendarOwner(account) {
  const address = account?.custom?.userSmtpAddress;
  if (!address) return {};
  const name = account?.custom?.userDisplayName;
  return {
    organizer: `mailto:${address}`,
    ...(name ? { organizerName: name } : {}),
  };
}

/** Resolve the local TB address-book / calendar name for a folder.
 *  Reuses whatever the user (or a previous bind) put in
 *  `folder.targetName`; otherwise falls back to
 *  `${displayName} (${accountName})`. */
function localNameForFolder(folder, ctx) {
  const stored = folder?.targetName?.trim?.();
  if (stored) return stored;
  return `${folder.displayName} (${ctx.account.accountName})`;
}

/** EAS requires a stable device identifier. 32 chars is the de-facto
 *  convention (some servers truncate to 32). The `MZTB` prefix marks the
 *  generator so the FriendlyName in `Settings/DeviceInformation/Set` can
 *  strip it for a cleaner label in Exchange's mobile-device list. */
function generateDeviceId() {
  const bytes = new Uint8Array(14);
  crypto.getRandomValues(bytes);
  const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, "0")).join(
    "",
  );
  return "MZTB" + hex;
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
      indexMap: [],
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
 *  re-running on an already-processed folder yields the same output.
 *
 *  Trash folders themselves (`type === "4"`) are always hidden. Folders
 *  whose parent is a Trash folder ("trash children") are hidden when
 *  `showItemsInTrash` is `false` and revealed (with a localized "Trash | "
 *  display-name prefix) when it is `true`. */
export function applyTrashVisibility(folder, trashServerIDs, showItemsInTrash) {
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
    hidden: isTrash || (inTrash && !showItemsInTrash),
    targetType: isTrash ? null : folder.targetType,
    custom: { ...c, displayNameRaw: raw },
  };
}

async function readShowItemsInTrash() {
  try {
    const rv = await browser.storage.local.get({ showItemsInTrash: false });
    return rv.showItemsInTrash === true;
  } catch (err) {
    console.debug("[eas] readShowItemsInTrash storage.get failed:", err);
    return false;
  }
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
async function mergeFolderDeltas(existingFolders, delta) {
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
    const rawName =
      upd.displayName ||
      existing.custom?.displayNameRaw ||
      existing.displayName;
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
      const rawName =
        add.displayName ||
        existing.custom?.displayNameRaw ||
        existing.displayName;
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
 *  the runtime sync path (initial + delta) and the migration upgrade.
 *  Reads the `showItemsInTrash` advanced option once so a runtime toggle
 *  is reflected on the next folder-list fetch. */
export async function finalizeFolderListForPush(folders) {
  const showItemsInTrash = await readShowItemsInTrash();
  const trashServerIDs = buildTrashServerIDs(folders);
  return folders
    .map((f) => applyTrashVisibility(f, trashServerIDs, showItemsInTrash))
    .map((f) => ({
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
        indexMap: f.custom?.indexMap ?? [],
        displayNameRaw: f.custom?.displayNameRaw,
      },
    }));
}
