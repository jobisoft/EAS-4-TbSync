/**
 * Provider-local one-shot upgrades.
 *
 * Runs work that has to happen exactly once after the user updates the
 * provider across a "split version" - typically a one-time data-shape
 * migration that the host's legacy migration deliberately couldn't do
 * because it's provider-specific.
 *
 * The trigger is `runtime.onInstalled` (with `reason === "update"` and a
 * `previousVersion` set), wired up in [background.mjs](../background.mjs).
 * Fresh installs never fire any upgrade. The list of pending upgrade IDs
 * persists in `browser.storage.local` under UPGRADE_QUEUE_KEY so a
 * partial run (host crash, network outage) is retried on the next
 * host-connect via the boot-time stale drain.
 *
 * While a drain is in flight, the host treats every account belonging
 * to this provider as "upgrading" - refuses every user-initiated RPC
 * and skips autosync ticks. The lock is acquired before the first
 * upgrade body runs and released in a `finally` so a crashing upgrade
 * still releases it.
 */

import { finalizeFolderListForPush } from "./eas-provider.mjs";

const UPGRADE_QUEUE_KEY = "eas.upgradeQueue";

/** Ordered list of split versions. An upgrade is *applicable* to an
 *  `(previousVersion, currentVersion)` pair iff
 *  `previousVersion < splitVersion <= currentVersion`. Strict on the
 *  prev side so a user already on the split doesn't re-run; inclusive
 *  on the cur side so installing exactly at the split triggers it. */
export const UPGRADES = [
  {
    splitVersion: "4.20",
    id: "eas.host-and-https-to-server",
    run: async (provider) => {
      await liftLegacyAccountState(provider);
    },
  },
];

/** Legacy add-on root pref branch ([provider.js:28](../../EAS-4-TbSync/content/provider.js#L28)).
 *  Used to find the global OAuth client-ID slot. */
const LEGACY_PREF_CLIENT_ID = "extensions.eas4tbsync.oauth.clientID";

/** Storage.local key the new add-on reads at refresh time
 *  ([modules/eas/oauth.mjs](eas/oauth.mjs)). One value shared by all
 *  accounts in the profile, matching legacy semantics. */
const STORAGE_KEY_CLIENT_ID = "tbsync.clientID";

/** nsILoginManager realm legacy used for EAS credentials
 *  ([content/includes/network.js:122](../../EAS-4-TbSync/content/includes/network.js#L122)).
 *  The origin is namespaced per-account as "TbSync#<accountID>" rather
 *  than the actual server hostname - legacy decoupled the credential
 *  from the host so Autodiscover-driven host changes don't orphan it
 *  ([content/includes/network.js:107-115](../../EAS-4-TbSync/content/includes/network.js#L107-L115)).
 *  The legacy accountID survives the host's profile migration unchanged,
 *  so we can reach the entry by reusing `account.accountId` here. */
const LEGACY_LOGIN_REALM = "TbSync/EAS";

function legacyLoginOrigin(accountId) {
  return `TbSync#${accountId}`;
}

/** Dotted-decimal version comparison. Sufficient for the version
 *  strings the legacy add-on shipped (e.g. `"4.17.2.ews.16.1"`) and
 *  the new add-on (`"5.0"`) - any non-numeric segment becomes NaN,
 *  which only matters if the *first differing* segment is non-numeric.
 *  In practice the legacy → new transition diverges at the first
 *  segment (4 → 5), so the comparator short-circuits before reaching
 *  any non-numeric tail. */
export function compareVersions(a, b) {
  const pa = String(a).split(".").map(Number);
  const pb = String(b).split(".").map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const x = pa[i] ?? 0, y = pb[i] ?? 0;
    if (x !== y) return x - y;
  }
  return 0;
}

let inFlight = null;

/** Drain `eas.upgradeQueue` against the UPGRADES table. Idempotent
 *  (each upgrade body is itself idempotent) and self-coalescing - a
 *  second caller while the first is mid-flight just awaits the same
 *  Promise.
 *
 *  The host upgrade lock is acquired before any upgrade body runs and
 *  released in `finally`, so:
 *    - User-initiated RPCs against this provider's accounts are refused
 *      while the drain is running.
 *    - Autosync ticks skip those accounts.
 *    - A throw inside an upgrade still releases the lock; the failed
 *      upgrade ID stays in the queue and is retried next boot. */
export function runUpgrades(provider) {
  if (inFlight) return inFlight;
  inFlight = (async () => {
    const rv = await browser.storage.local.get({ [UPGRADE_QUEUE_KEY]: [] });
    const queue = rv[UPGRADE_QUEUE_KEY];
    if (!queue.length) return;

    let lockAcquired = false;
    try {
      await provider.setProviderUpgradeLock(true);
      lockAcquired = true;
      provider.reportEventLog({
        level: "debug",
        message: `[upgrade] entering upgrade mode - sync and account/resource modifications are paused (${queue.length} upgrade(s) pending)`,
      });

      const remaining = [];
      for (const id of queue) {
        const upgrade = UPGRADES.find(u => u.id === id);
        if (!upgrade) continue;  // unknown id - silently drop
        try {
          provider.reportEventLog({ level: "debug", message: `[upgrade] ${id} starting` });
          await upgrade.run(provider);
          provider.reportEventLog({ level: "debug", message: `[upgrade] ${id} done` });
        } catch (err) {
          provider.reportEventLog({
            level: "error",
            message: `[upgrade] ${id} failed: ${err?.message ?? String(err)}`,
          });
          remaining.push(id);
        }
      }

      await browser.storage.local.set({ [UPGRADE_QUEUE_KEY]: remaining });
    } finally {
      if (lockAcquired) {
        await provider.setProviderUpgradeLock(false).catch(err =>
          console.warn("[eas-4-tbsync] failed to release upgrade lock:", err?.message ?? String(err))
        );
        provider.reportEventLog({
          level: "debug",
          message: `[upgrade] exiting upgrade mode - sync and account/resource modifications re-enabled`,
        });
      }
      inFlight = null;
    }
  })();
  return inFlight;
}

/** Compute the set of upgrades triggered by an update transition and
 *  merge their IDs into the persistent queue. No-op when nothing
 *  applies. Returns the new queue length. */
export async function enqueueUpgradesForUpdate(previousVersion, currentVersion) {
  const triggered = UPGRADES
    .filter(u =>
      compareVersions(previousVersion, u.splitVersion) < 0
      && compareVersions(u.splitVersion, currentVersion) <= 0
    )
    .map(u => u.id);
  if (!triggered.length) return 0;
  const rv = await browser.storage.local.get({ [UPGRADE_QUEUE_KEY]: [] });
  const next = Array.from(new Set([...rv[UPGRADE_QUEUE_KEY], ...triggered]));
  await browser.storage.local.set({ [UPGRADE_QUEUE_KEY]: next });
  return next.length;
}

// ── Upgrade bodies ───────────────────────────────────────────────────────

/** Lift the legacy account state into the shape the new EAS provider
 *  reads:
 *
 *    1. Global OAuth client-ID pref (`extensions.eas4tbsync.oauth.clientID`)
 *       → `browser.storage.local["tbsync.clientID"]`. Skipped when the
 *       legacy pref is missing or empty.
 *    2. Per-account `host` + `https` (bool) → `custom.server` URL.
 *       Strips trailing slashes and appends `/Microsoft-Server-ActiveSync`
 *       unless the host already ends in it.
 *    3. Per-account credentials lifted from `nsILoginManager` (origin =
 *       `TbSync#<accountID>`, realm = "TbSync/EAS", user = legacy user)
 *       into `custom`. Branches on `custom.servertype`:
 *         - "office365" → parse the JSON token blob, take `.refresh`,
 *           write `refreshToken`.
 *         - anything else → write `password` from the raw login-manager
 *           value.
 *
 *  Each step is idempotent. */
async function liftLegacyAccountState(provider) {
  await liftClientIDPref(provider);

  const accounts = await provider.listAccounts();
  for (const acc of accounts) {
    try {
      await liftHostAndHttpsToServer(provider, acc);
      await liftCredentials(provider, acc);
      await liftFolderVisibility(provider, acc);
    } catch (err) {
      provider.reportEventLog({
        level: "warning",
        accountId: acc.accountId,
        message: `[upgrade] failed to lift legacy state: ${err?.message ?? String(err)}`,
      });
    }
  }
}

async function liftFolderVisibility(provider, acc) {
  const rv = await provider.getAccount(acc.accountId);
  const folders = rv?.folders ?? [];
  if (!folders.length) return;
  const patched = finalizeFolderListForPush(folders);
  await provider.pushFolderList({ accountId: acc.accountId, folders: patched });
  provider.reportEventLog({
    level: "debug",
    accountId: acc.accountId,
    message: `[upgrade] applied trash visibility to ${patched.length} folder(s)`,
  });
}

async function liftClientIDPref(provider) {
  const existing = await browser.storage.local.get({ [STORAGE_KEY_CLIENT_ID]: "" });
  if (existing[STORAGE_KEY_CLIENT_ID]) return;
  const value = await browser.LegacyPrefs.getUserPref(LEGACY_PREF_CLIENT_ID);
  if (typeof value !== "string" || !value.trim()) return;
  await browser.storage.local.set({ [STORAGE_KEY_CLIENT_ID]: value.trim() });
  provider.reportEventLog({
    level: "debug",
    message: `[upgrade] lifted legacy '${LEGACY_PREF_CLIENT_ID}' pref into storage.local['${STORAGE_KEY_CLIENT_ID}']`,
  });
}

async function liftHostAndHttpsToServer(provider, acc) {
  if (acc.custom?.server) return;
  const host = acc.custom?.host;
  if (!host) return;
  const protocol = acc.custom?.https ? "https://" : "http://";
  let url = protocol + host;
  while (url.endsWith("/")) url = url.slice(0, -1);
  if (!url.endsWith("Microsoft-Server-ActiveSync")) url += "/Microsoft-Server-ActiveSync";
  await provider.updateAccount({
    accountId: acc.accountId,
    patch: { custom: { server: url } },
  });
  provider.reportEventLog({
    level: "debug",
    accountId: acc.accountId,
    message: `[upgrade] lifted legacy host+https to server="${url}"`,
  });
}

async function liftCredentials(provider, acc) {
  const c = acc.custom ?? {};
  const isOAuthLegacy = c.servertype === "office365";

  if (isOAuthLegacy && c.refreshToken) return;
  if (!isOAuthLegacy && c.password) return;

  const user = c.user;
  if (!user) {
    provider.reportEventLog({
      level: "warning",
      accountId: acc.accountId,
      message: `[upgrade] cannot lift credentials: missing legacy user`,
    });
    return;
  }

  const origin = legacyLoginOrigin(acc.accountId);
  const stored = await browser.LegacyLoginManager.getLoginInfo({
    origin,
    httpRealm: LEGACY_LOGIN_REALM,
    username: user,
  });
  if (stored == null) {
    provider.reportEventLog({
      level: "warning",
      accountId: acc.accountId,
      message: `[upgrade] no legacy nsILoginManager entry for (${origin}, ${LEGACY_LOGIN_REALM}, ${user})`,
    });
    return;
  }

  if (isOAuthLegacy) {
    let refreshToken = "";
    try { refreshToken = JSON.parse(stored)?.refresh ?? ""; }
    catch { /* malformed blob; refreshToken stays empty */ }
    if (!refreshToken) {
      provider.reportEventLog({
        level: "warning",
        accountId: acc.accountId,
        message: `[upgrade] legacy OAuth token blob has no 'refresh' field`,
      });
      return;
    }
    await provider.updateAccount({
      accountId: acc.accountId,
      patch: { custom: {
        refreshToken,
        authenticatedUserEmail: c.authenticatedUserEmail ?? null,
      }},
    });
    provider.reportEventLog({
      level: "debug",
      accountId: acc.accountId,
      message: `[upgrade] lifted legacy OAuth refresh token from nsILoginManager`,
    });
    return;
  }

  await provider.updateAccount({
    accountId: acc.accountId,
    patch: { custom: {
      password: stored,
    }},
  });
  provider.reportEventLog({
    level: "debug",
    accountId: acc.accountId,
    message: `[upgrade] lifted legacy basic-auth password from nsILoginManager`,
  });
}
