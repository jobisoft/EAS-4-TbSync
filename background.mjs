import { EasProvider } from "./modules/eas-provider.mjs";
import { startAuth } from "./modules/eas/oauth.mjs";
import { discoverEasServer } from "./modules/eas/autodiscover.mjs";
import { runUpgrades, enqueueUpgradesForUpdate } from "./modules/upgrades.mjs";
import { installAnchorMailboxInjector } from "./modules/anchor-mailbox.mjs";

/**
 * Provider entry point. All port / handshake plumbing lives inside the
 * TbSyncProviderImplementation base class - this file constructs the
 * concrete EasProvider, wires internal runtime messages from the UI
 * dialogs, and calls init().
 *
 * The provider carries no persistent storage: the host owns the account
 * and folder rows (including server URL, username, password-as-account-
 * custom, sync keys, and the changelog). The host also runs the address-
 * book observer; the provider is a pure consumer of the host's changelog
 * queue for contact sync.
 */

// Register the anchor-mailbox webRequest listener before the provider
// constructs and starts issuing requests, so the very first OPTIONS /
// FolderSync of the boot is already cookie-injected.
installAnchorMailboxInjector();

const provider = new EasProvider();

/** Resolves once the provider has finished its boot sequence (instance
 *  constructed, `init()` returned, host port open). The upgrade runner
 *  awaits this before issuing any host RPC. */
export const providerReady = (async () => {
  await new Promise(resolve => provider.onceConnectedToHost(resolve));
})();

// Internal messages from our own UI pages (setup.html, config.html).
// Errors are returned as structured { ok, error, code } rather than thrown,
// because runtime.sendMessage serialisation drops Error.code and the dialogs
// need the code to distinguish user-cancel from real failures.
browser.runtime.onMessage.addListener(async msg => {
  if (msg?.type === "eas.startOAuth") {
    try {
      const result = await startAuth({
        loginHint: msg.loginHint,
        servertype: msg.servertype,
      });
      return { ok: true, result };
    } catch (err) {
      return { ok: false, error: err.message ?? String(err), code: err.code ?? null };
    }
  }
  if (msg?.type === "eas.discoverServer") {
    try {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), 30_000);
      try {
        const result = await discoverEasServer({
          email: msg.email,
          password: msg.password,
          signal: controller.signal,
        });
        return { ok: true, result };
      } finally {
        clearTimeout(timer);
      }
    } catch (err) {
      return {
        ok: false,
        error: err.message ?? String(err),
        code: err.code ?? null,
        details: err.details ?? null,
      };
    }
  }
  if (msg?.type === "eas.createAccount") {
    try {
      // Forward the whole message minus `type` so the provider can branch
      // on `method` and read both basic-auth and OAuth-specific fields.
      const { type: _t, ...args } = msg;
      const result = await provider.createAccountFromSetup(args);
      return { ok: true, result };
    } catch (err) {
      return { ok: false, error: err.message ?? String(err), code: err.code ?? null };
    }
  }
  if (msg?.type === "eas.getAccount") {
    try {
      const result = await provider.getAccountForConfig(msg.accountId);
      return { ok: true, result };
    } catch (err) {
      return { ok: false, error: err.message ?? String(err), code: err.code ?? null };
    }
  }
  if (msg?.type === "eas.saveAccount") {
    try {
      const result = await provider.saveAccountFromConfig({
        accountId: msg.accountId,
        patch: msg.patch ?? {},
      });
      return { ok: true, result };
    } catch (err) {
      return { ok: false, error: err.message ?? String(err), code: err.code ?? null };
    }
  }
  return undefined;
});

provider.init();

// ── One-shot upgrade runner ──────────────────────────────────────────────
//
// `runtime.onInstalled` enqueues the IDs of every upgrade whose split
// version falls in `(previousVersion, currentVersion]`, then drains them
// once the provider is connected to the host. Fresh installs short-circuit
// at the reason check so no upgrade ever runs on a clean profile.
browser.runtime.onInstalled.addListener(async (details) => {
  if (details.reason !== "update" || !details.previousVersion) return;
  const cur = browser.runtime.getManifest().version;
  const enqueued = await enqueueUpgradesForUpdate(details.previousVersion, cur);
  if (!enqueued) return;
  await providerReady;
  await runUpgrades(provider);
});

// Boot-time stale-queue drain. `runtime.onInstalled` only fires on the
// boot where the install/update *actually* happened - if a previous run
// failed mid-flight, the queue persists in storage and we need a second,
// independent trigger to retry. `runUpgrades` is idempotent + self-
// coalescing, so a same-boot collision with the listener above is safe.
providerReady.then(() => runUpgrades(provider));
