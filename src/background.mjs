import { EasProvider } from "./modules/eas-provider.mjs";
import { startAuth } from "./modules/eas/oauth.mjs";
import { discoverEasServer } from "./modules/eas/autodiscover.mjs";
import { installAnchorMailboxInjector } from "./modules/anchor-mailbox.mjs";
import {
  registerCalendarProvider,
  setSyncHandlers,
} from "./modules/calendar-provider.mjs";

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

// Claim our calendar type before anything can ask for one. Calendars of a
// type nobody has registered come up as force-disabled placeholders, so the
// listeners have to exist by the time the calendar manager reads the
// profile - not once the first sync runs.
registerCalendarProvider();

const provider = new EasProvider();

// The hooks are handed user edits but own no storage, so give them the way
// back to the host: the changelog they record into, and the event log that
// makes a recorded edit visible next to every other sync decision.
setSyncHandlers({
  provider,
  reportEventLog: (payload) => provider.reportEventLog(payload),
});

// Internal messages from our own UI pages (setup.html, config.html).
//
// The listener is deliberately NOT async. Returning a promise from an
// onMessage listener claims the message and supplies its response, so an
// async listener would answer every message in this add-on - including
// `tbsync-setup-completed`, which belongs to the base class's own listener.
// Returning nothing for anything not in this table leaves those alone.
//
// Errors come back as structured { ok, error, code } rather than thrown,
// because runtime.sendMessage serialisation drops Error.code and the dialogs
// need the code to distinguish user-cancel from real failures.
const MESSAGE_HANDLERS = {
  "eas.startOAuth": (msg) =>
    startAuth({ loginHint: msg.loginHint, servertype: msg.servertype }),

  "eas.discoverServer": async (msg) => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 30_000);
    try {
      return await discoverEasServer({
        email: msg.email,
        password: msg.password,
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timer);
    }
  },

  // Forward the whole message minus `type` so the provider can branch on
  // `method` and read both basic-auth and OAuth-specific fields.
  "eas.createAccount": ({ type: _t, ...args }) =>
    provider.createAccountFromSetup(args),

  "eas.getAccount": (msg) => provider.getAccountForConfig(msg.accountId),

  "eas.saveAccount": (msg) =>
    provider.saveAccountFromConfig({
      accountId: msg.accountId,
      patch: msg.patch ?? {},
    }),
};

/** Run a handler and shape the reply the dialogs expect. `details` rides
 *  along because the setup dialog reads `reply.details.tried` from a failed
 *  Autodiscover; it is null for errors that carry none. */
async function replyEnvelope(handler, msg) {
  try {
    return { ok: true, result: await handler(msg) };
  } catch (err) {
    return {
      ok: false,
      error: err?.message ?? String(err),
      code: err?.code ?? null,
      details: err?.details ?? null,
    };
  }
}

browser.runtime.onMessage.addListener((msg) => {
  const handler = MESSAGE_HANDLERS[msg?.type];
  // Not ours - stay out of the way so the listener it belongs to can answer.
  if (!handler) return;
  return replyEnvelope(handler, msg);
});

provider.init();
