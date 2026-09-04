/**
 * Unit tests for the external-browser sign-in route.
 *
 * Thunderbird's internal OAuth window has no WebAuthn plumbing, so a tenant
 * that requires a passkey can never finish consent there - the ceremony
 * waits for a prompt Thunderbird's chrome never shows. Thunderbird's own
 * accounts escape through `mailnews.oauth.useExternalBrowser`; that pref
 * belongs to `OAuth2.sys.mjs` and never reaches this add-on, which drives
 * its own popup. The `oauth.useExternalBrowser` option is our equivalent:
 * hand the authorization URL to the system browser, and take the redirect
 * URL back by paste, because a foreign browser is not something
 * `tabs.onUpdated` can watch.
 *
 * What is pinned here is the handoff - which window opens where, and which
 * pasted URLs are allowed to end it. Every test carries a timeout: the
 * failure mode of a broken handoff is a sign-in that never settles, and an
 * assertion nobody reaches is not a test result.
 *
 * Run with `npm run test:unit`.
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

const oauth = await import("../../src/modules/eas/oauth.mjs");
const { startAuth, completeExternalConsent, reopenExternalConsent } = oauth;

const T = { timeout: 5000 };

const REDIRECT_URI =
  "https://login.microsoftonline.com/common/oauth2/nativeclient";
const AUTHORIZE =
  "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";

/** The host APIs the consent flow reaches for, recording what it did.
 *  Built per test, so nothing carries over. The waiters exist because the
 *  flow is asynchronous in ways a fixed number of microtask turns cannot
 *  be trusted to cover. */
function installHostEnv({ useExternalBrowser = false } = {}) {
  const state = { externalUrls: [], windows: [], removed: [], nextId: 1 };
  const removalListeners = new Set();
  const updateListeners = new Set();
  const windowWatchers = [];
  const externalWatchers = [];

  // Resolved from a macrotask, not a microtask: the flow registers its
  // listeners in the turns right after the call the waiter is watching, and
  // a test that woke between the two would drive a window nobody is
  // listening to yet.
  const settle = (watchers, list) => {
    for (const w of watchers.splice(0)) {
      if (list.length >= w.n) setTimeout(() => w.resolve(list[w.n - 1]), 0);
      else watchers.push(w);
    }
  };

  globalThis.browser.storage = {
    local: {
      async get(defaults) {
        const out = { ...defaults };
        if ("oauth.useExternalBrowser" in out) {
          out["oauth.useExternalBrowser"] = useExternalBrowser;
        }
        return out;
      },
    },
  };

  globalThis.browser.windows = {
    async openDefaultBrowser(url) {
      state.externalUrls.push(url);
      settle(externalWatchers, state.externalUrls);
    },
    async create(args) {
      const id = state.nextId++;
      state.windows.push({ id, ...args });
      settle(windowWatchers, state.windows);
      return { id };
    },
    async remove(id) {
      state.removed.push(id);
      // The real API fires onRemoved for a window it closed itself, and the
      // flow has to survive that re-entry rather than read it as a cancel.
      for (const fn of [...removalListeners]) fn(id);
    },
    onRemoved: {
      addListener: (fn) => removalListeners.add(fn),
      removeListener: (fn) => removalListeners.delete(fn),
    },
  };

  globalThis.browser.tabs = {
    onUpdated: {
      addListener: (fn) => updateListeners.add(fn),
      removeListener: (fn) => updateListeners.delete(fn),
    },
  };

  /** Resolves with the nth window once it has been created. */
  state.windowOpened = (n = 1) =>
    new Promise((resolve) => {
      if (state.windows.length >= n) return resolve(state.windows[n - 1]);
      windowWatchers.push({ n, resolve });
    });
  /** Resolves with the nth URL once it has reached the system browser. */
  state.externalOpened = (n = 1) =>
    new Promise((resolve) => {
      if (state.externalUrls.length >= n)
        return resolve(state.externalUrls[n - 1]);
      externalWatchers.push({ n, resolve });
    });
  /** Drive the internal popup the way Thunderbird would. */
  state.navigateInternalPopup = (windowId, url) => {
    for (const fn of [...updateListeners]) fn(1, { url }, { windowId, url });
  };
  /** Close a window from the outside - the user clicking the X. */
  state.closeWindow = (windowId) => {
    state.removed.push(windowId);
    for (const fn of [...removalListeners]) fn(windowId);
  };

  return state;
}

/** A token endpoint that answers the authorization-code exchange. */
function installTokenEndpoint() {
  const seen = [];
  globalThis.fetch = async (_url, init) => {
    seen.push(new URLSearchParams(init?.body ?? ""));
    return {
      ok: true,
      status: 200,
      text: async () =>
        JSON.stringify({
          access_token: "at-1",
          refresh_token: "rt-1",
          expires_in: 3600,
        }),
    };
  };
  return seen;
}

/** The `state` the flow put in the authorization URL. A URL without it is
 *  from a different sign-in attempt, which is what the parameter is for. */
const stateOf = (authUrl) => new URL(authUrl).searchParams.get("state");

/** The paste dialog's handoff token, which it reads from its own query
 *  string and quotes back on every message. */
const tokenOf = (dialogWindow) =>
  new URL(dialogWindow.url, "https://example.invalid/").searchParams.get(
    "token",
  );

/** The whole external handoff up to the point the user is pasting: returns
 *  the live sign-in, the dialog's token, and the URL Microsoft will land on. */
async function startExternalSignIn(host, extra = {}) {
  const signIn = startAuth({ servertype: "office365", ...extra });
  const dialog = await host.windowOpened(1);
  const authUrl = await host.externalOpened(1);
  return {
    signIn,
    dialog,
    token: tokenOf(dialog),
    landedUrl: `${REDIRECT_URI}?code=abc&state=${stateOf(authUrl)}`,
  };
}

test("the option off keeps consent inside Thunderbird", T, async () => {
  const host = installHostEnv({ useExternalBrowser: false });
  installTokenEndpoint();

  const signIn = startAuth({ servertype: "office365" });
  const popup = await host.windowOpened(1);

  assert.deepEqual(host.externalUrls, [], "the system browser was not used");
  assert.ok(
    popup.url.startsWith(AUTHORIZE),
    "the one window that opened is the authorization page itself",
  );

  host.navigateInternalPopup(
    popup.id,
    `${REDIRECT_URI}?code=abc&state=${stateOf(popup.url)}`,
  );
  assert.equal((await signIn).refreshToken, "rt-1");
});

test(
  "the option on sends the authorization page to the system browser",
  T,
  async () => {
    const host = installHostEnv({ useExternalBrowser: true });
    installTokenEndpoint();

    const { signIn, dialog, token, landedUrl } =
      await startExternalSignIn(host);

    assert.ok(
      host.externalUrls[0].startsWith(AUTHORIZE),
      "the browser was handed the authorization page",
    );
    assert.equal(host.windows.length, 1, "one Thunderbird window opened");
    assert.ok(
      dialog.url.includes("dialogs/oauth-paste/oauth-paste.html"),
      "and it is the paste dialog, not the authorization page",
    );

    assert.deepEqual(await completeExternalConsent({ token, url: landedUrl }), {
      accepted: true,
    });
    assert.equal((await signIn).refreshToken, "rt-1");
    assert.deepEqual(
      host.removed,
      [dialog.id],
      "the paste dialog was closed once it had served its purpose",
    );
  },
);

test(
  "a URL from somewhere else is refused without ending the sign-in",
  T,
  async () => {
    const host = installHostEnv({ useExternalBrowser: true });
    installTokenEndpoint();
    const { signIn, token, landedUrl } = await startExternalSignIn(host);

    assert.deepEqual(
      await completeExternalConsent({ token, url: "not a url at all" }),
      { accepted: false, reason: "notAUrl" },
    );
    assert.deepEqual(
      await completeExternalConsent({
        token,
        url: "https://example.com/?code=x",
      }),
      { accepted: false, reason: "notRedirect" },
    );
    assert.deepEqual(
      await completeExternalConsent({ token, url: `${REDIRECT_URI}?foo=bar` }),
      { accepted: false, reason: "noCode" },
    );
    assert.deepEqual(
      await completeExternalConsent({
        token,
        url: `${REDIRECT_URI}?code=abc&state=from-an-older-attempt`,
      }),
      { accepted: false, reason: "staleLink" },
    );

    assert.deepEqual(
      host.removed,
      [],
      "the dialog is still up for another try",
    );

    // And the sign-in is genuinely still live: the right URL still lands.
    await completeExternalConsent({ token, url: landedUrl });
    assert.equal((await signIn).refreshToken, "rt-1");
  },
);

test("Microsoft's own error rides back through the paste", T, async () => {
  const host = installHostEnv({ useExternalBrowser: true });
  installTokenEndpoint();
  const { signIn, token } = await startExternalSignIn(host);

  const verdict = await completeExternalConsent({
    token,
    url:
      `${REDIRECT_URI}?error=access_denied&error_description=Admin+consent+required` +
      `&state=${stateOf(host.externalUrls[0])}`,
  });
  assert.deepEqual(verdict, { accepted: true }, "the paste itself was fine");
  await assert.rejects(() => signIn, /access_denied/);
});

test("closing the paste dialog cancels the sign-in", T, async () => {
  const host = installHostEnv({ useExternalBrowser: true });
  installTokenEndpoint();
  const { signIn, dialog } = await startExternalSignIn(host);

  host.closeWindow(dialog.id);
  await assert.rejects(
    () => signIn,
    (err) => err.code === "E:CANCELLED",
  );
});

test("a token nobody is waiting on is refused, not thrown at", T, async () => {
  installHostEnv({ useExternalBrowser: true });
  assert.deepEqual(
    await completeExternalConsent({
      token: "never-issued",
      url: `${REDIRECT_URI}?code=abc`,
    }),
    { accepted: false, reason: "expired" },
  );
  assert.deepEqual(await reopenExternalConsent({ token: "never-issued" }), {
    accepted: false,
    reason: "expired",
  });
});

test(
  "reopen hands the browser the same URL it was given the first time",
  T,
  async () => {
    const host = installHostEnv({ useExternalBrowser: true });
    installTokenEndpoint();
    const { signIn, dialog, token } = await startExternalSignIn(host);

    assert.deepEqual(await reopenExternalConsent({ token }), {
      accepted: true,
    });
    assert.equal(host.externalUrls.length, 2, "the browser was asked again");
    assert.equal(
      host.externalUrls[1],
      host.externalUrls[0],
      "with the same authorization URL - a fresh one would carry a state that makes the paste the user is about to do stale",
    );

    host.closeWindow(dialog.id);
    await assert.rejects(() => signIn);
  },
);

test(
  "the paste dialog is the window a caller is told to track",
  T,
  async () => {
    // The provider registers this id so a second Connect click raises the
    // window already up rather than starting a competing sign-in. On the
    // external route there is no consent popup to register - the dialog is
    // the only thing Thunderbird owns.
    const host = installHostEnv({ useExternalBrowser: true });
    installTokenEndpoint();

    const seen = [];
    const { signIn, dialog } = await startExternalSignIn(host, {
      onWindowCreated: (id) => seen.push(id),
    });

    assert.deepEqual(seen, [dialog.id]);
    host.closeWindow(dialog.id);
    await assert.rejects(() => signIn);
  },
);
