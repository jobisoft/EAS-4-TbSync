/**
 * The Thunderbird half of the external-browser sign-in
 * (oauth.mjs::runExternalConsent). The authorization page is up in the
 * system browser, where a passkey works; this window is how its answer
 * gets back, because `tabs.onUpdated` cannot watch another application.
 *
 * The window carries a handoff token in its query string and quotes it on
 * every message, so a dialog left over from an abandoned attempt cannot
 * answer for the sign-in that is live now.
 *
 * Whether a pasted URL is good enough to end the sign-in is not decided
 * here - the background decides, and answers with a reason this file turns
 * into a sentence. A refusal leaves the window up for another try; on
 * acceptance the background closes it.
 */

import { localizeDocument } from "../../vendor/i18n/i18n.mjs";

const $ = (id) => document.getElementById(id);
const i18n = (key, fallback) => browser.i18n.getMessage(key) || fallback;

const token = new URLSearchParams(location.search).get("token");

/** One sentence per way the paste can be wrong. `expired` is the odd one:
 *  the sign-in is gone, so there is nothing to try again and the window is
 *  only waiting to be closed. */
const REASONS = {
  notAUrl: [
    "oauthPaste.error.notAUrl",
    "That does not look like a web address. Copy the whole address from your browser's address bar.",
  ],
  notRedirect: [
    "oauthPaste.error.notRedirect",
    "That is not the page sign-in ends on. Finish signing in, wait for the blank page, and copy the address from there.",
  ],
  noCode: [
    "oauthPaste.error.noCode",
    "That address carries no sign-in result. It was probably copied before sign-in finished.",
  ],
  staleLink: [
    "oauthPaste.error.staleLink",
    "That address is from an earlier sign-in attempt. Use “Open the sign-in page again”, then copy the new address.",
  ],
  expired: [
    "oauthPaste.error.expired",
    "This sign-in is no longer waiting for an answer. Close this window and start again from TbSync.",
  ],
  noBrowser: [
    "oauthPaste.error.noBrowser",
    "Could not open your default browser.",
  ],
};

function showError(text) {
  const el = $("error");
  el.textContent = text;
  el.classList.add("visible");
}

function clearError() {
  $("error").classList.remove("visible");
}

/** Send one message and unwrap the background's envelope. Returns the
 *  verdict, or null when the message itself failed - the reply envelope
 *  carries that as `ok: false`. */
async function ask(type) {
  const reply = await browser.runtime.sendMessage({
    type,
    token,
    url: $("redirect-url").value,
  });
  if (!reply?.ok) {
    showError(
      reply?.error ??
        i18n("oauthPaste.error.generic", "Could not reach the add-on."),
    );
    return null;
  }
  return reply.result;
}

async function submit() {
  clearError();
  if (!$("redirect-url").value.trim()) {
    showError(
      i18n(
        "oauthPaste.error.empty",
        "Paste the address your browser ended up on.",
      ),
    );
    $("redirect-url").focus();
    return;
  }

  $("btn-continue").disabled = true;
  try {
    const verdict = await ask("eas.completeExternalOAuth");
    if (!verdict) return;
    if (verdict.accepted) {
      // The sign-in has the URL and this window is on its way out; keep the
      // form quiet for the moment it is still on screen.
      $("redirect-url").disabled = true;
      return;
    }
    const [key, fallback] = REASONS[verdict.reason] ?? [
      "oauthPaste.error.generic",
      "That address could not be used.",
    ];
    showError(i18n(key, fallback));
  } finally {
    // Not in the accepted branch's interest, but harmless there and the
    // only way a refused paste gets a second attempt.
    if (!$("redirect-url").disabled) $("btn-continue").disabled = false;
  }
}

async function reopen() {
  clearError();
  const verdict = await ask("eas.reopenExternalOAuth");
  if (!verdict || verdict.accepted) return;
  const [key, fallback] = REASONS[verdict.reason] ?? [
    "oauthPaste.error.generic",
    "Could not open your default browser.",
  ];
  showError(i18n(key, fallback));
}

document.addEventListener("DOMContentLoaded", () => {
  localizeDocument();

  if (!token) {
    // Nothing here can work without one, and there is no sign-in to cancel.
    showError(
      i18n(
        "oauthPaste.error.missingToken",
        "Missing handoff token. Open this window through TbSync.",
      ),
    );
    $("btn-continue").disabled = true;
    $("btn-reopen").disabled = true;
    return;
  }

  $("btn-continue").addEventListener("click", submit);
  $("btn-reopen").addEventListener("click", reopen);
  // Cancel tells the background, which closes this window: a content page
  // is not reliably allowed to close the window it lives in, and the
  // sign-in has to be told either way.
  $("btn-cancel").addEventListener("click", () =>
    ask("eas.cancelExternalOAuth"),
  );

  // A URL pasted into the box is the whole interaction; Enter should finish
  // it. Shift+Enter stays a newline, since the box is a textarea only so a
  // long URL is readable.
  $("redirect-url").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      submit();
    }
  });

  $("redirect-url").focus();
});
