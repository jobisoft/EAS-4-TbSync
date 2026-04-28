/**
 * Setup popup. User picks an account type from a rich dropdown:
 *
 *   - Office 365 Business    → OAuth, host = outlook.office365.com
 *   - Personal Microsoft     → OAuth, host = eas.outlook.com
 *   - Auto-detect            → basic auth, server resolved via Autodiscover
 *   - Custom EAS server      → basic auth, user enters server URL
 *
 * For OAuth: ask the background to run consent in a separate popup
 * (oauth.mjs::runConsentPopup), receive the token bundle, then call
 * `eas.createAccount` and post `tbsync-setup-completed`.
 *
 * For auto-detect: collect email + password, run `eas.discoverServer`,
 * then call `eas.createAccount` with the resolved server URL. Errors
 * surface inline via the `auto-error-popover`.
 *
 * For custom basic auth: collect fields and call `eas.createAccount` directly.
 */

import { localizeDocument } from "../../vendor/i18n/i18n.mjs";
import { createDropdown } from "../shared/dropdown.mjs";

const i18n = (key, fallback, substitutions) =>
  browser.i18n.getMessage(key, substitutions) || fallback;

const params = new URLSearchParams(location.search);
const setupToken = params.get("setupToken");

const TYPE_OFFICE365   = "office365";
const TYPE_PERSONAL_MS = "personal-ms";
const TYPE_AUTO        = "auto";
const TYPE_CUSTOM      = "custom";

function $(id) { return document.getElementById(id); }
function val(id) { return $(id).value.trim(); }

function showError(message) {
  const el = $("error");
  el.textContent = message;
  el.classList.add("visible");
}
function clearError() { $("error").classList.remove("visible"); }

function applyType(type) {
  const isOAuth = type === TYPE_OFFICE365 || type === TYPE_PERSONAL_MS;
  const isAuto  = type === TYPE_AUTO;
  $("oauth-panel").hidden = !isOAuth;
  $("auto-panel").hidden  = !isAuto;
  $("basic-panel").hidden = isOAuth || isAuto;
  $("btn-submit").textContent = isOAuth
    ? i18n("setup.oauth.signIn", "Sign in with Microsoft")
    : isAuto
      ? i18n("setup.auto.discoverButton", "Discover")
      : i18n("setup.authButton", "Create account");
}

let typeDropdown;

// ── Submit handlers ──────────────────────────────────────────────────────

async function submitOAuth(type) {
  const email = val("oauth-email");

  if (!email) {
    showError(i18n("setup.oauth.error.emailRequired", "Please enter your Microsoft account email."));
    $("oauth-email").focus();
    return;
  }
  if (!setupToken) {
    showError(i18n("setup.error.missingToken", "Missing setup token. Open this window through TbSync."));
    return;
  }

  const btn = $("btn-submit");
  btn.disabled = true;
  try {
    // Run Microsoft consent in a separate popup driven by the background
    // (see oauth.mjs::runConsentPopup). On success we get the token bundle
    // back and feed it straight into createAccount.
    const authReply = await browser.runtime.sendMessage({
      type: "eas.startOAuth",
      loginHint: email,
      servertype: type,
    });
    if (!authReply?.ok) {
      if (authReply?.code === "E:CANCELLED") { btn.disabled = false; return; }
      throw new Error(authReply?.error ?? i18n("setup.oauth.error.signInFailed", "Sign-in failed"));
    }
    const tokens = authReply.result;

    const reply = await browser.runtime.sendMessage({
      type: "eas.createAccount",
      servertype: type,
      label: tokens.authenticatedUserEmail || email,
      refreshToken: tokens.refreshToken,
      authenticatedUserEmail: tokens.authenticatedUserEmail,
      loginHint: email,
    });
    if (!reply?.ok) {
      throw new Error(reply?.error ?? i18n("setup.error.createFailed", "Could not create the account"));
    }
    await browser.runtime.sendMessage({
      type: "tbsync-setup-completed",
      setupToken,
      accountName: reply.result.accountName,
      initialFolders: reply.result.initialFolders,
      custom: reply.result.custom,
    });
    window.close();
  } catch (err) {
    btn.disabled = false;
    showError(err.message ?? String(err));
  }
}

async function submitAuto() {
  const email    = val("auto-email");
  const password = $("auto-password").value;

  if (!email)    { showError(i18n("setup.oauth.error.emailRequired", "Please enter your email address.")); $("auto-email").focus();    return; }
  if (!password) { showError(i18n("setup.error.passwordRequired",    "Please enter the password."));        $("auto-password").focus(); return; }
  if (!setupToken) {
    showError(i18n("setup.error.missingToken", "Missing setup token. Open this window through TbSync."));
    return;
  }

  const btn = $("btn-submit");
  const status = $("auto-status");
  btn.disabled = true;
  status.textContent = i18n("setup.auto.status.searching", "Looking up server settings…");
  status.hidden = false;

  try {
    const reply = await browser.runtime.sendMessage({
      type: "eas.discoverServer",
      email, password,
    });

    if (!reply?.ok) {
      showAutoError(reply);
      return;
    }

    const created = await browser.runtime.sendMessage({
      type: "eas.createAccount",
      servertype: TYPE_AUTO,
      email, password,
      server: reply.result.server,
      user: reply.result.user || email,
    });
    if (!created?.ok) {
      throw new Error(created?.error ?? i18n("setup.error.createFailed", "Could not create the account"));
    }
    await browser.runtime.sendMessage({
      type: "tbsync-setup-completed",
      setupToken,
      accountName: created.result.accountName,
      initialFolders: created.result.initialFolders,
      custom: created.result.custom,
    });
    window.close();
  } catch (err) {
    showAutoError({ ok: false, code: null, error: err?.message ?? String(err), details: null });
  } finally {
    if (!document.hidden) {
      btn.disabled = false;
      status.hidden = true;
    }
  }
}

function showAutoError(reply) {
  const detailEl   = $("auto-error-detail");
  const triedLabel = $("auto-error-tried-label");
  const triedList  = $("auto-error-tried");
  const code = reply?.code ?? null;

  let msgKey, msgFallback;
  if (code === "E:AUTH") {
    msgKey = "setup.auto.error.auth";
    msgFallback = "The username or password was rejected.";
  } else if (code === "E:NETWORK") {
    msgKey = "setup.auto.error.network";
    msgFallback = "Could not reach the discovery endpoints.";
  } else {
    msgKey = "setup.auto.error.noServer";
    msgFallback = "No EAS server found for this email.";
  }
  detailEl.textContent = i18n(msgKey, msgFallback);

  triedList.replaceChildren();
  const tried = Array.isArray(reply?.details?.tried) ? reply.details.tried : [];
  if (tried.length) {
    triedLabel.hidden = false;
    for (const t of tried) {
      const li = document.createElement("li");
      li.textContent = `${t.url} — ${t.status}`;
      triedList.appendChild(li);
    }
  } else {
    triedLabel.hidden = true;
  }
  $("auto-error-popover").showPopover();
}

async function submitBasic() {
  const label    = val("account-name");
  const server   = val("server");
  const user     = val("user");
  const password = $("password").value;

  if (!label)    { showError(i18n("setup.error.labelRequired",    "Please enter an account label."));   $("account-name").focus(); return; }
  if (!server)   { showError(i18n("setup.error.serverRequired",   "Please enter the server URL."));     $("server").focus();       return; }
  if (!user)     { showError(i18n("setup.error.userRequired",     "Please enter the username."));       $("user").focus();         return; }
  if (!password) { showError(i18n("setup.error.passwordRequired", "Please enter the password."));       $("password").focus();     return; }
  if (!setupToken) {
    showError(i18n("setup.error.missingToken", "Missing setup token. Open this window through TbSync."));
    return;
  }

  const btn = $("btn-submit");
  btn.disabled = true;
  try {
    const reply = await browser.runtime.sendMessage({
      type: "eas.createAccount",
      servertype: TYPE_CUSTOM,
      label, server, user, password,
    });
    if (!reply?.ok) {
      throw new Error(reply?.error ?? i18n("setup.error.createFailed", "Could not create the account"));
    }
    await browser.runtime.sendMessage({
      type: "tbsync-setup-completed",
      setupToken,
      accountName: reply.result.accountName,
      initialFolders: reply.result.initialFolders,
      custom: reply.result.custom,
    });
    window.close();
  } catch (err) {
    btn.disabled = false;
    showError(err.message ?? String(err));
  }
}

async function onSubmit() {
  clearError();
  const type = typeDropdown.getValue();
  if (type === TYPE_CUSTOM) return submitBasic();
  if (type === TYPE_AUTO)   return submitAuto();
  return submitOAuth(type);
}

// ── Boot ─────────────────────────────────────────────────────────────────

localizeDocument();

typeDropdown = createDropdown($("account-type"), {
  options: [
    {
      value: TYPE_OFFICE365,
      label: i18n("setup.accountType.office365", "Office 365 Business"),
      hint:  i18n("setup.accountType.office365.hint", ""),
    },
    {
      value: TYPE_PERSONAL_MS,
      label: i18n("setup.accountType.personalMs", "Personal Microsoft account"),
      hint:  i18n("setup.accountType.personalMs.hint", ""),
    },
    {
      value: TYPE_AUTO,
      label: i18n("setup.accountType.auto", "Auto-detect"),
      hint:  i18n("setup.accountType.auto.hint", ""),
    },
    {
      value: TYPE_CUSTOM,
      label: i18n("setup.accountType.custom", "Custom EAS server"),
      hint:  i18n("setup.accountType.custom.hint", ""),
    },
  ],
  value: TYPE_OFFICE365,
  onChange: applyType,
});

applyType(typeDropdown.getValue());

$("btn-cancel").addEventListener("click", () => window.close());
$("btn-submit").addEventListener("click", onSubmit);
$("auto-error-close").addEventListener("click", () => $("auto-error-popover").hidePopover());

// ESC closes the dialog; Enter while focused in a text input fires the
// primary action (when enabled and visible). `defaultPrevented` lets the
// dropdown's own Escape handler swallow the key when its panel is open.
document.addEventListener("keydown", e => {
  if (e.defaultPrevented) return;
  if (e.key === "Escape") {
    window.close();
    return;
  }
  if (e.key === "Enter" && e.target?.tagName === "INPUT") {
    const btn = document.querySelector("button.primary:not([hidden])");
    if (btn && !btn.disabled) {
      e.preventDefault();
      btn.click();
    }
  }
});
