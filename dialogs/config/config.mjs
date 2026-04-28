/**
 * Config popup controller. Single scrollable column of sectioned
 * settings: Account, Connection (custom-EAS-server only), Protocol,
 * Contacts, Calendar.
 *
 * `readOnly=true` means the account is currently connected in TbSync;
 * we must not allow edits while it's live. Banner explains why; every
 * input renders disabled; Save is hidden; Cancel becomes Close.
 *
 * Account type is set at setup and not editable here. The dropdown is
 * always rendered in `locked` mode - purely for visual parity with the
 * setup dialog.
 */

import { localizeDocument } from "../../vendor/i18n/i18n.mjs";
import { createDropdown } from "../shared/dropdown.mjs";

const i18n = (key, fallback, substitutions) =>
  browser.i18n.getMessage(key, substitutions) || fallback;

const params = new URLSearchParams(location.search);
const accountId = params.get("accountId");
const readOnly = params.get("readOnly") === "1";

const KNOWN_AS_VERSIONS = ["2.5", "14.0", "14.1", "16.1"];

const TYPE_OFFICE365   = "office365";
const TYPE_PERSONAL_MS = "personal-ms";
const TYPE_AUTO        = "auto";
const TYPE_CUSTOM      = "custom";

function deriveAccountType(account) {
  if (account.servertype === TYPE_OFFICE365)   return TYPE_OFFICE365;
  if (account.servertype === TYPE_PERSONAL_MS) return TYPE_PERSONAL_MS;
  if (account.servertype === TYPE_AUTO)        return TYPE_AUTO;
  return TYPE_CUSTOM;
}

const FIELD_IDS = [
  "account-name",
  "server", "user", "password",
  "as-version-selected", "provision",
  "contacts-display-override", "contacts-name-separator",
  "calendar-sync-limit", "sync-recurrence",
];

function $(id) { return document.getElementById(id); }

function showError(message) {
  const el = $("error");
  el.textContent = message;
  el.classList.add("visible");
}
function clearError() { $("error").classList.remove("visible"); }

async function load() {
  if (!accountId) {
    showError(i18n("config.error.missingAccountId", "Missing account identifier."));
    return;
  }
  const reply = await browser.runtime.sendMessage({ type: "eas.getAccount", accountId });
  if (!reply?.ok) {
    showError(reply?.error ?? i18n("config.error.loadFailed", "Failed to load account."));
    return;
  }
  const account = reply.result;

  // ── Account section ────────────────────────────────────────────────────
  $("account-name").value = account.accountName ?? "";

  const accountType = deriveAccountType(account);
  createDropdown($("account-type"), {
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
    value: accountType,
    locked: true,
  });

  if (accountType !== TYPE_CUSTOM && account.authenticatedUserEmail) {
    $("oauth-identity-row").hidden = false;
    $("oauth-identity").textContent = account.authenticatedUserEmail;
  }

  // ── Connection section ─────────────────────────────────────────────────
  // Visible for custom EAS and auto-detect accounts. For auto-detect, the
  // server and user came from the Autodiscover response and stay readonly;
  // only the password is editable.
  if (accountType === TYPE_CUSTOM || accountType === TYPE_AUTO) {
    $("connection-section").hidden = false;
    $("server").value = account.server ?? "";
    $("user").value   = account.user ?? "";
    const lockServerUser = accountType === TYPE_AUTO;
    $("server").readOnly = lockServerUser;
    $("user").readOnly   = lockServerUser;
    // Password is always blank on load.
  } else {
    $("connection-section").hidden = true;
  }

  // ── Protocol section ───────────────────────────────────────────────────
  $("device-id").textContent = account.deviceId ?? "";
  populateAsVersionDropdown(account);
  $("provision").checked = account.provision !== false;

  // ── Contacts section ───────────────────────────────────────────────────
  $("contacts-display-override").checked = !!account.contactsDisplayOverride;
  $("contacts-name-separator").value = account.contactsNameSeparator || "10";

  // ── Calendar section ───────────────────────────────────────────────────
  $("calendar-sync-limit").value = account.calendarSyncLimit || "7";
  $("sync-recurrence").checked = !!account.syncRecurrence;

  applyReadOnly();
}

/** Build the AS-version dropdown: "auto" plus the known fixed list,
 *  matching the legacy add-on. The hint underneath shows the currently
 *  negotiated version, but only while "auto" is selected. */
function populateAsVersionDropdown(account) {
  const sel = $("as-version-selected");
  sel.innerHTML = "";
  const autoOpt = document.createElement("option");
  autoOpt.value = "auto";
  autoOpt.textContent = i18n("config.protocol.asVersion.auto", "");
  sel.appendChild(autoOpt);

  for (const v of KNOWN_AS_VERSIONS) {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    sel.appendChild(opt);
  }
  sel.value = account.asVersionSelected || "auto";
  sel.addEventListener("change", () => updateAsVersionHint(account));
  updateAsVersionHint(account);
}

function updateAsVersionHint(account) {
  const sel = $("as-version-selected");
  const hintEl = $("as-version-hint");
  if (sel.value === "auto" && account.asVersion) {
    hintEl.textContent = i18n("config.protocol.asVersion.negotiatedHint", "", [account.asVersion]);
  } else {
    hintEl.textContent = "";
  }
}

function applyReadOnly() {
  const banner = $("readonly-banner");
  if (readOnly) {
    banner.textContent = i18n("config.readOnlyBanner",
      "To prevent synchronization errors, settings cannot be edited while the account is enabled.");
    banner.classList.add("visible");
  } else {
    banner.classList.remove("visible");
  }
  for (const id of FIELD_IDS) {
    const el = $(id);
    if (el) el.disabled = readOnly;
  }
  $("btn-save").hidden = readOnly;
  const cancelBtn = $("btn-cancel");
  cancelBtn.textContent = readOnly
    ? i18n("config.close", "Close")
    : i18n("config.cancel", "Cancel");
}

async function onSave() {
  if (readOnly) return;
  clearError();

  const accountName = $("account-name").value.trim();
  if (!accountName) {
    showError(i18n("config.error.accountNameRequired", "Account name is required."));
    return;
  }

  const asVersionSelected = $("as-version-selected").value;
  if (asVersionSelected !== "auto" && !KNOWN_AS_VERSIONS.includes(asVersionSelected)) {
    showError(i18n("config.error.invalidAsVersion", "Invalid ActiveSync version selection."));
    return;
  }

  const patch = {
    accountName,
    asVersionSelected,
    provision: $("provision").checked,
    contactsDisplayOverride: $("contacts-display-override").checked,
    contactsNameSeparator: $("contacts-name-separator").value,
    calendarSyncLimit: $("calendar-sync-limit").value,
    syncRecurrence: $("sync-recurrence").checked,
  };

  // Connection fields only flow through when the section is visible. For
  // auto-detect accounts the server/user inputs are readOnly, so only the
  // (optional) password actually changes.
  if (!$("connection-section").hidden) {
    if (!$("server").readOnly) patch.server = $("server").value.trim();
    if (!$("user").readOnly)   patch.user   = $("user").value.trim();
    const pw = $("password").value;
    if (pw) patch.password = pw;
  }

  $("btn-save").disabled = true;
  try {
    const reply = await browser.runtime.sendMessage({ type: "eas.saveAccount", accountId, patch });
    if (!reply?.ok) {
      throw new Error(reply?.error ?? i18n("config.error.saveFailed", "Save failed"));
    }
    window.close();
  } catch (err) {
    showError(err.message ?? String(err));
    $("btn-save").disabled = false;
  }
}

localizeDocument();
load();
$("btn-cancel").addEventListener("click", () => window.close());
$("btn-save").addEventListener("click", onSave);

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
