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
import { normalizeCustomServerUrl } from "../../modules/eas/server-url.mjs";

const i18n = (key, fallback, substitutions) =>
  browser.i18n.getMessage(key, substitutions) || fallback;

const params = new URLSearchParams(location.search);
const accountId = params.get("accountId");
const readOnly = params.get("readOnly") === "1";

const KNOWN_AS_VERSIONS = ["2.5", "14.0", "14.1", "16.1"];

const TYPE_OFFICE365 = "office365";
const TYPE_PERSONAL_MS = "personal-ms";
const TYPE_AUTO = "auto";
const TYPE_CUSTOM = "custom";

// Per-type 32px icons. Same set the setup popup uses; the trigger
// shows the icon for the selected (locked) type.
const ICON_OFFICE365 = browser.runtime.getURL("icons/365_32.png");
const ICON_PERSONAL_MS = browser.runtime.getURL("icons/365_32.png");
const ICON_AUTO = browser.runtime.getURL("icons/eas32.png");
const ICON_CUSTOM = browser.runtime.getURL("icons/eas32.png");

function deriveAccountType(account) {
  if (account.servertype === TYPE_OFFICE365) return TYPE_OFFICE365;
  if (account.servertype === TYPE_PERSONAL_MS) return TYPE_PERSONAL_MS;
  if (account.servertype === TYPE_AUTO) return TYPE_AUTO;
  return TYPE_CUSTOM;
}

const FIELD_IDS = [
  "account-name",
  "server",
  "user",
  "password",
  "as-version-selected",
  "provision",
  "conflict-policy",
  "contacts-display-override",
  "contacts-name-separator",
  "calendar-sync-limit",
  "sync-recurrence",
  "gal-enabled",
  "freebusy-enabled",
];

function $(id) {
  return document.getElementById(id);
}

function showError(message) {
  const el = $("error");
  el.textContent = message;
  el.classList.add("visible");
}
function clearError() {
  $("error").classList.remove("visible");
}

async function load() {
  if (!accountId) {
    showError(
      i18n("config.error.missingAccountId", "Missing account identifier."),
    );
    return;
  }
  const reply = await browser.runtime.sendMessage({
    type: "eas.getAccount",
    accountId,
  });
  if (!reply?.ok) {
    showError(
      reply?.error ??
        i18n("config.error.loadFailed", "Failed to load account."),
    );
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
        hint: i18n("setup.accountType.office365.hint", ""),
        icon: ICON_OFFICE365,
      },
      {
        value: TYPE_PERSONAL_MS,
        label: i18n(
          "setup.accountType.personalMs",
          "Personal Microsoft account",
        ),
        hint: i18n("setup.accountType.personalMs.hint", ""),
        icon: ICON_PERSONAL_MS,
      },
      {
        value: TYPE_AUTO,
        label: i18n("setup.accountType.auto", "Auto-detect"),
        hint: i18n("setup.accountType.auto.hint", ""),
        icon: ICON_AUTO,
      },
      {
        value: TYPE_CUSTOM,
        label: i18n("setup.accountType.custom", "Custom EAS server"),
        hint: i18n("setup.accountType.custom.hint", ""),
        icon: ICON_CUSTOM,
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
    $("user").value = account.user ?? "";
    const lockServerUser = accountType === TYPE_AUTO;
    $("server").readOnly = lockServerUser;
    $("user").readOnly = lockServerUser;
    // Password is always blank on load.
  } else {
    $("connection-section").hidden = true;
  }

  // ── Protocol section ───────────────────────────────────────────────────
  $("device-id").textContent = account.deviceId ?? "";
  populateAsVersionDropdown(account);
  $("provision").checked = account.provision !== false;
  $("conflict-policy").value = account.conflict || "1";

  // ── Contacts section ───────────────────────────────────────────────────
  $("contacts-display-override").checked = !!account.contactsDisplayOverride;
  $("contacts-name-separator").value = account.contactsNameSeparator || "10";

  // ── Calendar section ───────────────────────────────────────────────────
  $("calendar-sync-limit").value = account.calendarSyncLimit || "7";
  $("sync-recurrence").checked = !!account.syncRecurrence;
  $("sync-on-change").value = String(account.syncOnChange ?? "15");

  // ── Server lookups ─────────────────────────────────────────────────────
  // Each checkbox is forced off + disabled when the server's OPTIONS probe
  // didn't advertise what it needs - the Search command for the address
  // list, availability for free/busy. The hint flips to a "not supported by
  // your server" line so the user understands why.
  $("gal-enabled").checked = account.galSupported && !!account.galEnabled;
  $("freebusy-enabled").checked =
    account.freeBusySupported && !!account.freeBusyEnabled;

  applyReadOnly();

  // applyReadOnly clears `disabled` on every field when the account is
  // editable; the GAL "not supported" disable must outlive that pass.
  if (!account.galSupported) {
    $("gal-enabled").disabled = true;
    $("gal-enabled-hint").textContent = i18n(
      "config.gal.notSupported",
      "Your server does not advertise the Search command, so the Global Address List is not available.",
    );
  }
  if (!account.freeBusySupported) {
    $("freebusy-enabled").disabled = true;
    $("freebusy-enabled-hint").textContent = i18n(
      "config.freebusy.notSupported",
      "Your server does not offer attendee availability, so free/busy times cannot be shown.",
    );
  }
}

/** Which versions to offer: the ones this add-on speaks and the server
 *  says it speaks, intersected.
 *
 *  Not the server's list on its own - a server advertising 12.0 must not
 *  offer a version we cannot talk. And not a filtered list when the
 *  server named none: that is the case where the probe never answered,
 *  which is exactly when the user has to choose one blind, so an empty
 *  intersection would leave a dropdown offering nothing but "auto".
 *
 *  A value already selected stays on the list whatever the server now
 *  advertises, so opening this window cannot silently change the setting
 *  it is showing. */
function offeredAsVersions(account) {
  const advertised = Array.isArray(account.allowedAsVersions)
    ? account.allowedAsVersions
    : [];
  const base = advertised.length
    ? KNOWN_AS_VERSIONS.filter((v) => advertised.includes(v))
    : KNOWN_AS_VERSIONS;
  const list = base.length ? base : KNOWN_AS_VERSIONS;
  const selected = account.asVersionSelected;
  return selected && selected !== "auto" && !list.includes(selected)
    ? [...list, selected]
    : list;
}

/** Build the AS-version dropdown: "auto" plus whatever this server can be
 *  asked for. The hint underneath says what the account is using and what
 *  the server would pick for it. */
function populateAsVersionDropdown(account) {
  const sel = $("as-version-selected");
  sel.innerHTML = "";
  const autoOpt = document.createElement("option");
  autoOpt.value = "auto";
  autoOpt.textContent = i18n("config.protocol.asVersion.auto", "");
  sel.appendChild(autoOpt);

  for (const v of offeredAsVersions(account)) {
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
  const advertised = Array.isArray(account.allowedAsVersions)
    ? account.allowedAsVersions
    : [];
  if (!advertised.length) {
    // Said out loud, because the list is the full one and the reason is
    // not visible: the server never answered the version probe, so this
    // is a choice the user has to make without its help.
    hintEl.textContent = i18n("config.protocol.asVersion.unknownHint", "");
    return;
  }
  // Two different questions, and the popup is where they get confused.
  // What the account speaks is settled when it connects and cannot move
  // while it is enabled; what the server suggests is re-read every day.
  // They differ whenever a version has been chosen by hand, so both are
  // said rather than one standing in for the other.
  const lines = [];
  if (readOnly && account.asVersion) {
    lines.push(
      i18n("config.protocol.asVersion.usingHint", "", [account.asVersion]),
    );
  }
  if (account.asVersionSuggested) {
    lines.push(
      i18n("config.protocol.asVersion.suggestedHint", "", [
        account.asVersionSuggested,
      ]),
    );
  }
  hintEl.textContent = lines.join(" ");
}

function applyReadOnly() {
  const banner = $("readonly-banner");
  if (readOnly) {
    banner.textContent = i18n(
      "config.readOnlyBanner",
      "To prevent synchronization errors, settings cannot be edited while the account is enabled.",
    );
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
    showError(
      i18n("config.error.accountNameRequired", "Account name is required."),
    );
    return;
  }

  const asVersionSelected = $("as-version-selected").value;
  if (
    asVersionSelected !== "auto" &&
    !KNOWN_AS_VERSIONS.includes(asVersionSelected)
  ) {
    showError(
      i18n(
        "config.error.invalidAsVersion",
        "Invalid ActiveSync version selection.",
      ),
    );
    return;
  }

  const patch = {
    accountName,
    asVersionSelected,
    provision: $("provision").checked,
    conflict: $("conflict-policy").value,
    contactsDisplayOverride: $("contacts-display-override").checked,
    contactsNameSeparator: $("contacts-name-separator").value,
    calendarSyncLimit: $("calendar-sync-limit").value,
    syncRecurrence: $("sync-recurrence").checked,
    syncOnChange: $("sync-on-change").value,
  };

  // Only thread `galEnabled` through when the field is actually
  // interactive - sending a forced-off value for an unsupported server
  // would clobber the per-account preference if support is restored.
  if (!$("gal-enabled").disabled) {
    patch.galEnabled = $("gal-enabled").checked;
  }
  if (!$("freebusy-enabled").disabled) {
    patch.freeBusyEnabled = $("freebusy-enabled").checked;
  }

  // Connection fields only flow through when the section is visible. For
  // auto-detect accounts the server/user inputs are readOnly, so only the
  // (optional) password actually changes.
  if (!$("connection-section").hidden) {
    // Editable only for custom accounts; auto-detect keeps the address
    // Autodiscover returned. Same rule as the setup dialog.
    if (!$("server").readOnly) {
      const server = $("server").value.trim();
      if (!normalizeCustomServerUrl(server)) {
        showError(
          i18n("setup.error.serverInvalid", "Not a usable server URL.", [
            server,
          ]),
        );
        $("server").focus();
        return;
      }
      patch.server = server;
    }
    if (!$("user").readOnly) patch.user = $("user").value.trim();
    const pw = $("password").value;
    if (pw) patch.password = pw;
  }

  $("btn-save").disabled = true;
  try {
    const reply = await browser.runtime.sendMessage({
      type: "eas.saveAccount",
      accountId,
      patch,
    });
    if (!reply?.ok) {
      throw new Error(
        reply?.error ?? i18n("config.error.saveFailed", "Save failed"),
      );
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
document.addEventListener("keydown", (e) => {
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
