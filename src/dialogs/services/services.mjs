/**
 * Services window - settings kept on the server rather than by TbSync.
 *
 * Out of office is the only one so far. Unlike the account settings, which
 * are stored here and locked while the account is connected, everything in
 * this window needs the connection: it is read from the mailbox on open and
 * written back on save, and nothing is cached in between. A stale value
 * shown here would be worse than a slow one.
 *
 * The three audiences Exchange keeps - inside the organisation, known
 * outside contacts, everyone else - are shown as two controls, because the
 * distinction between the last two is not one most people want to make. The
 * two external audiences are written together, which is what Outlook's own
 * "outside my organisation" checkbox does.
 */

import { localizeDocument } from "../../vendor/i18n/i18n.mjs";

const $ = (id) => document.getElementById(id);
const i18n = (key, fallback) => browser.i18n.getMessage(key) || fallback;

const params = new URLSearchParams(location.search);
const accountId = params.get("accountId");

/** What the server said when the window opened, so a save can send back the
 *  audiences the user did not touch exactly as they were. */
let loaded = null;

function showError(text) {
  const box = $("error");
  box.textContent = text;
  box.classList.add("visible");
}

function clearError() {
  $("error").classList.remove("visible");
}

/** Oof times are UTC, and unlike the rest of EAS they are full ISO with
 *  milliseconds - `2026-08-13T18:00:00.000Z`. The compact `20260813T180000Z`
 *  form is accepted too, because it costs three lines and a server that
 *  sent it would otherwise show the user two empty fields.
 *
 *  `datetime-local` speaks the user's own clock, so both directions go
 *  through `Date` rather than slicing the string: displaying 18:00 for an
 *  18:00Z the user reads as 20:00 would be a lie, and writing their 09:00
 *  back as 09:00Z would move it. */
function toInputValue(wire) {
  const compact = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z?$/.exec(
    wire ?? "",
  );
  const date = new Date(
    compact
      ? `${compact[1]}-${compact[2]}-${compact[3]}T${compact[4]}:${compact[5]}:${compact[6]}Z`
      : (wire ?? ""),
  );
  if (isNaN(date)) return "";
  const pad = (n) => String(n).padStart(2, "0");
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
    `T${pad(date.getHours())}:${pad(date.getMinutes())}`
  );
}

function toWireValue(input) {
  if (!input) return "";
  const date = new Date(input);
  return isNaN(date) ? "" : date.toISOString();
}

/** Say so when the mailbox holds a message we are about to flatten. This
 *  window has no rich editor, so an HTML reply is shown as its markup and
 *  would be saved back as plain text. */
function formatHint(el, message) {
  el.textContent =
    String(message?.bodyType ?? "Text").toUpperCase() === "HTML"
      ? i18n(
          "services.oof.htmlWarning",
          "This message is formatted on the server. Saving here replaces it with plain text.",
        )
      : "";
}

function render(oof) {
  $("oof-state").value = String(oof.state ?? "0");
  $("oof-start").value = toInputValue(oof.startTime);
  $("oof-end").value = toInputValue(oof.endTime);

  const internal = oof.messages?.internal;
  $("oof-internal").value = internal?.reply ?? "";
  formatHint($("oof-internal-hint"), internal);

  // The two external audiences share one control. If the mailbox has them
  // saying different things, the known-contacts one is shown - it is the
  // one a reply actually reaches most often.
  const known = oof.messages?.externalKnown;
  const unknown = oof.messages?.externalUnknown;
  $("oof-external-on").checked = !!(known?.enabled || unknown?.enabled);
  $("oof-external").value = known?.reply || unknown?.reply || "";
  formatHint($("oof-external-hint"), known?.enabled ? known : unknown);

  updateVisibility();
}

function updateVisibility() {
  // The state is the master switch: with it off nothing else is sent, so
  // nothing else is shown. A box that takes text and then drops it on save
  // is worse than no box at all.
  const state = $("oof-state").value;
  const on = state !== "0";
  // A server that named neither external audience does not have them, so
  // the control goes rather than sitting there switched off. See
  // `readOofFromDoc`.
  const hasExternal = !!(
    loaded?.messages?.externalKnown || loaded?.messages?.externalUnknown
  );
  $("internal-row").hidden = !on;
  $("external-block").hidden = !on || !hasExternal;
  $("window-row").hidden = state !== "2";
  $("external-row").hidden = !$("oof-external-on").checked;
}

function collect() {
  const state = $("oof-state").value;
  const on = state !== "0";
  const external = $("oof-external-on").checked && on;
  const externalReply = $("oof-external").value;
  // Send back the audiences the mailbox itself named, and no others.
  const messages = {};
  if (loaded?.messages?.internal) {
    messages.internal = { enabled: on, reply: $("oof-internal").value };
  }
  // Written together: the window offers one control for both, so sending
  // them separately would let them drift apart invisibly.
  for (const name of ["externalKnown", "externalUnknown"]) {
    if (loaded?.messages?.[name]) {
      messages[name] = { enabled: external, reply: externalReply };
    }
  }
  return {
    state,
    startTime: toWireValue($("oof-start").value),
    endTime: toWireValue($("oof-end").value),
    messages,
  };
}

async function load() {
  clearError();
  const reply = await browser.runtime.sendMessage({
    type: "eas.readOof",
    accountId,
  });
  $("loading").hidden = true;
  if (!reply?.ok) {
    showError(
      reply?.error ??
        i18n("services.error.loadFailed", "Could not read the settings."),
    );
    return;
  }
  if (!reply.result) {
    showError(
      i18n(
        "services.oof.unsupported",
        "This server does not offer an out-of-office reply.",
      ),
    );
    return;
  }
  loaded = reply.result;
  render(loaded);
  $("form").hidden = false;
  $("btn-save").disabled = false;
}

async function save() {
  clearError();
  $("btn-save").disabled = true;
  const reply = await browser.runtime.sendMessage({
    type: "eas.writeOof",
    accountId,
    settings: collect(),
  });
  if (!reply?.ok) {
    showError(
      reply?.error ??
        i18n("services.error.saveFailed", "Could not save the settings."),
    );
    $("btn-save").disabled = false;
    return;
  }
  // Read back rather than trust the write: the server normalises what it
  // stores, and this window is only useful if it shows the mailbox.
  $("loading").hidden = false;
  $("form").hidden = true;
  await load();
}

$("oof-state").addEventListener("change", updateVisibility);
$("oof-external-on").addEventListener("change", updateVisibility);
$("btn-cancel").addEventListener("click", () => window.close());
$("btn-save").addEventListener("click", save);

localizeDocument();
load();
