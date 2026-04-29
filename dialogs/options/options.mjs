/**
 * Advanced settings page. Six storage.local keys, all optional - empty
 * input means "use the bundled default", so an empty value is persisted
 * as `storage.local.remove(key)` rather than as an empty string. The
 * msTodoCompat boolean is similar: only the `true` state is written.
 */

import { localizeDocument } from "../../vendor/i18n/i18n.mjs";

const $ = (id) => document.getElementById(id);

const STRING_FIELDS = [
  { key: "timeout", inputId: "opt-timeout", type: "number" },
  { key: "tbsync.useragent", inputId: "opt-useragent", type: "string" },
  { key: "tbsync.type", inputId: "opt-devicetype", type: "string" },
  { key: "maxItems", inputId: "opt-maxitems", type: "number" },
  { key: "oauth.clientID", inputId: "opt-clientid", type: "string" },
];

async function load() {
  const keys = STRING_FIELDS.map((f) => f.key).concat(
    "msTodoCompat",
    "showItemsInTrash",
  );
  const stored = await browser.storage.local.get(keys);

  for (const f of STRING_FIELDS) {
    const v = stored[f.key];
    if (v === undefined || v === null || v === "") continue;
    $(f.inputId).value = String(v);
  }
  $("opt-mstodo").checked = stored.msTodoCompat === true;
  $("opt-show-trash").checked = stored.showItemsInTrash === true;
}

function bindStringField({ key, inputId, type }) {
  $(inputId).addEventListener("change", async (e) => {
    const raw = e.target.value.trim();
    if (raw === "") {
      await browser.storage.local.remove(key);
      e.target.value = "";
      return;
    }
    if (type === "number") {
      const n = Number(raw);
      if (!Number.isFinite(n) || n <= 0) {
        // Reject invalid input: revert to whatever is currently stored.
        const stored = await browser.storage.local.get({ [key]: null });
        const v = stored[key];
        e.target.value = v === null || v === undefined ? "" : String(v);
        return;
      }
      await browser.storage.local.set({ [key]: n });
      e.target.value = String(n);
    } else {
      await browser.storage.local.set({ [key]: raw });
      e.target.value = raw;
    }
  });
}

function bindCheckbox(inputId, key) {
  $(inputId).addEventListener("change", async (e) => {
    if (e.target.checked) await browser.storage.local.set({ [key]: true });
    else await browser.storage.local.remove(key);
  });
}

document.addEventListener("DOMContentLoaded", async () => {
  localizeDocument();
  await load();
  for (const f of STRING_FIELDS) bindStringField(f);
  bindCheckbox("opt-mstodo", "msTodoCompat");
  bindCheckbox("opt-show-trash", "showItemsInTrash");
});
