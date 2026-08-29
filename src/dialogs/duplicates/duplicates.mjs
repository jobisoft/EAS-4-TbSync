/**
 * Duplicate copies window.
 *
 * Opened by the provider after a sync found UIDs the server holds more
 * than once. Thunderbird keeps one item per UID, so none of this is
 * visible in the calendar itself - the surplus copies sit on the server,
 * come down again on every full sync, and are only ever seen here.
 *
 * One row per duplicated item, not per copy: the copies of one UID are the
 * same item, and listing 533 of them would be a list nobody can act on.
 * The count excludes the copy that stays, so it reads as "this many will
 * be removed".
 *
 * Deleting is the only thing here that touches the server, and it names
 * UIDs rather than ServerIds: which copies those are is decided in the
 * background against the finding the sync made, so a window left open
 * while the mailbox changed cannot ask for something else.
 */

import { localizeDocument } from "../../vendor/i18n/i18n.mjs";

const $ = (id) => document.getElementById(id);
const i18n = (key, fallback, subs) =>
  browser.i18n.getMessage(key, subs) || fallback;

const params = new URLSearchParams(location.search);
const accountId = params.get("accountId");

function showError(text) {
  const box = $("error");
  box.textContent = text;
  box.classList.add("visible");
}

function checkboxes() {
  return [...document.querySelectorAll("#rows input[type=checkbox]")];
}

function syncDeleteButton() {
  const picked = checkboxes().filter((c) => c.checked);
  $("btn-delete").disabled = picked.length === 0;
  const all = checkboxes();
  $("check-all").checked = picked.length === all.length && all.length > 0;
}

function render(rows) {
  const body = $("rows");
  body.replaceChildren();
  for (const row of rows) {
    const tr = document.createElement("tr");

    const check = document.createElement("td");
    const box = document.createElement("input");
    box.type = "checkbox";
    box.checked = true;
    box.dataset.uid = row.uid;
    box.addEventListener("change", syncDeleteButton);
    check.append(box);

    const calendar = document.createElement("td");
    calendar.textContent = row.folderName;

    const item = document.createElement("td");
    // An item with no title still has to be pickable, so it is named by
    // what it is rather than left as an empty cell.
    item.textContent =
      row.title || i18n("duplicates.item.untitled", "(no title)");

    const uid = document.createElement("td");
    uid.className = "uid";
    uid.textContent = row.uid;

    const copies = document.createElement("td");
    copies.className = "copies";
    copies.textContent = String(row.copies);

    tr.append(check, calendar, item, uid, copies);
    body.append(tr);
  }
  syncDeleteButton();
}

async function load() {
  const reply = await browser.runtime.sendMessage({
    type: "eas.readDuplicates",
    accountId,
  });
  $("loading").hidden = true;
  if (!reply?.ok) {
    showError(
      reply?.error ??
        i18n("duplicates.error.loadFailed", "Could not read the findings."),
    );
    return;
  }
  if (!reply.result?.rows?.length) {
    showError(i18n("duplicates.empty", "Nothing is duplicated any more."));
    return;
  }
  render(reply.result.rows);
  $("content").hidden = false;
}

/** Chunk-by-chunk progress, pushed from the background while the sync that
 *  does the deleting runs. Scoped to this account, since another one could
 *  be cleaning up at the same time. */
browser.runtime.onMessage.addListener((msg) => {
  if (msg?.type !== "eas.duplicatesProgress" || msg.accountId !== accountId) {
    return;
  }
  $("progress").textContent = i18n(
    "duplicates.progress",
    `Removed ${msg.deleted} of ${msg.total}.`,
    [String(msg.deleted), String(msg.total)],
  );
  $("progress").hidden = false;
});

function setBusy(busy) {
  $("btn-delete").disabled = busy;
  $("check-all").disabled = busy;
  for (const box of checkboxes()) box.disabled = busy;
  // Close stays live: the removal runs inside a sync and finishes either
  // way, and several hundred copies take minutes.
}

async function remove() {
  const uids = checkboxes()
    .filter((c) => c.checked)
    .map((c) => c.dataset.uid);
  if (!uids.length) return;
  setBusy(true);
  $("btn-delete").textContent = i18n("duplicates.button.deleting", "Removing…");
  const reply = await browser.runtime.sendMessage({
    type: "eas.cleanupDuplicates",
    accountId,
    uids,
  });
  const restore = () => {
    $("progress").hidden = true;
    setBusy(false);
    $("btn-delete").textContent = i18n(
      "duplicates.button.delete",
      "Remove the extra copies",
    );
    syncDeleteButton();
  };
  if (!reply?.ok) {
    showError(
      reply?.error ??
        i18n("duplicates.error.deleteFailed", "Could not remove the copies."),
    );
    restore();
    return;
  }
  // The removal happens inside a sync, so the answer is what the row says
  // afterwards rather than what was asked for. Anything still listed was
  // refused by the server or never reached, and the list is reloaded so
  // the user sees which.
  if (reply.result?.remaining) {
    showError(
      i18n("duplicates.error.partial", "Some copies could not be removed."),
    );
    restore();
    await load();
    return;
  }
  window.close();
}

$("check-all").addEventListener("change", (e) => {
  for (const box of checkboxes()) box.checked = e.target.checked;
  syncDeleteButton();
});
$("btn-close").addEventListener("click", () => window.close());
$("btn-delete").addEventListener("click", remove);

localizeDocument();
load();
