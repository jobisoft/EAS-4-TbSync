/**
 * EAS calendar provider event listeners.
 *
 * Bridges Lightning's `messenger.calendar.provider.*` events into the
 * host's changelog. Replaces the host's pre-migration
 * `messenger.calendar.items.on*` watcher for calendar/task folders:
 *
 *   - `onItemCreated/Updated/Removed` (request hooks for user-driven
 *     edits) → append `*_by_user` rows to `folder.changelog` via the
 *     `changelogAppendUserEntry` RPC. The runner drains those rows on
 *     the next manager-triggered sync.
 *   - `onSync(calendar)` is registered but no-op — sync remains
 *     manager-triggered. Lightning's refresh button just no-ops.
 *   - `calendar.calendars.onUpdated` (rename) and `onRemoved` (user
 *     deletes the calendar) take over the lifecycle signals the host's
 *     watcher used to handle.
 *
 * Provider-driven writes (sync from server → cache via
 * `calendar-store.mjs`) do NOT echo through `provider.onItem*`, so we
 * don't need a `_by_server` pre-tag dance for calendar items at all.
 *
 * Folder resolution: calendars are created with
 * `url = "ext+eas://<accountId>/<folderId>"`, so the listeners parse
 * the URL to find the owning folder. Accounts whose folders aren't yet
 * bound (or that belong to a different provider) silently no-op.
 */

const URL_PREFIX = "ext+eas://";

function parseExtUrl(url) {
  if (typeof url !== "string" || !url.startsWith(URL_PREFIX)) return null;
  const tail = url.slice(URL_PREFIX.length);
  const slash = tail.indexOf("/");
  if (slash <= 0) return null;
  const accountId = tail.slice(0, slash);
  const folderId = tail.slice(slash + 1);
  if (!accountId || !folderId) return null;
  return { accountId, folderId };
}

async function resolveFolder(provider, calendar) {
  const parsed = parseExtUrl(calendar?.url);
  if (!parsed) return null;
  const { accountId, folderId } = parsed;
  const rv = await provider.getAccount(accountId).catch(() => null);
  if (!rv?.account) return null;
  const folder = rv.folders?.find((f) => f.folderId === folderId);
  if (!folder) return null;
  return { accountId, folderId, folder, account: rv.account };
}

function changelogKindForItem(item) {
  return item?.type === "task" ? "task" : "event";
}

async function recordUserChange(provider, calendar, item, op, options) {
  const owner = await resolveFolder(provider, calendar);
  if (!owner || !item?.id) return item;

  // iTIP accept / decline / tentative arrives as `onItemUpdated` with
  // `options.invitation: true`. Per-calendar `capabilities.organizer`
  // (set at create time from the EAS account email) tells TB's iTIP
  // handler that the user is an attendee, not the organizer of any
  // received invitation — so the local ORGANIZER property is preserved
  // on accept. The standard Sync Change push then carries the original
  // ORGANIZER, the user's flipped PARTSTAT, and the server records the
  // response without re-attributing the meeting. A future enhancement
  // can route this row through the EAS `MeetingResponse` command (codec
  // primitive at `appendMeetingResponseBody`); for now the diagnostic
  // log makes the path visible.
  if (options?.invitation) {
    provider
      .reportEventLog({
        level: "info",
        accountId: owner.accountId,
        folderId: owner.folderId,
        message: `[calendar-provider] iTIP ${op} on ${item.id}: routed via Sync Change with capabilities.organizer-preserved ORGANIZER`,
      })
      .catch(() => {});
  }

  await provider
    .changelogAppendUserEntry({
      accountId: owner.accountId,
      folderId: owner.folderId,
      parentId: calendar.id,
      itemId: item.id,
      kind: changelogKindForItem(item),
      op,
    })
    .catch((err) =>
      console.warn(
        `[eas] changelogAppendUserEntry failed for ${owner.folderId}/${item.id}:`,
        err?.message ?? err,
      ),
    );
  return item;
}

async function handleCalendarRename(provider, calendar, changes) {
  if (!changes || !("name" in changes)) return;
  const owner = await resolveFolder(provider, calendar);
  if (!owner) return;
  if (owner.folder.targetName === calendar.name) return;
  await provider
    .updateFolder({
      accountId: owner.accountId,
      folderId: owner.folderId,
      patch: { targetName: calendar.name },
    })
    .catch((err) =>
      console.warn(
        `[eas] target-rename update failed for ${owner.folderId}:`,
        err?.message ?? err,
      ),
    );
}

async function handleCalendarRemoved(provider, calendarId) {
  // The user removed the local calendar from Lightning's UI. We don't
  // know which folder it belonged to without consulting the
  // calendar-id ↔ folder map; the host already reacts to "target
  // missing" on the next sync, so this listener is a soft hint:
  // unbind the folder eagerly so the manager reflects the change.
  const accounts = await provider.listAccounts().catch(() => []);
  for (const accountId of accounts) {
    const rv = await provider.getAccount(accountId).catch(() => null);
    const folder = rv?.folders?.find((f) => f.targetID === calendarId);
    if (!folder) continue;
    await provider
      .updateFolder({
        accountId,
        folderId: folder.folderId,
        patch: {
          targetID: null,
          targetName: null,
          selected: false,
          custom: {
            ...(folder.custom ?? {}),
            cacheId: null,
            providerCalendar: false,
          },
        },
      })
      .catch((err) =>
        console.warn(
          `[eas] target-removed update failed for ${folder.folderId}:`,
          err?.message ?? err,
        ),
      );
    return;
  }
}

export function init(provider) {
  // onSync is registered as a no-op so Lightning's refresh affordance
  // doesn't show an error when the user clicks it. Real syncs are
  // triggered by the TbSync manager (or autosync alarm) via
  // `onSyncFolder`.
  messenger.calendar.provider.onSync.addListener(() => null);

  messenger.calendar.provider.onItemCreated.addListener(
    async (calendar, item, options) =>
      recordUserChange(provider, calendar, item, "created", options),
  );
  messenger.calendar.provider.onItemUpdated.addListener(
    async (calendar, item, _oldItem, options) =>
      recordUserChange(provider, calendar, item, "updated", options),
  );
  messenger.calendar.provider.onItemRemoved.addListener(
    async (calendar, item, options) =>
      recordUserChange(provider, calendar, item, "deleted", options),
  );

  messenger.calendar.calendars.onUpdated.addListener((calendar, changes) =>
    handleCalendarRename(provider, calendar, changes),
  );
  messenger.calendar.calendars.onRemoved.addListener((id) =>
    handleCalendarRemoved(provider, id),
  );
}

// Kept exported so other modules (upgrade runner, tests) can re-use the
// URL parser without duplicating the format.
export { parseExtUrl };
