/**
 * Thin wrapper over `messenger.calendar.calendars.*` and
 * `messenger.calendar.items.*`. Always exchanges iCal strings - the
 * wrapper pins `format: "ical"` on every write and `returnFormat: "ical"`
 * on every read so callers never need to think about jCal.
 *
 * Calendars are created as **ext-type** Lightning calendars owned by
 * this extension. The user-facing calendar id (`calendar.id`) is what
 * Lightning routes `provider.onItem*` events on; the cache id
 * (`calendar.cacheId`) is the writable storage substrate the runner
 * uses for `items.*` reads/writes during sync. The two are surfaced
 * separately so callers can keep them straight.
 *
 * Calendar operations tolerate "not found" (the user may have deleted
 * the calendar manually); item-level reads/writes throw on real errors
 * so the sync orchestrator can surface them.
 */

const ICAL_FORMAT = "ical";

const EXT_TYPE = "ext-" + browser.runtime.id;

/* ── Calendar level ───────────────────────────────────────────────── */

/**
 * Create an ext-type Lightning calendar bound to this extension. Per-
 * calendar `capabilities` narrow the manifest defaults (e.g. `events:
 * true, tasks: false` for a Calendar folder; `events: false, tasks:
 * true` for a Tasks folder) and pin `organizer` / `organizerName` to
 * the EAS account identity so TB's iTIP code knows who owns the
 * calendar.
 *
 * Returns `{ id, cacheId }`. `id` is what `provider.onItem*` events
 * route on; `cacheId` is the underlying writable cache calendar.
 */
export async function createCalendar({ name, color, url, capabilities }) {
  if (!name || typeof name !== "string" || !name.trim()) {
    throw new Error("createCalendar requires a non-empty name");
  }
  if (!url || typeof url !== "string") {
    throw new Error("createCalendar requires a unique url");
  }
  const props = {
    name: name.trim(),
    type: EXT_TYPE,
    url,
  };
  if (color) props.color = color;
  if (capabilities) props.capabilities = capabilities;
  const calendar = await messenger.calendar.calendars.create(props);
  if (!calendar?.id) {
    throw new Error("createCalendar: calendars.create returned no id");
  }
  return { id: calendar.id, cacheId: calendar.cacheId ?? calendar.id };
}

export async function deleteCalendar(id) {
  if (!id) return;
  try {
    await messenger.calendar.calendars.remove(id);
  } catch (err) {
    if (isNotFoundError(err)) return;
    throw err;
  }
}

export async function calendarExists(id) {
  if (!id) return false;
  try {
    const cal = await messenger.calendar.calendars.get(id);
    return !!cal;
  } catch {
    return false;
  }
}

export async function renameCalendar(id, name) {
  if (!id) throw new Error("renameCalendar requires an id");
  await messenger.calendar.calendars.update(id, { name });
}

/**
 * Mirror a folder's effective read-only state onto the local Thunderbird
 * calendar. When set, TB greys out event editing in the UI; the experiment's
 * sync write path bypasses the flag, so the runner can still apply server
 * changes to the local store. Tolerant of "calendar not found" because the
 * user may have deleted it manually since the folder row was bound.
 */
export async function setCalendarReadOnly(id, readOnly) {
  if (!id) return;
  try {
    await messenger.calendar.calendars.update(id, { readOnly: !!readOnly });
  } catch (err) {
    if (isNotFoundError(err)) return;
    throw err;
  }
}

/* ── Item level ───────────────────────────────────────────────────── */

/**
 * List items in a calendar, optionally filtered by type ("event" or
 * "task"). Returns `[{ id, type, item: <iCal string> }]`. Tolerates
 * "calendar not found" by returning [].
 */
export async function listItems(calendarId, type) {
  if (!calendarId) return [];
  try {
    const queryOpts = { calendarId, returnFormat: ICAL_FORMAT };
    if (type) queryOpts.type = type;
    const list = await messenger.calendar.items.query(queryOpts);
    return list.map(normalizeItem);
  } catch (err) {
    if (isNotFoundError(err)) return [];
    throw err;
  }
}

export async function getItem(calendarId, id) {
  if (!calendarId || !id) return null;
  try {
    const node = await messenger.calendar.items.get(calendarId, id, {
      returnFormat: ICAL_FORMAT,
    });
    return node ? normalizeItem(node) : null;
  } catch (err) {
    if (isNotFoundError(err)) return null;
    throw err;
  }
}

/**
 * Create a calendar item. Pre-specifies `id` so the changelog freeze
 * key matches the announced `onCreated` event.
 */
export async function createItem(calendarId, { id, type, ical }) {
  if (!calendarId) throw new Error("createItem requires a calendarId");
  if (!id) throw new Error("createItem requires an id");
  if (!type) throw new Error("createItem requires a type");
  if (!ical) throw new Error("createItem requires an iCal string");
  const created = await messenger.calendar.items.create(calendarId, {
    id,
    type,
    format: ICAL_FORMAT,
    item: ical,
    returnFormat: ICAL_FORMAT,
  });
  return normalizeItem(created);
}

export async function updateItem(calendarId, id, { ical }) {
  if (!calendarId) throw new Error("updateItem requires a calendarId");
  if (!id) throw new Error("updateItem requires an id");
  if (!ical) throw new Error("updateItem requires an iCal string");
  await messenger.calendar.items.update(calendarId, id, {
    format: ICAL_FORMAT,
    item: ical,
    returnFormat: ICAL_FORMAT,
  });
}

export async function deleteItem(calendarId, id) {
  if (!calendarId || !id) return;
  try {
    await messenger.calendar.items.remove(calendarId, id);
  } catch (err) {
    if (isNotFoundError(err)) return;
    throw err;
  }
}

/* ── Helpers ──────────────────────────────────────────────────────── */

function normalizeItem(node) {
  if (!node) return node;
  return { id: node.id, type: node.type, item: node.item };
}

function isNotFoundError(err) {
  const msg = String(err?.message ?? err ?? "");
  return /no such|not found|invalid id|unknown calendar/i.test(msg);
}
