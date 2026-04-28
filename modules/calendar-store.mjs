/**
 * Thin wrapper over `messenger.calendar.calendars.*` and
 * `messenger.calendar.items.*`. Always exchanges iCal strings - the
 * wrapper pins `format: "ical"` on every write and `returnFormat: "ical"`
 * on every read so callers never need to think about jCal.
 *
 * Calendar operations tolerate "not found" (the user may have deleted
 * the calendar manually); item-level reads/writes throw on real errors
 * so the sync orchestrator can surface them.
 */

const ICAL_FORMAT = "ical";

const STORAGE_TYPE = "storage";
const STORAGE_URL  = "moz-storage-calendar://";

/* ── Calendar level ───────────────────────────────────────────────── */

/**
 * Create a local storage calendar. `kind` is "events" or "tasks", but
 * the experiment rejects `capabilities` on foreign calendar types
 * ("storage" is foreign from our extension's perspective) - so we
 * don't pass it. Each EAS folder syncs into its own calendar and the
 * sync codec dispatch already constrains what gets written, so the
 * local calendar accepting both kinds is harmless.
 * Returns the new calendar id.
 */
export async function createCalendar({ name, kind, color }) {
  if (!name || typeof name !== "string" || !name.trim()) {
    throw new Error("createCalendar requires a non-empty name");
  }
  if (kind !== "events" && kind !== "tasks") {
    throw new Error(`createCalendar requires kind: 'events' | 'tasks' (got ${kind})`);
  }
  const props = {
    name: name.trim(),
    type: STORAGE_TYPE,
    url: STORAGE_URL,
  };
  if (color) props.color = color;
  const calendar = await messenger.calendar.calendars.create(props);
  return calendar?.id ?? calendar;
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
    id, type,
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
