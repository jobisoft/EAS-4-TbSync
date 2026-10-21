/** Local iTIP identity for invitations stored on EAS as private appointments.
 * The source stays in extension storage, independent of the calendar cache.
 * No mail, network request or calendar write belongs in this module. */
import ICAL from "../../vendor/ical.min.js";
import { isReceivedMeeting, readEasServerIdFromIcal } from "./calendar-codec.mjs";

const MARKER = "x-eas-client-invitation";
const SOURCE_PROPERTIES = [
  "uid", "organizer", "attendee", "sequence", "dtstamp",
  "x-moz-received-sequence", "x-moz-received-dtstamp",
  "x-moz-invited-attendee", MARKER,
];
const KEY_PREFIX = "eas.clientInvitations.";
const invitationLocks = new Map();

/** Serialize a whole read/prepare/write operation for one invitation.
 * This mutex is not reentrant: do not await another operation on the same
 * key from inside its callback. Other invitations remain independent. */
export async function withClientInvitationLock(targetID, itemId, operation) {
  const key = JSON.stringify([targetID, itemId]);
  const previous = invitationLocks.get(key) ?? Promise.resolve();
  const current = previous.catch(() => {}).then(operation);
  invitationLocks.set(key, current);
  try {
    return await current;
  } finally {
    if (invitationLocks.get(key) === current) invitationLocks.delete(key);
  }
}

function parse(blob) {
  try {
    const component = new ICAL.Component(ICAL.parse(blob));
    return component.name === "vcalendar" ? component : null;
  } catch {
    return null;
  }
}

function master(calendar) {
  return calendar?.getAllSubcomponents("vevent")
    .find((event) => !event.hasProperty("recurrence-id")) ?? null;
}

export function isClientInvitation(blob) {
  return String(master(parse(blob))?.getFirstPropertyValue(MARKER) ?? "") === "1";
}

export function prepareClientInvitation({ blob, priorBlob, enabled, userEmail }) {
  const calendar = parse(blob);
  const event = master(calendar);
  if (!event) return blob;
  const prior = master(parse(priorBlob));
  const sameUid = prior && prior.getFirstPropertyValue("uid") === event.getFirstPropertyValue("uid");
  if ((sameUid && isClientInvitation(priorBlob)) ||
      (enabled && isReceivedMeeting(blob, userEmail) && !readEasServerIdFromIcal(blob))) {
    event.updatePropertyWithValue(MARKER, "1");
    return calendar.toString();
  }
  return blob;
}

/** Return a separate outbound projection; never overwrite the source ICS. */
export function toClientAppointment(blob) {
  if (!isClientInvitation(blob)) return blob;
  const calendar = parse(blob);
  for (const event of calendar.getAllSubcomponents("vevent")) {
    for (const name of ["organizer", "attendee", "x-moz-invited-attendee"]) {
      event.removeAllProperties(name);
    }
  }
  return calendar.toString();
}

function recurrenceKey(event) {
  const rid = event.getFirstPropertyValue("recurrence-id");
  if (!rid) return "master";
  if (rid instanceof ICAL.Time) {
    if (rid.isDate) return `date:${rid.toString()}`;
    if (rid.zone?.tzid !== "floating") return `utc:${rid.toUnixTime()}`;
  }
  return `local:${rid.toString()}`;
}

/** A provider hook may receive only the edited occurrence. Merge it onto
 * its marked parent so saving it cannot replace the entire series. An
 * unrelated or unavailable parent is not a trustworthy base. */
export function mergeClientInvitationOccurrence({ blob, masterBlob }) {
  const incoming = parse(blob);
  if (!incoming || master(incoming)) return blob;
  const overrides = incoming.getAllSubcomponents("vevent");
  if (!overrides.length || overrides.some((event) => !event.hasProperty("recurrence-id"))) return blob;
  if (!isClientInvitation(masterBlob)) return blob;
  const merged = parse(masterBlob);
  const uid = master(merged)?.getFirstPropertyValue("uid");
  if (!uid || overrides.some((event) => event.getFirstPropertyValue("uid") !== uid)) return blob;
  for (const override of overrides) {
    const key = recurrenceKey(override);
    for (const old of merged.getAllSubcomponents("vevent")) {
      if (old.hasProperty("recurrence-id") && recurrenceKey(old) === key) merged.removeSubcomponent(old);
    }
    merged.addSubcomponent(new ICAL.Component(structuredClone(override.toJSON())));
  }
  // A new override can use a zone absent from the original invitation.
  // Keep existing definitions so other occurrences keep their meaning.
  const zoneIds = new Set(merged.getAllSubcomponents("vtimezone")
    .map((zone) => zone.getFirstPropertyValue("tzid")));
  for (const zone of incoming.getAllSubcomponents("vtimezone")) {
    const id = zone.getFirstPropertyValue("tzid");
    if (id && !zoneIds.has(id)) {
      merged.addSubcomponent(new ICAL.Component(structuredClone(zone.toJSON())));
      zoneIds.add(id);
    }
  }
  return merged.toString();
}

/** Restore only invitation identity and revision data onto server content.
 * Iterate the built components: a server-deleted override stays deleted. */
export function restoreClientInvitation({ builtBlob, sourceBlob }) {
  if (!isClientInvitation(sourceBlob)) return builtBlob;
  const built = parse(builtBlob);
  const source = parse(sourceBlob);
  if (!built || !source) return builtBlob;
  const originals = new Map(source.getAllSubcomponents("vevent")
    .map((event) => [recurrenceKey(event), event]));
  const originalUid = master(source)?.getFirstPropertyValue("uid");
  for (const event of built.getAllSubcomponents("vevent")) {
    // A new server exception inherits invitation identity from the source
    // master; a source override supplies its more specific revision data.
    if (originalUid) event.updatePropertyWithValue("uid", originalUid);
    const original = originals.get(recurrenceKey(event)) ?? master(source);
    if (!original) continue;
    for (const name of SOURCE_PROPERTIES) {
      event.removeAllProperties(name);
      for (const property of original.getAllProperties(name)) {
        event.addProperty(new ICAL.Property(structuredClone(property.toJSON())));
      }
    }
  }
  return built.toString();
}

// Serialize read/modify/write operations so simultaneous hooks cannot lose
// another item's source. Different calendars have separate storage tables.
const pendingWrites = new Map();

function mutate(targetID, change) {
  const key = KEY_PREFIX + targetID;
  const previous = pendingWrites.get(key) ?? Promise.resolve();
  const operation = previous.catch(() => {}).then(async () => {
    const stored = await browser.storage.local.get({ [key]: {} });
    const table = { ...stored[key] };
    change(table);
    if (Object.keys(table).length) await browser.storage.local.set({ [key]: table });
    else await browser.storage.local.remove(key);
  });
  pendingWrites.set(key, operation);
  return operation.finally(() => {
    if (pendingWrites.get(key) === operation) pendingWrites.delete(key);
  });
}

async function tableFor(targetID) {
  const key = KEY_PREFIX + targetID;
  await pendingWrites.get(key);
  const stored = await browser.storage.local.get({ [key]: {} });
  return stored[key] ?? {};
}

export async function rememberClientInvitation(targetID, itemId, blob, serverID = null) {
  if (!targetID || !itemId || !isClientInvitation(blob)) return;
  await mutate(targetID, (table) => {
    const slot = `item:${itemId}`;
    table[slot] = {
      itemId, blob,
      serverID: serverID || readEasServerIdFromIcal(blob) || table[slot]?.serverID || null,
    };
  });
}

export async function clientInvitationFor(targetID, itemId) {
  if (!targetID || !itemId) return null;
  return (await tableFor(targetID))[`item:${itemId}`] ?? null;
}

export async function clientInvitationByServerId(targetID, serverID) {
  if (!targetID || !serverID) return null;
  const matches = Object.values(await tableFor(targetID)).filter((entry) => entry.serverID === serverID);
  return matches.length === 1 ? matches[0] : null;
}

export async function forgetClientInvitation(targetID, itemId = null) {
  if (!targetID) return;
  await mutate(targetID, (table) => {
    if (itemId !== null) delete table[`item:${itemId}`];
    else for (const slot of Object.keys(table)) delete table[slot];
  });
}
