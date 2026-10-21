/** Locate a source meeting-request email without synchronising mail.
 * MS-ASCMD 2.2.3.98.2 permits its Search LongId in MeetingResponse.
 * MS-ASEMAIL 2.2.2.48 supplies Calendar UID on 16.x only. A subject is
 * merely a bounded search hint; it is never accepted as meeting identity. */
import ICAL from "../../vendor/ical.min.js";
import { createWBXML } from "../wbxml.mjs";
import { easRequest } from "../network.mjs";
import { easCommandAdvertised } from "./allowed-commands.mjs";
import { childByTag, readPathFrom } from "./wbxml-helpers.mjs";

export function buildMeetingRequestSearchBody(summary) {
  const w = createWBXML();
  w.switchpage("Search");
  w.otag("Search");
  w.otag("Store");
  w.atag("Name", "Mailbox");
  w.otag("Query");
  w.otag("And");
  w.switchpage("AirSync");
  w.atag("Class", "Email");
  w.switchpage("Search");
  w.atag("FreeText", summary);
  w.ctag();
  w.ctag();
  w.otag("Options");
  w.atag("RebuildResults");
  w.atag("DeepTraversal");
  w.atag("Range", "0-99");
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

function address(value) {
  const text = String(value ?? "").trim();
  const angle = /<([^<>]+)>$/.exec(text);
  return (angle?.[1] ?? text).replace(/^mailto:/i, "").toLowerCase();
}

function identityOf(blob) {
  try {
    const vcal = new ICAL.Component(ICAL.parse(blob));
    const event = vcal.getAllSubcomponents("vevent")
      .find((v) => !v.hasProperty("recurrence-id"));
    if (!event) return null;
    const uid = String(event.getFirstPropertyValue("uid") ?? "");
    const summary = String(event.getFirstPropertyValue("summary") ?? "").trim();
    if (!uid || !summary) return null;
    const start = event.getFirstPropertyValue("dtstart");
    // Floating and unresolved TZID times have no trustworthy UTC instant.
    const startMs = start instanceof ICAL.Time && !start.isDate &&
      start.zone?.tzid !== "floating"
      ? start.toUnixTime() * 1000 : null;
    return { uid, summary, organizer: address(event.getFirstPropertyValue("organizer")), startMs };
  } catch {
    return null;
  }
}

/** Return one unambiguous, answerable source email's LongId, or null.
 * The decoder drops namespace prefixes, so every lookup is anchored to
 * the exact Properties/MeetingRequest path. Incomplete results cannot
 * establish uniqueness and are deliberately left for the next sync. */
export function readMeetingRequestLongId(doc, blob) {
  const identity = identityOf(blob);
  if (!identity) return null;
  const root = doc?.documentElement;
  if (root?.tagName !== "Search" || readPathFrom(root, ["Status"]) !== "1") return null;
  const store = childByTag(childByTag(root, "Response"), "Store");
  if (readPathFrom(store, ["Status"]) !== "1") return null;
  const rows = Array.from(store?.children ?? []).filter((n) => n.tagName === "Result");
  const totalText = readPathFrom(store, ["Total"]);
  if (!/^\d+$/.test(totalText ?? "")) return null;
  const total = Number(totalText);
  if (total > 100 || total !== rows.length) return null;
  const uidRows = rows.filter((row) =>
    readPathFrom(row, ["Properties", "MeetingRequest", "UID"]) === identity.uid);
  // More than one source for the same UID can be an older invitation
  // and its update. Do not select one merely because its optional fields
  // happen to match the local copy; answering the stale request is unsafe.
  const sourceIds = new Set(uidRows.map((row) => readPathFrom(row, ["LongId"])));
  if (sourceIds.size !== 1) return null;
  const matches = new Set();
  for (const row of uidRows) {
    const props = childByTag(row, "Properties");
    const meeting = childByTag(props, "MeetingRequest");
    const itemClass = readPathFrom(row, ["Class"]);
    if (itemClass && itemClass !== "Email") continue;
    if (readPathFrom(props, ["MessageClass"]) !== "IPM.Schedule.Meeting.Request") continue;
    // MS-ASCMD 3.1.5.6: only types 1/2 can be answered. In particular,
    // forwarded/non-owner/delegated notifications (types 5/6) must not be.
    if (!["1", "2"].includes(readPathFrom(meeting, ["MeetingMessageType"]))) continue;
    const organizer = readPathFrom(meeting, ["Organizer"]);
    if (organizer && identity.organizer && address(organizer) !== identity.organizer) continue;
    const start = readPathFrom(meeting, ["StartTime"]);
    if (start && identity.startMs !== null && Date.parse(start) !== identity.startMs) continue;
    const longId = readPathFrom(row, ["LongId"]);
    if (!longId || longId.length > 256) continue;
    matches.add(longId);
  }
  return matches.size === 1 ? [...matches][0] : null;
}

export async function findMeetingRequestLongId({ account, asVersion, blob }) {
  if (!["16.0", "16.1"].includes(asVersion) || !easCommandAdvertised(account, "Search")) return null;
  const identity = identityOf(blob);
  if (!identity) return null;
  try {
    const { doc } = await easRequest({
      account, asVersion, command: "Search",
      body: buildMeetingRequestSearchBody(identity.summary.slice(0, 256)),
    });
    return readMeetingRequestLongId(doc, blob);
  } catch {
    // Search is only a fallback; a transport failure must leave the
    // original invitation and its queued answer available for retry.
    return null;
  }
}
