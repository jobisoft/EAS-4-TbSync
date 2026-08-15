/**
 * The meeting response e-mail, and the `SendMail` command that carries it.
 *
 * [MS-ASCMD] *Receiving and Accepting Meeting Requests*, step 5: after a
 * successful `MeetingResponse` the client sends the reply to the organiser
 * itself, and that step "applies only to protocol versions 2.5, 12.0, 12.1,
 * 14.0, and 14.1". From 16.0 the server generates it. So on 14.x nothing
 * reaches the organiser unless we send it - `Status 1` only means the
 * user's own calendar was updated.
 *
 * Only after the response succeeded, never before: the guidance is explicit
 * that doing it the other way round leaves the invitee's calendar and what
 * the organiser has been told disagreeing with each other.
 *
 * The message is what Outlook sends: a subject that states the answer, and
 * an iCalendar REPLY naming the responder with their PARTSTAT. The
 * organiser's server reads that part, not the prose, and turns it into the
 * AttendeeStatus the organiser sees.
 */

import ICAL from "../../vendor/ical.min.js";
import { createWBXML } from "../wbxml.mjs";
import { easRequest } from "../network.mjs";
import { readPath } from "./wbxml-helpers.mjs";

/** UserResponse → the words and the PARTSTAT that go with it.
 *  [MS-ASCMD] gives the subject prefixes; the PARTSTAT is what the
 *  organiser's server actually acts on. */
const RESPONSE_SHAPE = Object.freeze({
  1: { prefix: "Accepted", partstat: "ACCEPTED" },
  2: { prefix: "Tentative", partstat: "TENTATIVE" },
  3: { prefix: "Declined", partstat: "DECLINED" },
});

const CRLF = "\r\n";

/** Fold nothing, escape nothing: a header value that could carry a newline
 *  would let a caller inject headers, and a meeting subject is the one
 *  field here that comes from another person. */
function headerSafe(text) {
  return String(text ?? "")
    .replace(/[\r\n]+/g, " ")
    .trim();
}

/** `Name <addr>` when we have a name, bare address otherwise. */
function addressOf(name, email) {
  const clean = headerSafe(name);
  return clean ? `${clean} <${email}>` : email;
}

function icalTimestamp(date) {
  return (
    date
      .toISOString()
      .replace(/[-:]/g, "")
      .replace(/\.\d{3}/, "")
      .slice(0, 15) + "Z"
  );
}

/**
 * The reply message for one answered meeting, as raw MIME, or null when the
 * item does not carry what a reply needs.
 *
 * `now` is passed in rather than read, so the DTSTAMP is testable.
 */
export function buildMeetingResponseMime({
  blob,
  userResponse,
  userEmail,
  userName = "",
  now,
}) {
  const shape = RESPONSE_SHAPE[userResponse];
  if (!shape || !blob || !userEmail) return null;

  let vevent;
  try {
    const vcal = new ICAL.Component(ICAL.parse(blob));
    vevent =
      vcal
        .getAllSubcomponents("vevent")
        .find((v) => !v.getFirstProperty("recurrence-id")) ??
      vcal.getFirstSubcomponent("vevent");
  } catch {
    return null;
  }
  if (!vevent) return null;

  const uid = vevent.getFirstPropertyValue("uid");
  const organizerProp = vevent.getFirstProperty("organizer");
  const organizer = String(organizerProp?.getFirstValue() ?? "").replace(
    /^mailto:/i,
    "",
  );
  // No organiser means nobody to tell. That is not a failure - a meeting we
  // hold without one is not one we were invited to.
  if (!uid || !organizer) return null;

  const summary = headerSafe(vevent.getFirstPropertyValue("summary") ?? "");
  const organizerName = headerSafe(organizerProp?.getParameter("cn") ?? "");
  const stamp = icalTimestamp(now);
  const boundary = `eas-reply-${uid}`
    .replace(/[^A-Za-z0-9-]/g, "")
    .slice(0, 60);

  const ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//TbSync//EAS-4-TbSync//EN",
    "METHOD:REPLY",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${stamp}`,
    `ORGANIZER:mailto:${organizer}`,
    `ATTENDEE;PARTSTAT=${shape.partstat}:mailto:${userEmail}`,
    ...(summary ? [`SUMMARY:${shape.prefix}: ${summary}`] : []),
    "END:VEVENT",
    "END:VCALENDAR",
  ].join(CRLF);

  return [
    `From: ${addressOf(userName, userEmail)}`,
    `To: ${addressOf(organizerName, organizer)}`,
    `Subject: ${shape.prefix}${summary ? `: ${summary}` : ""}`,
    "MIME-Version: 1.0",
    `Content-Type: multipart/alternative; boundary="${boundary}"`,
    "",
    `--${boundary}`,
    "Content-Type: text/plain; charset=utf-8",
    "",
    // Deliberately bare. Outlook writes the answer into the subject and the
    // iCalendar part, and anything we invent here is text the user did not
    // write appearing over their name.
    `${shape.prefix}: ${summary}`,
    "",
    `--${boundary}`,
    'Content-Type: text/calendar; charset=utf-8; method=REPLY; name="meeting.ics"',
    "",
    ics,
    "",
    `--${boundary}--`,
    "",
  ].join(CRLF);
}

/**
 * Send one message. Resolves with the `Status` the server reported, or null
 * when it answered with an empty body - which is what success looks like
 * for this command ([MS-ASCMD]: "If the message was sent successfully, the
 * server returns an empty response").
 */
export async function sendMail({ account, asVersion, mime, clientId }) {
  const w = createWBXML();
  w.switchpage("ComposeMail");
  w.otag("SendMail");
  w.atag("ClientId", clientId);
  w.atag("SaveInSentItems");
  w.atag("Mime", mime);
  w.ctag();

  const { doc } = await easRequest({
    account,
    command: "SendMail",
    body: w.getBytes(),
    asVersion,
  });
  return doc ? readPath(doc, ["Status"]) : null;
}
