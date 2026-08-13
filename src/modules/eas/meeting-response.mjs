/**
 * EAS MeetingResponse — tells the server the user Accepted / Tentatively
 * accepted / Declined a meeting invite, by referencing the existing
 * Calendar-collection item ([MS-ASCMD] §2.2.2.10 / §3.1.5.5).
 *
 * This is the protocol-correct replacement for pushing the itip-driven
 * self-attendee PARTSTAT edit back as a plain `Sync <Change>`: the server
 * updates the organizer's copy and notifies attendees itself, so the
 * client must not *also* push a generic item edit for the same change. The
 * push phase decides that: an item whose meeting somebody else organises is
 * never sent as an Add or a Change, and its queue entry is answered here
 * instead.
 *
 * Contributed by Tomas Kovacik <kovacik@dgtfactory.com> in PR #339 and
 * validated against Exchange Online on 16.1. Ported unchanged apart from
 * this note; what surrounds it - when to send, and never waiting for the
 * server to apply it - is ours.
 *
 * Wire shape:
 *
 *   <MeetingResponse>
 *     <Request>
 *       <UserResponse>1|2|3</UserResponse>
 *       <airsync:CollectionId>…</airsync:CollectionId>
 *       <airsync:RequestId>…</airsync:RequestId>
 *       [<InstanceId>…</InstanceId>]
 *     </Request>
 *   </MeetingResponse>
 *
 * Response: `<MeetingResponse><Result><RequestId/><Status/>
 * [<CalendarId/>]</Result></MeetingResponse>`. `CalendarId` is only present
 * when the server created/moved an item into the Calendar folder as a
 * result of the response; the caller's next regular pull sync reconciles
 * that, so it's surfaced but not acted on here.
 */

import { createWBXML } from "../wbxml.mjs";
import { easRequest } from "../network.mjs";
import { readPathFrom } from "./wbxml-helpers.mjs";

function buildBody({ collectionId, serverID, userResponse, instanceId }) {
  const w = createWBXML();
  w.switchpage("MeetingResponse");
  w.otag("MeetingResponse");
  w.otag("Request");
  w.atag("UserResponse", String(userResponse));
  // CollectionId/RequestId are native tokens of the MeetingResponse
  // codepage itself (0x06/0x08) - distinct from AirSync's own
  // CollectionId (0x12) and from RequestId, which AirSync doesn't
  // define at all. No switchpage needed/wanted here.
  w.atag("CollectionId", collectionId);
  w.atag("RequestId", serverID);
  // A whole recurring series is one EAS item, so responding to a single
  // occurrence names that instance rather than addressing a different item.
  // InstanceId follows RequestId in the schema and is valid from 14.1 on.
  if (instanceId) w.atag("InstanceId", instanceId);
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** Send a MeetingResponse for a single item. Returns `{ status, calendarId }`
 *  on any response with a `<Result>` element (caller checks `status === "1"`
 *  for success), or `null` on a network/transport failure or a malformed
 *  response with no `<Result>`. Callers gate on
 *  `easCommandLikelyAvailable(account, "MeetingResponse")` before calling. */
export async function sendMeetingResponse({
  account,
  asVersion,
  collectionId,
  serverID,
  userResponse,
  instanceId = null,
}) {
  if (!collectionId || !serverID || !userResponse) return null;
  // MeetingResponseRequest.xsd restricts InstanceId to exactly 24 characters
  // (the extended form 2026-08-10T07:45:00.000Z) - unlike the AirSyncBase
  // InstanceId, which is the 16-character basic form. Getting it wrong is
  // answered with Status 2, so refuse to send a malformed value rather than
  // have the server reject the response as an invalid meeting.
  if (instanceId && String(instanceId).length !== 24) return null;
  let resp;
  try {
    resp = await easRequest({
      account,
      command: "MeetingResponse",
      body: buildBody({ collectionId, serverID, userResponse, instanceId }),
      asVersion,
    });
  } catch {
    return null;
  }
  if (!resp?.doc) return null;

  const resultNode = resp.doc.getElementsByTagName("Result")[0];
  if (!resultNode) return null;
  const status = readPathFrom(resultNode, ["Status"]);
  const calendarId = readPathFrom(resultNode, ["CalendarId"]);
  return { status, calendarId };
}
