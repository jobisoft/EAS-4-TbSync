/**
 * EAS ResolveRecipients command (codepage 10), for the availability of
 * an address the user is about to invite.
 *
 * Request:
 *
 *   <ResolveRecipients>
 *     <To>someone@example.org</To>
 *     <Options>
 *       <Availability>
 *         <StartTime>2026-08-12T09:00:00.000Z</StartTime>
 *         <EndTime>2026-08-12T12:00:00.000Z</EndTime>
 *       </Availability>
 *     </Options>
 *   </ResolveRecipients>
 *
 * Reply:
 *
 *   <ResolveRecipients>
 *     <Status>1</Status>
 *     <Response>
 *       <To>…</To><Status>1</Status><RecipientCount>1</RecipientCount>
 *       <Recipient>
 *         <Type>1</Type><DisplayName>…</DisplayName><EmailAddress>…</EmailAddress>
 *         <Availability>
 *           <Status>1</Status>
 *           <MergedFreeBusy>0000222200000000</MergedFreeBusy>
 *         </Availability>
 *       </Recipient>
 *     </Response>
 *   </ResolveRecipients>
 *
 * `MergedFreeBusy` is one digit per time slot, laid end to end from
 * StartTime: 0 free, 1 tentative, 2 busy, 3 out of office, 4 no data.
 * A slot is 30 minutes; `eas/free-busy.mjs` owns that arithmetic and
 * says why it must not be derived from the reply's own length.
 */

import { createWBXML } from "../wbxml.mjs";
import { easRequest } from "../network.mjs";
import { childByTag, readPathFrom } from "./wbxml-helpers.mjs";

/** The extended ISO form with milliseconds, which is what [MS-ASCMD]
 *  uses for this command - unlike the basic `YYYYMMDDTHHMMSSZ` the
 *  Calendar codepage wants. */
function isoWithMillis(date) {
  return date.toISOString().replace(/\.\d{3}Z$/, ".000Z");
}

function buildBody({ to, start, end }) {
  const w = createWBXML("ResolveRecipients");
  w.otag("ResolveRecipients");
  w.atag("To", to);
  w.otag("Options");
  w.otag("Availability");
  w.atag("StartTime", isoWithMillis(start));
  w.atag("EndTime", isoWithMillis(end));
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** Pull what we asked for out of the reply, tolerating every level of
 *  absence: a recipient the server could not resolve carries no
 *  `Availability` at all, and a refused request carries only the
 *  top-level `Status`. Missing reads as null, never as an error - the
 *  caller decides that no availability simply means nothing to draw. */
export function readResolveRecipients(doc) {
  const root = doc?.documentElement;
  if (!root) return null;
  const response = childByTag(root, "Response");
  const recipient = childByTag(response, "Recipient");
  const availability = childByTag(recipient, "Availability");
  return {
    status: readPathFrom(root, ["Status"]) ?? null,
    responseStatus: readPathFrom(response, ["Status"]) ?? null,
    availabilityStatus: readPathFrom(availability, ["Status"]) ?? null,
    mergedFreeBusy: readPathFrom(availability, ["MergedFreeBusy"]) ?? null,
    displayName: readPathFrom(recipient, ["DisplayName"]) ?? null,
  };
}

/** Ask the server how busy `to` is between `start` and `end`. Returns
 *  the raw digits, so the caller owns the slot arithmetic. */
export async function runResolveRecipients({
  account,
  asVersion,
  to,
  start,
  end,
}) {
  const { doc } = await easRequest({
    account,
    command: "ResolveRecipients",
    body: buildBody({ to, start, end }),
    asVersion,
  });
  return readResolveRecipients(doc);
}
