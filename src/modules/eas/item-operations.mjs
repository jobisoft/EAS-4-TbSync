/**
 * EAS ItemOperations.Fetch — pulls a single item's full ApplicationData by
 * (CollectionId, ServerId), bypassing the per-folder Sync state machine.
 *
 * Used by the read-only revert path: when a folder is in download-only mode
 * we can't push the user's local edits back, so we re-fetch the server's
 * canonical copy and overwrite the local store. Falls back gracefully when
 * the server didn't advertise ItemOperations (legacy then walks Sync.Fetch
 * or, ultimately, a synckey-reset; the new code keeps the synckey-reset as
 * the only fallback to avoid a third wire path).
 *
 * Wire shape ([MS-ASCMD] §2.2.2.10):
 *
 *   <ItemOperations>
 *     <Fetch>
 *       <Store>Mailbox</Store>
 *       <airsync:CollectionId>…</airsync:CollectionId>
 *       <airsync:ServerId>…</airsync:ServerId>
 *       <Options>
 *         <airsyncbase:BodyPreference>
 *           <airsyncbase:Type>1</airsyncbase:Type>
 *         </airsyncbase:BodyPreference>
 *       </Options>
 *     </Fetch>
 *   </ItemOperations>
 *
 * Response: `<ItemOperations><Status>…</Status><Response><Fetch>…<Properties>
 * <ApplicationData…/></Properties></Fetch></Response></ItemOperations>`. The
 * `<Properties>` element wraps the same per-type fields that `<Sync>` puts
 * inside `<ApplicationData>` — codecs iterate children regardless of the
 * wrapper tag name, so we hand the `<Properties>` node straight to
 * `codec.applicationDataToBlob`.
 */

import { createWBXML } from "../wbxml.mjs";
import { easRequest } from "../network.mjs";
import { readPath, readPathFrom } from "./wbxml-helpers.mjs";

function buildBody({ collectionId, serverID, bodyType }) {
  const w = createWBXML();
  w.switchpage("ItemOperations");
  w.otag("ItemOperations");
  w.otag("Fetch");
  w.atag("Store", "Mailbox");
  w.switchpage("AirSync");
  w.atag("CollectionId", collectionId);
  w.atag("ServerId", serverID);
  w.switchpage("ItemOperations");
  w.otag("Options");
  w.switchpage("AirSyncBase");
  w.otag("BodyPreference");
  // The caller names the format it needs, because the two callers need
  // different ones: a revert wants the note as plain text, and the pull's
  // note resolver asks for HTML when NativeBodyType says the server holds a
  // rich note that the Sync response flattened. Sync itself always asks for
  // plain, whatever the class.
  w.atag("Type", bodyType ?? "1");
  w.ctag();
  w.switchpage("ItemOperations");
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** Fetch the server's current `<Properties>` (same shape as
 *  `<ApplicationData>`) for `(collectionId, serverID)`. Returns the
 *  `<Properties>` element on Status 1, or `null` when the server answered
 *  and the item is not there. A request that failed to reach the server
 *  throws instead: callers read null as "the server deleted it" and act
 *  on that, so an unanswered question must never look like an answer.
 *  Callers gate on
 *  `easCommandLikelyAvailable(account, "ItemOperations")` before calling. */
export async function fetchServerItem({
  account,
  asVersion,
  collectionId,
  serverID,
  bodyType,
}) {
  if (!collectionId || !serverID) return null;
  let resp;
  try {
    resp = await easRequest({
      account,
      command: "ItemOperations",
      body: buildBody({ collectionId, serverID, bodyType }),
      asVersion,
    });
  } catch (err) {
    // "Could not ask" and "the server says it is gone" must not look the
    // same to the caller. `revertLocalChanges` reads null as the latter
    // and DELETES the local item - so any failure to reach the server
    // would wipe the user's pending edit on a read-only folder, and one
    // throttling window would take every remaining item with it. A
    // request that never got an answer says nothing about the server's
    // state: it always throws on, and the sync unwinds with an error the
    // user can see. Only a served answer can mean "gone".
    throw err;
  }
  if (!resp?.doc) return null;

  const topStatus = readPath(resp.doc, ["Status"]);
  if (topStatus && topStatus !== "1") return null;

  // The first <Fetch> under <Response> is ours; we only ever send one.
  const fetchNode = resp.doc.getElementsByTagName("Fetch")[0];
  if (!fetchNode) return null;
  const fetchStatus = readPathFrom(fetchNode, ["Status"]);
  if (fetchStatus && fetchStatus !== "1") return null;
  const properties = readChild(fetchNode, "Properties");
  if (!properties) return null;
  return properties;
}

function readChild(node, tag) {
  if (!node?.children) return null;
  for (const c of node.children) if (c.tagName === tag) return c;
  return null;
}
