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

function buildBody({ collectionId, serverID }) {
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
  w.atag("Type", "1");
  w.ctag();
  w.switchpage("ItemOperations");
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** Fetch the server's current `<Properties>` (same shape as
 *  `<ApplicationData>`) for `(collectionId, serverID)`. Returns the
 *  `<Properties>` element on Status 1, or `null` on any other status /
 *  network-level failure. Callers gate on
 *  `easCommandLikelyAvailable(account, "ItemOperations")` before calling. */
export async function fetchServerItem({
  account,
  asVersion,
  collectionId,
  serverID,
}) {
  if (!collectionId || !serverID) return null;
  let resp;
  try {
    resp = await easRequest({
      account,
      command: "ItemOperations",
      body: buildBody({ collectionId, serverID }),
      asVersion,
    });
  } catch {
    return null;
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
