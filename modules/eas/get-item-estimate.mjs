/**
 * EAS GetItemEstimate command. Asks the server for the number of items
 * a subsequent Sync would deliver. Soft-fails when the server doesn't
 * advertise GetItemEstimate or returns a non-success status; the caller
 * proceeds without a precise total in that case.
 *
 * Wire shape differs by AS version (legacy
 * `EAS-4-TbSync/content/includes/network.js:946-1006`):
 *
 *   2.5: <Class>, <CollectionId>, <FilterType>, <SyncKey>
 *   14.0+: <SyncKey>, <CollectionId>, <Options><Class/></Options>
 */

import { easRequest } from "../network.mjs";
import { createWBXML } from "../wbxml.mjs";
import { readPath } from "./wbxml-helpers.mjs";
import { easCommandLikelyAvailable } from "./allowed-commands.mjs";

export async function runGetItemEstimate({
  account, asVersion, collectionId, synckey,
  className = "Contacts", filterType = "0",
}) {
  if (!easCommandLikelyAvailable(account, "GetItemEstimate")) return null;
  if (!collectionId || !synckey || synckey === "0") return null;

  let resp;
  try {
    const body = buildBody({ asVersion, collectionId, synckey, className, filterType });
    resp = await easRequest({ account, command: "GetItemEstimate", body, asVersion });
  } catch {
    return null;
  }
  if (!resp?.doc) return null;

  const status = readPath(resp.doc, ["Response", "Status"]);
  if (status !== "1") return null;

  const estimateNode = resp.doc.getElementsByTagName("Estimate")[0];
  if (!estimateNode) return null;
  const n = Number(estimateNode.textContent);
  return Number.isFinite(n) ? n : null;
}

function buildBody({ asVersion, collectionId, synckey, className, filterType }) {
  const w = createWBXML();
  w.switchpage("GetItemEstimate");
  w.otag("GetItemEstimate");
    w.otag("Collections");
      w.otag("Collection");
        if (asVersion === "2.5") {
          w.atag("Class", className);
          w.atag("CollectionId", collectionId);
          w.switchpage("AirSync");
          w.atag("FilterType", filterType);
          w.atag("SyncKey", synckey);
          w.switchpage("GetItemEstimate");
        } else {
          w.switchpage("AirSync");
          w.atag("SyncKey", synckey);
          w.switchpage("GetItemEstimate");
          w.atag("CollectionId", collectionId);
          w.switchpage("AirSync");
          w.otag("Options");
            w.atag("Class", className);
          w.ctag();
          w.switchpage("GetItemEstimate");
        }
      w.ctag();
    w.ctag();
  w.ctag();
  return w.getBytes();
}
