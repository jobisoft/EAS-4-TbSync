/**
 * EAS FolderSync command. Pulls the server's folder hierarchy and
 * returns the new continuation key plus the list of additions /
 * updates / deletions since the previous sync.
 *
 * Wire shape (after WBXML decode):
 *
 *   <FolderSync>
 *     <Status>1</Status>
 *     <SyncKey>$nextKey</SyncKey>
 *     <Changes>
 *       <Count>n</Count>
 *       <Add><ServerId/><ParentId/><DisplayName/><Type/></Add> …
 *       <Update><ServerId/><ParentId/><DisplayName/><Type/></Update> …
 *       <Delete><ServerId/></Delete> …
 *     </Changes>
 *   </FolderSync>
 *
 * On the very first sync (`SyncKey=0`), only `Add` entries appear -
 * the server treats every folder as new. Subsequent syncs may include
 * any combination.
 *
 * Status 1 = success. Any other status is surfaced as an error; in
 * particular Status 9 ("invalid synchronization key") means the caller
 * should reset `foldersynckey` to 0 and retry - handled at the
 * orchestration layer, not here.
 */

import { ERR, withCode } from "../../vendor/tbsync/provider.mjs";
import { createWBXML } from "../wbxml.mjs";
import { easRequest } from "../network.mjs";
import { readPath, readPathFrom } from "./wbxml-helpers.mjs";

function buildFolderSyncBody(currentKey) {
  const w = createWBXML();
  w.switchpage("FolderHierarchy");
  w.otag("FolderSync");
    w.atag("SyncKey", currentKey);
  w.ctag();
  return w.getBytes();
}

function readFolderEntries(doc, tagName) {
  const result = [];
  const nodes = doc.getElementsByTagName(tagName);
  for (const n of nodes) {
    const serverID = readPathFrom(n, ["ServerId"]);
    if (!serverID) continue;
    result.push({
      serverID,
      parentID: readPathFrom(n, ["ParentId"]) ?? "0",
      displayName: readPathFrom(n, ["DisplayName"]) ?? "",
      type: readPathFrom(n, ["Type"]) ?? "",
    });
  }
  return result;
}

export async function runFolderSync({ account, asVersion }) {
  const currentKey = account.custom?.foldersynckey ?? "0";
  const body = buildFolderSyncBody(currentKey);
  const { doc } = await easRequest({
    account,
    command: "FolderSync",
    body,
    asVersion,
  });
  if (!doc) {
    throw withCode(new Error("Empty FolderSync response"), ERR.UNKNOWN_COMMAND);
  }

  const status = readPath(doc, ["Status"]);
  if (status !== "1") {
    const err = new Error(`FolderSync rejected (Status=${status ?? "missing"})`);
    err.code = ERR.UNKNOWN_COMMAND;
    err.folderSyncStatus = status;
    throw err;
  }

  const synckey = readPath(doc, ["SyncKey"]);
  if (!synckey) {
    throw withCode(new Error("FolderSync response missing SyncKey"), ERR.UNKNOWN_COMMAND);
  }

  return {
    synckey,
    adds:    readFolderEntries(doc, "Add"),
    updates: readFolderEntries(doc, "Update"),
    deletes: readFolderEntries(doc, "Delete"),
  };
}
