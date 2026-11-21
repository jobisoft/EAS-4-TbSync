#!/usr/bin/env node
/**
 * A small mock ActiveSync server, for testing what the add-on does when the
 * server changes its folders - which no real account lets a script do.
 *
 * It speaks only as much EAS as a connect and an empty sync of a folder need:
 *
 *   OPTIONS      versions and commands (no Provision, no Settings)
 *   FolderSync   the folder tree, as Add / Update / Delete since a key
 *   Sync         every collection is empty
 *
 * Any credentials are accepted. The wire format is the add-on's own codec
 * (`src/modules/wbxml.mjs`), so a request the mock cannot read is one the
 * codec cannot read either.
 *
 * Control, over plain HTTP under /__mock/ (JSON in, JSON out):
 *
 *   GET  /__mock/folders                      the tree
 *   POST /__mock/rename   {serverId, name?, parentId?}
 *                                             rename a folder, and/or move it
 *   POST /__mock/add      {serverId?, name, type, parentId?}
 *   POST /__mock/delete   {serverId}
 *   GET  /__mock/log                          every request seen so far
 *   POST /__mock/reset                        back to the starting tree
 *
 * Usage:  node mock/server.mjs [port]      (default 47660)
 */

import http from "node:http";
import { decodeWBXML, createWBXML } from "../src/modules/wbxml.mjs";

const PORT = Number(process.argv[2] ?? 47660);

// EAS folder types: 4 deleted items, 7 tasks, 8 calendar, 9 contacts,
// 13 user calendar, 14 user contacts, 15 user tasks.
const START = () => [
  { serverId: "tasks", parentId: "0", name: "Tasks", type: 7 },
  { serverId: "calendar", parentId: "0", name: "Calendar", type: 8 },
  { serverId: "contacts", parentId: "0", name: "Contacts", type: 9 },
  { serverId: "trash", parentId: "0", name: "Deleted Items", type: 4 },
  { serverId: "mycal", parentId: "0", name: "Team calendar", type: 13 },
  { serverId: "mybook", parentId: "0", name: "Customers", type: 14 },
  { serverId: "mytasks", parentId: "0", name: "Projects", type: 15 },
];

let folders = START();
/** What changed at which version of the tree: `{ version, kind, folder }`. */
let changes = [];
let version = 1;
let log = [];
/** Per-collection Sync keys, so each Sync moves them on. */
const collectionKeys = new Map();

function bump(kind, folder) {
  version += 1;
  changes.push({ version, kind, folder: { ...folder } });
}

// ── The wire ───────────────────────────────────────────────────────────────

function text(xml, tag) {
  const m = xml.match(new RegExp(`<${tag}>([^<]*)</${tag}>`));
  return m ? decodeURIComponent(m[1]) : null;
}

function folderSyncResponse(clientKey) {
  const w = createWBXML();
  w.switchpage("FolderHierarchy");
  w.otag("FolderSync");

  const known = clientKey === "0" || /^\d+$/.test(clientKey);
  if (!known || Number(clientKey) > version) {
    w.atag("Status", "9"); // invalid synchronization key
    w.ctag();
    return w.getBytes();
  }

  let entries;
  if (clientKey === "0") {
    entries = folders.map((folder) => ({ kind: "Add", folder }));
  } else {
    entries = changes
      .filter((c) => c.version > Number(clientKey))
      .map((c) => ({ kind: c.kind, folder: c.folder }));
  }

  w.atag("Status", "1");
  w.atag("SyncKey", String(version));
  w.otag("Changes");
  w.atag("Count", String(entries.length));
  for (const { kind, folder } of entries) {
    w.otag(kind);
    w.atag("ServerId", folder.serverId);
    if (kind !== "Delete") {
      w.atag("ParentId", folder.parentId);
      w.atag("DisplayName", folder.name);
      w.atag("Type", String(folder.type));
    }
    w.ctag();
  }
  w.ctag();
  w.ctag();
  return w.getBytes();
}

function syncResponse(xml) {
  const collectionId = text(xml, "CollectionId") ?? "0";
  const sent = text(xml, "SyncKey") ?? "0";
  const next = String(Number(sent) + 1);
  collectionKeys.set(collectionId, next);

  const w = createWBXML();
  w.otag("Sync");
  w.otag("Collections");
  w.otag("Collection");
  w.atag("SyncKey", next);
  w.atag("CollectionId", collectionId);
  w.atag("Status", "1");
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

// ── The server ─────────────────────────────────────────────────────────────

function readBody(req) {
  return new Promise((resolve) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => resolve(Buffer.concat(chunks)));
  });
}

function json(res, status, value) {
  res.writeHead(status, { "Content-Type": "application/json" });
  res.end(JSON.stringify(value, null, 2));
}

async function control(req, res, path) {
  const body = req.method === "POST" ? await readBody(req) : null;
  const args = body?.length ? JSON.parse(body.toString("utf8")) : {};
  const find = () => folders.find((f) => f.serverId === args.serverId);

  if (path === "folders") return json(res, 200, { version, folders });
  if (path === "log") return json(res, 200, log);
  if (path === "reset") {
    folders = START();
    changes = [];
    version = 1;
    log = [];
    collectionKeys.clear();
    return json(res, 200, { ok: true });
  }
  if (path === "rename") {
    const folder = find();
    if (!folder || !(args.name || args.parentId)) {
      return json(res, 404, { error: "no such folder" });
    }
    if (args.name) folder.name = args.name;
    if (args.parentId) folder.parentId = args.parentId;
    bump("Update", folder);
    return json(res, 200, { version, folder });
  }
  if (path === "add") {
    const folder = {
      serverId: args.serverId ?? `f${version + 1}`,
      parentId: args.parentId ?? "0",
      name: args.name,
      type: args.type,
    };
    folders.push(folder);
    bump("Add", folder);
    return json(res, 200, { version, folder });
  }
  if (path === "delete") {
    const folder = find();
    if (!folder) return json(res, 404, { error: "no such folder" });
    folders = folders.filter((f) => f !== folder);
    bump("Delete", folder);
    return json(res, 200, { version });
  }
  return json(res, 404, { error: "unknown control path" });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host}`);
  try {
    if (url.pathname.startsWith("/__mock/")) {
      return await control(req, res, url.pathname.slice("/__mock/".length));
    }

    if (req.method === "OPTIONS") {
      log.push({ at: Date.now(), command: "OPTIONS" });
      res.writeHead(200, {
        "MS-ASProtocolVersions": "14.1",
        "MS-ASProtocolCommands": "FolderSync,Sync",
      });
      return res.end();
    }

    const command = url.searchParams.get("Cmd");
    const raw = await readBody(req);
    const xml = raw.length ? decodeWBXML(new Uint8Array(raw)) : "";
    log.push({ at: Date.now(), command, request: xml });

    let out;
    if (command === "FolderSync") {
      out = folderSyncResponse(text(xml, "SyncKey") ?? "0");
    } else if (command === "Sync") {
      out = syncResponse(xml);
    } else {
      res.writeHead(501);
      return res.end();
    }
    res.writeHead(200, { "Content-Type": "application/vnd.ms-sync.wbxml" });
    res.end(Buffer.from(out));
  } catch (err) {
    log.push({ at: Date.now(), error: String(err?.stack ?? err) });
    res.writeHead(500);
    res.end(String(err?.message ?? err));
  }
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`mock EAS server on http://127.0.0.1:${PORT}`);
});
