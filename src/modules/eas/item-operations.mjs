/**
 * EAS ItemOperations.Fetch — pulls items' full ApplicationData by
 * (CollectionId, ServerId), bypassing the per-folder Sync state machine.
 *
 * Two callers, two shapes:
 *   - the read-only revert re-fetches one item at a time to overwrite the
 *     local copy with the server's canonical one;
 *   - the pull's note upgrade collects every item in a window whose
 *     NativeBodyType says the server holds an HTML note the Sync response
 *     flattened, and fetches them all in ONE request — [MS-ASCMD] allows any
 *     number of `<Fetch>` elements per `<ItemOperations>`.
 *
 * Wire shape ([MS-ASCMD] §2.2.2.10), one `<Fetch>` per item:
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
 * Response: `<ItemOperations><Status>…</Status><Response><Fetch>…
 * <Properties/></Fetch>…</Response></ItemOperations>` — one `<Fetch>` per
 * requested item, each echoing its `<ServerId>`, which is what results are
 * keyed by. `<Properties>` wraps the same per-type fields `<Sync>` puts in
 * `<ApplicationData>`; codecs iterate children regardless of the wrapper tag,
 * so the node goes straight to `codec.applicationDataToBlob`.
 */

import { createWBXML } from "../wbxml.mjs";
import { easRequest } from "../network.mjs";
import { readPathFrom } from "./wbxml-helpers.mjs";

/** One `<ItemOperations>` holding a `<Fetch>` per server id. Exported for
 *  the unit tests; production goes through `fetchServerItems`. */
export function buildFetchBody({ collectionId, serverIDs, bodyType }) {
  const w = createWBXML();
  w.switchpage("ItemOperations");
  w.otag("ItemOperations");
  for (const serverID of serverIDs) {
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
    // different ones: a revert wants the note as plain text, and the note
    // upgrade asks for HTML when NativeBodyType says the server holds a
    // rich note that the Sync response flattened. Sync itself always asks
    // for plain, whatever the class.
    w.atag("Type", bodyType ?? "1");
    w.ctag();
    w.switchpage("ItemOperations");
    w.ctag();
    w.ctag();
  }
  w.ctag();
  return w.getBytes();
}

/** Read a decoded ItemOperations response document into
 *  `{ status, items: Map<serverId, PropertiesElement> }`.
 *
 *  Results are keyed by the `<ServerId>` each `<Fetch>` echoes — response
 *  order is not guaranteed to match request order. When a node carries no
 *  echo and the counts line up, it is keyed positionally instead: a server
 *  that omits the echo works today only because the singular caller took the
 *  first node blindly, and that tolerance must survive the plural form.
 *
 *  A `<Fetch>` with a non-1 Status or no `<Properties>` is simply absent from
 *  the map; its siblings are unaffected. Exported for the unit tests. */
export function readFetchResults(root, serverIDs) {
  const items = new Map();
  if (!root) return { status: null, items };

  const status = readPathFrom(root, ["Status"]);
  if (status && status !== "1") return { status, items };

  const nodes = [];
  collectByTag(root, "Fetch", nodes);
  nodes.forEach((node, i) => {
    const fetchStatus = readPathFrom(node, ["Status"]);
    if (fetchStatus && fetchStatus !== "1") return;
    const properties = readChild(node, "Properties");
    if (!properties) return;
    const echoed = readPathFrom(node, ["ServerId"]);
    const key =
      echoed ?? (nodes.length === serverIDs.length ? serverIDs[i] : null);
    if (key != null) items.set(key, properties);
  });
  return { status: status ?? "1", items };
}

/** Depth-first collection by tag name. A recursive walk over `children`
 *  rather than `getElementsByTagName`, because the latter exists only on a
 *  real Document while the unit tests hand this reader the same lightweight
 *  node shape the capture parser builds - and `children` is the one part of
 *  the contract both share. */
function collectByTag(node, tag, out) {
  if (!node?.children) return;
  for (const c of node.children) {
    if (c.tagName === tag) out.push(c);
    collectByTag(c, tag, out);
  }
}

/** Fetch the server's current `<Properties>` for each of `serverIDs`, in one
 *  request. Returns `{ status, items }`; an id absent from `items` means the
 *  server answered and did not supply that item. A request that failed to
 *  reach the server throws instead — an unanswered question must never look
 *  like an answer (see `fetchServerItem` for why that distinction is
 *  load-bearing). Callers gate on
 *  `easCommandLikelyAvailable(account, "ItemOperations")`. */
export async function fetchServerItems({
  account,
  asVersion,
  collectionId,
  serverIDs,
  bodyType,
}) {
  const wanted = [...new Set(serverIDs)].filter(Boolean);
  if (!collectionId || wanted.length === 0) {
    return { status: null, items: new Map() };
  }
  const resp = await easRequest({
    account,
    command: "ItemOperations",
    body: buildFetchBody({ collectionId, serverIDs: wanted, bodyType }),
    asVersion,
  });
  return readFetchResults(resp?.doc?.documentElement ?? null, wanted);
}

/** Fetch one item's `<Properties>`. Returns the element on Status 1, or
 *  `null` when the server answered and the item is not there. A request that
 *  failed to reach the server throws instead: `revertLocalChanges` reads null
 *  as "the server deleted it" and DELETES the local item — so any failure to
 *  reach the server would wipe the user's pending edit on a read-only folder,
 *  and one throttling window would take every remaining item with it. A
 *  request that never got an answer says nothing about the server's state;
 *  only a served answer can mean "gone". */
export async function fetchServerItem({
  account,
  asVersion,
  collectionId,
  serverID,
  bodyType,
}) {
  if (!collectionId || !serverID) return null;
  const { items } = await fetchServerItems({
    account,
    asVersion,
    collectionId,
    serverIDs: [serverID],
    bodyType,
  });
  // A single-Fetch request can only be answered about that item, so a lone
  // result is ours whatever key it carries - a server echoing the id in a
  // different encoding must not read as "gone", because the revert path
  // deletes the local item on null.
  if (items.size === 1) return items.values().next().value;
  return items.get(serverID) ?? null;
}

function readChild(node, tag) {
  if (!node?.children) return null;
  for (const c of node.children) if (c.tagName === tag) return c;
  return null;
}
