/**
 * EAS `Settings/Oof` - the out-of-office reply, read and written.
 *
 * The first of what the host calls a *service*: a setting that lives on the
 * server rather than in TbSync, reachable only while the account is
 * connected. Everything in the account's config popup is the opposite -
 * stored here, and locked while connected.
 *
 * Three audiences, each with its own switch and its own text: people inside
 * the organisation, outside contacts the mailbox knows, and everyone else.
 * They stay separate on the wire, whatever a dialog chooses to show. A
 * server need not offer all three - one it does not name is *absent* rather
 * than switched off, and only what it named is reported and sent back.
 *
 * `OofState` is 0 disabled, 1 enabled, 2 enabled between two times. The
 * times only mean anything in state 2, and Exchange ignores them otherwise,
 * so they are sent only then.
 *
 * Body type is Text throughout. Exchange will hand back HTML if asked, and
 * Thunderbird has no rich editor to put in a popup like this, so asking for
 * HTML would mean showing the user markup and saving it back as if they had
 * written it. The cost is that saving flattens a message the mailbox holds
 * as HTML, so `readOofFromDoc` reports the server's own `BodyType` and the
 * dialog warns before it happens.
 *
 * What a server does with a message it is not currently sending is its own
 * affair, and they differ: a 16.1 mailbox stops reporting the internal
 * reply the moment out of office is off, a 14.1 one keeps handing it back.
 * Nothing here can change that, so nothing here tries - it only makes sure
 * we never contribute to the loss by sending an empty message of our own.
 */

import { ERR, withCode } from "../../vendor/tbsync/provider.mjs";
import { createWBXML } from "../wbxml.mjs";
import { EasHttpError, NET_ERR, easRequest } from "../network.mjs";
import { childByTag, readPath, readPathFrom } from "./wbxml-helpers.mjs";

const PROVISION_REQUIRED_STATUSES = new Set(["141", "142", "143", "144"]);

/** Wire tag → the audience name this module and the dialog use. */
const AUDIENCES = Object.freeze({
  internal: "AppliesToInternal",
  externalKnown: "AppliesToExternalKnown",
  externalUnknown: "AppliesToExternalUnknown",
});

/** The `<Oof>` block of a decoded reply. A `Document` is not its own root
 *  element, so the walk has to start at `documentElement` - the same
 *  anchoring `readPath` does. */
function oofOf(doc) {
  return childByTag(doc?.documentElement, "Oof");
}

function buildGetBody() {
  const w = createWBXML();
  w.switchpage("Settings");
  w.otag("Settings");
  w.otag("Oof");
  w.otag("Get");
  w.atag("BodyType", "Text");
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** Exported for the unit test: which audiences carry their text, and when,
 *  is the part of this module that can lose a user's message. */
export function buildSetBody({ state, startTime, endTime, messages }) {
  const w = createWBXML();
  w.switchpage("Settings");
  w.otag("Settings");
  w.otag("Oof");
  w.otag("Set");
  w.atag("OofState", String(state));
  // Only meaningful in state 2, and a server that is handed a window it
  // was not asked for is being told something the user did not say.
  if (String(state) === "2" && startTime && endTime) {
    w.atag("StartTime", startTime);
    w.atag("EndTime", endTime);
  }
  for (const [name, tag] of Object.entries(AUDIENCES)) {
    const message = messages?.[name];
    if (!message) continue;
    w.otag("OofMessage");
    w.atag(tag);
    w.atag("Enabled", message.enabled ? "1" : "0");
    // The text rides along whenever we have it, switched off or not: a
    // 16.1 mailbox will not report the internal reply while out of office
    // is off, so sending it back is the only way we know it is still what
    // the user wrote.
    //
    // Never empty, and never for an audience we hold no text for. An empty
    // ReplyMessage erases, and an empty box is not evidence of an empty
    // message - the same 16.1 server hides a stored reply behind exactly
    // that. Preserving a message the user cannot see beats honouring a
    // clear they probably did not ask for.
    if (message.reply) {
      w.atag("ReplyMessage", message.reply);
      w.atag("BodyType", "Text");
    }
    w.ctag();
  }
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** Turn a non-1 Status into the shape the caller can act on: a demand to
 *  re-provision is raised the way the sync loops already recognise, and
 *  anything else is an error naming the status. */
function assertOk(status, what) {
  if (status === null || status === "1") return;
  if (PROVISION_REQUIRED_STATUSES.has(status)) {
    throw new EasHttpError(NET_ERR.PROVISION_REQUIRED, 0, {
      message: `Settings/Oof ${what} rejected (Status=${status}), server demands re-Provision`,
    });
  }
  throw withCode(
    new Error(`Settings/Oof ${what} rejected (Status=${status})`),
    ERR.UNKNOWN_COMMAND,
  );
}

/** Parse the `<Get>` reply into the shape the dialog renders. */
export function readOofFromDoc(doc) {
  const get = childByTag(oofOf(doc), "Get");
  if (!get) return null;
  // Only what the server actually reported. An audience it does not name
  // is absent rather than off - Exchange 14.1 answers with
  // AppliesToInternal alone, and takes the other two without complaint
  // while keeping neither, so a window that offered them would be
  // promising something the mailbox never agreed to.
  const messages = {};
  for (const [name, tag] of Object.entries(AUDIENCES)) {
    for (const node of get.children ?? []) {
      if (node.tagName !== "OofMessage") continue;
      if (!childByTag(node, tag)) continue;
      messages[name] = {
        enabled: readPathFrom(node, ["Enabled"]) === "1",
        reply: readPathFrom(node, ["ReplyMessage"]) ?? "",
        // Reported rather than acted on, so the dialog can say when it is
        // about to flatten a message the mailbox holds as HTML.
        bodyType: readPathFrom(node, ["BodyType"]) ?? "Text",
      };
      break;
    }
  }
  return {
    state: readPathFrom(get, ["OofState"]) ?? "0",
    startTime: readPathFrom(get, ["StartTime"]) ?? "",
    endTime: readPathFrom(get, ["EndTime"]) ?? "",
    messages,
  };
}

/** What the mailbox currently says, or null when the server answered with
 *  no Oof block at all. Throws on a refusal, so the dialog can show it. */
export async function readOof({ account, asVersion }) {
  const { doc } = await easRequest({
    account,
    command: "Settings",
    body: buildGetBody(),
    asVersion,
  });
  if (!doc) return null;
  assertOk(readPath(doc, ["Status"]), "Get");
  assertOk(readPathFrom(oofOf(doc), ["Status"]), "Get");
  return readOofFromDoc(doc);
}

/** Write the state and the three messages back. Resolves on success. */
export async function writeOof({ account, asVersion, ...settings }) {
  const { doc } = await easRequest({
    account,
    command: "Settings",
    body: buildSetBody(settings),
    asVersion,
  });
  if (!doc)
    throw withCode(new Error("Empty Settings response"), ERR.UNKNOWN_COMMAND);
  assertOk(readPath(doc, ["Status"]), "Set");
  assertOk(readPathFrom(oofOf(doc), ["Status"]), "Set");
  return null;
}
