/**
 * Building the `Sync` request body.
 *
 * Split out from the runner so that anything which has to speak to a
 * collection can state its options correctly without importing the runner
 * itself. The Options block is the reason this is one function rather than
 * a rule each call site follows - see `appendOptions`.
 */

import { createWBXML } from "../wbxml.mjs";

/** Cheap pre-filter for the instance phase: does this blob carry anything
 *  `listInstanceCommands` could emit? A false positive costs one no-op
 *  call, a false negative is impossible - both shapes it walks leave one
 *  of these two markers in the iCal text. */
export function blobHasInstanceOverrides(blob) {
  return (
    typeof blob === "string" &&
    (blob.includes("RECURRENCE-ID") || blob.includes("EXDATE"))
  );
}

/**
 * The collection's options, complete, or nothing at all.
 *
 * [MS-ASCMD] keeps the last `<Options>` block per collection and a new one
 * REPLACES it - so every element that should still be active has to appear
 * in every block we write. A block naming one option silently retires the
 * rest, and the server carries that loss until something states them again.
 *
 * That is why this is one function rather than a rule each call site
 * follows. The rule was written down once, in a comment on the push batch,
 * and the same commit that wrote it restored `BodyPreference` while leaving
 * `FilterType` out - so every push reset the calendar window to unfiltered
 * and the next filtered pull removed what it had just added. Measured on
 * Exchange Online 16.1: a push-triggered sync pulled 45 items as adds with
 * a six-month window and none with the window set to "all".
 *
 * A request that writes no options at all is safe and sometimes required -
 * the server keeps what it has. The instance-command request relies on
 * that, and does not come through here.
 */
function appendOptions(w, { asVersion, className, filterType, conflict }) {
  if (asVersion === "2.5") {
    // 2.5 has no Class or BodyPreference inside Options - both arrived in
    // AS 12 - and it keeps the server's conflict default, that branch being
    // minimal-touch by policy and untestable. Calendar is the only class
    // left with a window worth stating, and without it the server treats
    // the initial pull as "every event ever". Matches legacy
    // sync.js:409-412.
    if (className !== "Calendar") return;
    w.otag("Options");
    w.atag("FilterType", String(filterType));
    w.ctag();
    return;
  }

  w.otag("Options");
  // FilterType narrows Calendar pulls to a window (e.g. last 2 weeks).
  // Only meaningful for Calendar - Contacts/Tasks have no time axis.
  // Legacy gates this on `type == "Calendar"` (sync.js:401); we mirror by
  // emitting the tag only for the Calendar class.
  if (className === "Calendar") w.atag("FilterType", String(filterType));
  w.atag("Class", className);
  // The account's conflict preference, on every request that states
  // options at all - including the one carrying Commands, which is the
  // request the server resolves conflicts in.
  if (conflict != null) w.atag("Conflict", conflict);
  w.switchpage("AirSyncBase");
  w.otag("BodyPreference");
  // Plain text, for every class.
  //
  // Asking for HTML looks harmless and is not. Exchange answers a
  // plain-text note by generating an HTML document around it - its own
  // wrapper, its own styles - which we would store in the ALTREP and, on
  // the next push of that item, send straight back as the note. The server
  // takes that as a real edit to the body and bumps the item's version, so
  // an event nobody touched churns on every sync, and an instance command
  // travelling behind its master's <Change> in the same sync is answered
  // with Status 7 (conflict) against our own write. Section 3.4 is where
  // that surfaced.
  //
  // A note the user formats in Thunderbird still goes out as Type 2 -
  // `appendBodyFromDescription` sends the HTML whenever a DESCRIPTION
  // carries an ALTREP - so authored formatting reaches the server. Only
  // formatting the server invented for itself stays out of the ALTREP.
  w.atag("Type", "1");
  w.ctag();
  w.switchpage("AirSync");
  w.ctag();
}

/** Exported for `test/unit/sync-options.test.mjs`, which decodes what this
 *  emits: the options block is worth pinning and there is no other seam
 *  that reaches it. Nothing outside this module calls it in production. */
export function buildSyncBody({
  synckey,
  collectionId,
  asVersion,
  withChanges,
  withCommands,
  withInstanceCommand,
  className,
  filterType,
  windowSize,
  conflict,
}) {
  const options = { asVersion, className, filterType, conflict };
  const w = createWBXML();
  w.switchpage("AirSync");
  w.otag("Sync");
  w.otag("Collections");
  w.otag("Collection");
  if (asVersion === "2.5") w.atag("Class", className);
  w.atag("SyncKey", synckey);
  w.atag("CollectionId", collectionId);
  if (withChanges) {
    w.atag("DeletesAsMoves");
    w.atag("GetChanges");
    w.atag("WindowSize", String(windowSize ?? 25));
    appendOptions(w, options);
  } else if (withCommands || withInstanceCommand) {
    // [MS-ASCMD] 2.2.3.84: "If the client does not want the server changes
    // returned, the request MUST include the GetChanges element with a
    // value of 0" - and when the element is absent with a non-zero SyncKey,
    // "the request is handled as if the GetChanges element were set to 1".
    // So a push that says nothing is a push that asks for changes.
    //
    // It must not. The server answers with a snapshot taken while our own
    // commands are still going out, and applying that snapshot deletes what
    // it does not yet know about. Measured on Exchange Online 16.1: a
    // series was added, its cancellation sent, and the reply to that
    // cancellation carried the master back with an <Exceptions> block
    // holding only the cancellation - truthfully, because the override was
    // still queued behind it. Applying it dropped the override locally; the
    // next request put it on the server, and the two sides disagreed.
    //
    // The pull that follows this push asks properly, once the queue is
    // empty and a snapshot means what it says.
    w.atag("GetChanges", "0");
  }
  // A Commands-only push states them too - it is the request the server
  // resolves conflicts in, and it must not leave the collection holding
  // fewer options than it had. Options precedes Commands in the Collection
  // schema.
  //
  // Instance-command requests are deliberately excluded, and write no
  // options at all: an exception <Change> always follows our own master
  // push in the same sync, and with an explicit <Conflict> Exchange
  // conflict-checks it against exactly that master change and discards it
  // with Status 7 (measured on Exchange Online; without the element the
  // same request is accepted). A conflict verdict against our own change is
  // never wanted - and writing no block leaves the collection's options
  // exactly as the request before it set them.
  //
  // 2.5 is left alone: it never stated options on a push before, its pull
  // states the only one it has, and nothing in a Commands request there can
  // retire it. Minimal-touch on a version we cannot test.
  if (!withChanges && withCommands && asVersion !== "2.5") {
    appendOptions(w, options);
  }
  if (withCommands) appendCommands(w, withCommands);
  if (withInstanceCommand) appendInstanceCommand(w, withInstanceCommand);
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** `<Commands>` holding a single per-instance command - see the instance
 *  phase for why it is never more than one. The descriptor writes the
 *  element itself and leaves the builder on AirSync, same contract as the
 *  codec calls in `appendCommands`. */
export function appendInstanceCommand(w, command) {
  w.otag("Commands");
  command.emit(w);
  w.ctag();
}

function appendCommands(
  w,
  {
    adds,
    mods,
    dels,
    separator,
    asVersion,
    codec,
    defaultTimezone,
    userEmail,
    fallbackOrganizerName,
    eventLog,
  },
) {
  if (!adds.length && !mods.length && !dels.length) return;
  w.otag("Commands");
  for (const a of adds) {
    w.otag("Add");
    w.atag("ClientId", a.clientId);
    w.otag("ApplicationData");
    codec.appendApplicationDataFromBlob({
      builder: w,
      op: "add",
      blob: a.item.blob,
      asVersion,
      separator,
      defaultTimezone,
      // An Add never carries embedded exceptions on ≤14.x - they follow
      // as a <Change> once the ack yields a ServerId (followUpPhase).
      // On 16.1 the writer never embeds them anyway.
      suppressExceptions:
        asVersion !== "16.1" && blobHasInstanceOverrides(a.item.blob),
      userEmail,
      fallbackOrganizerName,
      eventLog,
    });
    w.switchpage("AirSync");
    w.ctag();
    w.ctag();
  }
  for (const m of mods) {
    w.otag("Change");
    w.atag("ServerId", m.serverID);
    w.otag("ApplicationData");
    codec.appendApplicationDataFromBlob({
      builder: w,
      op: "change",
      blob: m.item.blob,
      asVersion,
      separator,
      defaultTimezone,
      userEmail,
      fallbackOrganizerName,
      eventLog,
    });
    w.switchpage("AirSync");
    w.ctag();
    w.ctag();
    // A 16.1 master's exceptions are *not* emitted here. They share this
    // master's ServerId, and Exchange rejects a second command against a
    // ServerId in the same block - the instance phase sends them one
    // request at a time once this push has landed.
  }
  for (const d of dels) {
    w.otag("Delete");
    w.atag("ServerId", d.serverID);
    w.ctag();
  }
  w.ctag();
}
