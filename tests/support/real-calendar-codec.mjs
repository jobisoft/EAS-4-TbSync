/**
 * The same `itemKind.codec` adapter shape `calendar-sync.mjs`'s (private,
 * unexported) `makeCodec()` builds around `calendar-codec.mjs` - kept
 * here so integration tests can drive `sync-runner.mjs` against the
 * real codec logic instead of a stub, without duplicating production's
 * wiring by hand in every test file.
 */
import * as modCodec from "../../src/modules/eas/calendar-codec.mjs";

export function realCalendarCodec() {
  return {
    applicationDataToBlob({
      adNode,
      existingBlob,
      serverID,
      asVersion,
      defaultTimezone,
      syncRecurrence,
      msTodoCompat,
      uid,
      userEmail,
      eventLog,
    }) {
      return modCodec.applicationDataToIcal({
        adNode,
        existingIcal: existingBlob,
        serverID,
        asVersion,
        defaultTimezone,
        syncRecurrence,
        msTodoCompat,
        uid,
        userEmail,
        eventLog,
      });
    },
    appendApplicationDataFromBlob(args) {
      return modCodec.appendApplicationDataFromIcal({
        ...args,
        ical: args.blob,
      });
    },
    readEasServerIdFromBlob: modCodec.readEasServerIdFromIcal,
    stampEasServerId: modCodec.stampEasServerId,
    applyInstanceChange: modCodec.applyInstanceChange,
    applyInstanceDelete: modCodec.applyInstanceDelete,
    appendInstanceChanges: modCodec.appendInstanceChanges,
  };
}
