/* global ExtensionCommon */

"use strict";

var { MailServices } = ChromeUtils.importESModule(
  "resource:///modules/MailServices.sys.mjs",
);

/** Legacy property-bag keys that EAS4 wrote via `nsIAbCard.setProperty()`
 *  ([legacy/EAS4/content/includes/contactsync.js:385, 391](legacy/EAS4/content/includes/contactsync.js#L385)).
 *  Modern Thunderbird does not surface these through the WebExtension-
 *  visible vCard, so the migration walks them via XPCOM and lifts each
 *  to the matching `X-EAS-*` vCard property the new codec round-trips.
 *  `Children` is the odd one out - legacy stored it without an `EAS-`
 *  prefix. Other items keep the `EAS-` prefix in this list so the
 *  caller can derive `X-EAS-<NAME>` mechanically. */
const LEGACY_PROPERTY_BAG_FIELDS = [
  "EAS-Alias",
  "EAS-WeightedRank",
  "EAS-YomiCompanyName",
  "EAS-YomiFirstName",
  "EAS-YomiLastName",
  "EAS-CompressedRTF",
  "EAS-MMS",
  "EAS-ManagerName",
  "EAS-AssistantName",
  "EAS-Spouse",
  "Children",
];

var LegacyAbProperties = class extends ExtensionCommon.ExtensionAPI {
  getAPI(_context) {
    return {
      LegacyAbProperties: {
        /** Returns one entry per non-list card in the address book.
         *  Each entry carries the card's UID (which legacy EAS4 used as
         *  the EAS ServerId via `card.primaryKey`) plus a `stamps` map
         *  of any legacy property-bag values present. The migration
         *  upgrade reads from this to stamp `X-EAS-SERVERID` on every
         *  card and lift property-bag values to vCard `X-EAS-*`
         *  properties. */
        readEasStamps: async (bookId) => {
          const dir = MailServices.ab.getDirectoryFromUID(bookId);
          if (!dir) {
            throw new Error(`No address book found for UID ${bookId}`);
          }
          const out = [];
          for (const card of dir.childCards) {
            if (card.isMailList) continue;
            const stamps = {};
            for (const name of LEGACY_PROPERTY_BAG_FIELDS) {
              const v = card.getProperty(name, "");
              if (v) stamps[name] = v;
            }
            out.push({ contactId: card.UID, stamps });
          }
          return out;
        },
      },
    };
  }
};
