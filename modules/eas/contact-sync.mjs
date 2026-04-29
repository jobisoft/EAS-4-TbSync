/**
 * EAS contact-sync entry point. Wires the Contacts itemKind into the
 * shared `runItemSync` framework. Codec lives in `contact-codec.mjs`;
 * local store reads/writes go through `address-book.mjs`.
 */

import { runItemSync } from "./sync-runner.mjs";
import * as addressBook from "../address-book.mjs";
import {
  applicationDataToVCard,
  appendApplicationDataFromVCard,
  readEasServerIdFromVCard,
  stampEasServerId,
} from "./contact-codec.mjs";

const codec = {
  async applicationDataToBlob({ adNode, serverID, asVersion, separator, uid }) {
    return await applicationDataToVCard({
      adNode,
      serverID,
      asVersion,
      separator,
      uid,
    });
  },
  appendApplicationDataFromBlob({ builder, blob, asVersion, separator }) {
    return appendApplicationDataFromVCard({
      builder,
      vCard: blob,
      asVersion,
      separator,
    });
  },
  readEasServerIdFromBlob: readEasServerIdFromVCard,
  stampEasServerId,
};

function contactStoreFactory(targetID) {
  return {
    async list() {
      const all = await addressBook.listContacts(targetID);
      return all.map((c) => ({ id: c.id, blob: c.vCard }));
    },
    async get(id) {
      const c = await addressBook.getContact(id);
      return c ? { id: c.id, blob: c.vCard } : null;
    },
    async create(_id, blob) {
      // The vCard's UID is set by the codec to the pre-assigned UUID, and
      // Thunderbird derives the contact id from that, so the returned id
      // matches `_id`. The runner asserts the match.
      return await addressBook.createContact(targetID, blob);
    },
    async update(id, blob) {
      await addressBook.updateContact(id, blob);
    },
    async delete(id) {
      await addressBook.deleteContact(id);
    },
  };
}

const contactItemKind = {
  className: "Contacts",
  filterType: "0",
  changelogKind: "contact",
  mapField: "contactMap",
  codec,
  storeFactory: contactStoreFactory,
};

export async function syncContactFolder({
  provider,
  account,
  folder,
  accountId,
  folderId,
  asVersion,
}) {
  return runItemSync({
    provider,
    account,
    folder,
    accountId,
    folderId,
    asVersion,
    itemKind: contactItemKind,
  });
}
