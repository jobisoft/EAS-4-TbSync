/**
 * Thin wrapper over `messenger.addressBooks.*` and `messenger.contacts.*`.
 * Book operations tolerate "not found" (the user may have removed the book
 * manually). Contact writes all take a vCard string via `{ vCard }`.
 */

/** Create a book and return its id. */
export async function createBook(name) {
  if (!name || typeof name !== "string" || !name.trim()) {
    throw new Error("createBook requires a non-empty name");
  }
  const id = await messenger.addressBooks.create({ name: name.trim() });
  // Workaround for TB bug in ext-addressBook.js:707-710 where the
  // contact-node cache misses property updates until contacts.list()
  // has been called at least once for the parent book. Without this
  // seed, the host's content-hash gate sees stale vCards on the
  // first edit and treats real changes as ghosts.
  await messenger.contacts.list(id);
  return id;
}

/** Delete a book, tolerating "not found". */
export async function deleteBook(id) {
  if (!id) return;
  try {
    await messenger.addressBooks.delete(id);
  } catch (err) {
    if (isNotFoundError(err)) return;
    throw err;
  }
}

export async function bookExists(id) {
  if (!id) return false;
  try {
    const node = await messenger.addressBooks.get(id);
    return !!node;
  } catch {
    return false;
  }
}

// ── Contact-level ──────────────────────────────────────────────────────────

/** List all contacts in a book, with vCard normalised to the top level. */
export async function listContacts(bookId) {
  if (!bookId) return [];
  try {
    const list = await messenger.contacts.list(bookId);
    return list.map(normalizeCard);
  } catch (err) {
    if (isNotFoundError(err)) return [];
    throw err;
  }
}

/** Fetch a contact by id with vCard normalised. Null on "not found". */
export async function getContact(id) {
  if (!id) return null;
  try {
    const node = await messenger.contacts.get(id);
    return node ? normalizeCard(node) : null;
  } catch (err) {
    if (isNotFoundError(err)) return null;
    throw err;
  }
}

/** Promote `properties.vCard` (older Thunderbird versions) to a top-level field. */
function normalizeCard(node) {
  if (!node) return node;
  const vCard = node.vCard ?? node.properties?.vCard ?? null;
  return { ...node, vCard };
}

/** Create a contact from a vCard. Returns the new id. */
export async function createContact(bookId, vCard) {
  if (!bookId) throw new Error("createContact requires a bookId");
  if (!vCard) throw new Error("createContact requires a vCard string");
  return await messenger.contacts.create(bookId, { vCard });
}

/** Replace an existing contact's vCard. */
export async function updateContact(contactId, vCard) {
  if (!contactId) throw new Error("updateContact requires a contactId");
  if (!vCard) throw new Error("updateContact requires a vCard string");
  await messenger.contacts.update(contactId, { vCard });
}

/** Delete a contact by id, tolerating "not found". */
export async function deleteContact(contactId) {
  if (!contactId) return;
  try {
    await messenger.contacts.delete(contactId);
  } catch (err) {
    if (isNotFoundError(err)) return;
    throw err;
  }
}

// ── Mailing-list-level ────────────────────────────────────────────────────

/** List all mailing lists in a book. */
export async function listMailingLists(bookId) {
  if (!bookId) return [];
  try {
    return await messenger.mailingLists.list(bookId);
  } catch (err) {
    if (isNotFoundError(err)) return [];
    throw err;
  }
}

/** Fetch a mailing list by id; null on "not found". */
export async function getMailingList(id) {
  if (!id) return null;
  try {
    return await messenger.mailingLists.get(id);
  } catch (err) {
    if (isNotFoundError(err)) return null;
    throw err;
  }
}

/** Create a mailing list. Returns the new id. */
export async function createMailingList(bookId, { name }) {
  if (!bookId) throw new Error("createMailingList requires a bookId");
  if (!name) throw new Error("createMailingList requires a name");
  return await messenger.mailingLists.create(bookId, { name });
}

/** Rename / update a mailing list. */
export async function updateMailingList(id, { name }) {
  if (!id) throw new Error("updateMailingList requires an id");
  await messenger.mailingLists.update(id, { name });
}

/** Delete a mailing list, tolerating "not found". */
export async function deleteMailingList(id) {
  if (!id) return;
  try {
    await messenger.mailingLists.delete(id);
  } catch (err) {
    if (isNotFoundError(err)) return;
    throw err;
  }
}

/** List the contacts in a mailing list, tolerating "not found". */
export async function listMailingListMembers(listId) {
  if (!listId) return [];
  try {
    return await messenger.mailingLists.listMembers(listId);
  } catch (err) {
    if (isNotFoundError(err)) return [];
    throw err;
  }
}

/** Add a contact to a mailing list, tolerating "not found". */
export async function addMailingListMember(listId, contactId) {
  if (!listId || !contactId) return;
  try {
    await messenger.mailingLists.addMember(listId, contactId);
  } catch (err) {
    if (isNotFoundError(err)) return;
    throw err;
  }
}

/** Remove a contact from a mailing list, tolerating "not found". */
export async function removeMailingListMember(listId, contactId) {
  if (!listId || !contactId) return;
  try {
    await messenger.mailingLists.removeMember(listId, contactId);
  } catch (err) {
    if (isNotFoundError(err)) return;
    throw err;
  }
}

/** True for Thunderbird's "there is no such id" error, and nothing else.
 *
 *  Matching on prose is not a choice we get to make: these APIs throw a bare
 *  `ExtensionError`, and only its message survives the boundary - no code, no
 *  name, no subclass. So the message is the whole signal.
 *
 *  What we can choose is to match it exactly. `ext-addressBook.js` raises this
 *  from three places - `findAddressBookById`, `findMailingListById`,
 *  `findContactById` - all with one shape, so the pattern is that shape and
 *  not a bag of keywords. The previous `/no such|not found|invalid id/`
 *  matched **none** of them (Thunderbird says "could *not be* found"), which
 *  made every tolerant branch in this file throw instead of returning null.
 *  That is how one mailing list could fail an entire contacts sync: the push
 *  path already skips an item it cannot read, but the read threw before it
 *  got the chance.
 *
 *  Erring tight is deliberate. Too loose swallows a real defect - "Invalid
 *  vCard data", "The card's UID may not be changed", "Duplicate contact id"
 *  are all our own bug, and losing them silently loses data. Too tight fails
 *  a sync loudly, which is recoverable and gets reported. If a future
 *  Thunderbird rewords this, we want the loud one. */
const NOT_FOUND =
  /^(addressBook|mailingList|contact) with id=.* could not be found\.$/;

function isNotFoundError(err) {
  return NOT_FOUND.test(String(err?.message ?? err ?? ""));
}
