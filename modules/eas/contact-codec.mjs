/**
 * EAS Contact ⇆ vCard 4.0 codec.
 *
 * Mirrors the legacy `EAS-4-TbSync/content/includes/contactsync.js`
 * mapping between EAS Contacts (and Contacts2 + AirSyncBase Body) WBXML
 * and the per-card vCard string Thunderbird now exposes via the
 * `messenger.contacts` API.
 *
 * Reads use `readPathFrom` from the wbxml-helpers (which decodes the
 * `encodeURIComponent` escape applied by `wbxml.mjs::decodeWBXML`).
 *
 * Writes append Contacts / Contacts2 / AirSyncBase tags onto an
 * already-open `<ApplicationData>` builder. The caller switches back to
 * the AirSync codepage afterwards if it has more siblings to emit.
 *
 * The address-street legacy quirk (multi-line streets joined / split by
 * an ASCII separator code stored in `account.custom.seperator`) is
 * carried over verbatim.
 */

import ICAL from "../../vendor/ical.min.js";
import { readPathFrom, readChildTexts } from "./wbxml-helpers.mjs";

/** vCard property name we use to track the EAS server-side serverId so
 *  pull-update / pull-delete can find the right local card without a
 *  separate id map. */
const X_EAS_SERVERID = "x-eas-serverid";

/** Bag of EAS fields with no clean vCard equivalent. We round-trip them
 *  through `X-EAS-<TAG>` properties so the server's value comes back
 *  unchanged on push. */
const PASS_THROUGH_FIELDS = [
  "Alias",
  "WeightedRank",
  "YomiCompanyName",
  "YomiFirstName",
  "YomiLastName",
  "CompressedRTF",
  "MMS",
  "ManagerName",
  "AssistantName",
  "Spouse",
];

/** Legacy "misuses" four EAS fields to round-trip through TB's standard
 *  Custom 1-4 vCard slots ([contactsync.js:118-119, 132-135](legacy/EAS4/content/includes/contactsync.js#L118-L119)).
 *  TB surfaces `x-custom1`..`x-custom4` as the four user-visible Custom
 *  fields in its card editor, so the user can read/edit these EAS-only
 *  values via the standard UI. We mirror that mapping verbatim. */
const CUSTOM_FIELD_MAP = {
  OfficeLocation: "x-custom1", // Contacts namespace
  CustomerId: "x-custom2", // Contacts2 namespace
  GovernmentId: "x-custom3", // Contacts2 namespace
  AccountName: "x-custom4", // Contacts2 namespace
};

/** EAS phone fields without a TB-supported TEL TYPE. Legacy
 *  ([contactsync.js:124-127, 141](legacy/EAS4/content/includes/contactsync.js#L124-L127))
 *  routes them through TEL with a PascalCase TYPE param and prefixes the
 *  value with the type label so TB's UI (which doesn't surface
 *  non-standard TYPE values) at least shows the user *what kind* of
 *  number this is. We mirror that shape verbatim for round-trip parity.
 *  `BusinessFaxNumber` is included even though `fax` is a standard type,
 *  because TB has only one fax slot — legacy used the prefixed form to
 *  distinguish home fax (plain `TYPE=fax`) from business fax. */
const PREFIXED_PHONES = {
  BusinessFaxNumber: { type: "WorkFax", prefix: "WorkFax: " },
  AssistantPhoneNumber: { type: "Assistant", prefix: "Assistant: " },
  CarPhoneNumber: { type: "Car", prefix: "Car: " },
  RadioPhoneNumber: { type: "Radio", prefix: "Radio: " },
  CompanyMainPhone: { type: "Company", prefix: "Company: " },
};

/* ── Reader: ApplicationData → vCard ───────────────────────────────── */

/** Build a vCard 4.0 string from an EAS `<ApplicationData>` DOM node.
 *
 *    adNode      DOM node positioned at <ApplicationData>
 *    serverID    EAS server-side identifier; written as X-EAS-SERVERID
 *    asVersion   "2.5" | "14.0" | "14.1" | "16.1" - selects body codepage
 *    separator   ASCII char code (string) for multi-line address streets
 *    uid         optional vCard UID to embed (TB derives the contact id from it) */
export async function applicationDataToVCard({
  adNode,
  serverID,
  asVersion,
  separator,
  uid,
}) {
  const comp = newVCard();
  if (uid) comp.addPropertyWithValue("uid", uid);
  comp.addPropertyWithValue(X_EAS_SERVERID, serverID);

  readNames(adNode, comp);
  readFileAs(adNode, comp);
  readDates(adNode, comp);
  await readEmails(adNode, comp);
  readWeb(adNode, comp);
  readPhones(adNode, comp);
  readAddresses(adNode, comp, separator);
  readOrganization(adNode, comp);
  readNote(adNode, comp, asVersion);
  readPicture(adNode, comp);
  readCategories(adNode, comp);
  readNickName(adNode, comp);
  readIMs(adNode, comp);
  readChildren(adNode, comp);
  readPassThroughs(adNode, comp);
  readCustomFields(adNode, comp);
  return comp.toString();
}

/* ── Writer: vCard → ApplicationData WBXML ─────────────────────────── */

/** Append Contacts / Contacts2 / AirSyncBase tags onto an open
 *  `<ApplicationData>` builder. Caller is responsible for switching the
 *  codepage back afterwards if more sibling commands follow. */
export function appendApplicationDataFromVCard({
  builder,
  vCard,
  asVersion,
  separator,
}) {
  const comp = parseVCard(vCard);
  if (!comp) return;

  // Phone bucketing happens once: standard types go to the Contacts
  // page below and CompanyMainPhone goes to the Contacts2 page at the
  // end. Both halves consume the same bucketed result.
  const phoneBuckets = bucketPhones(comp);

  // Contacts-page field order mirrors legacy's `Object.keys` iteration
  // over `map_EAS_properties_to_vCard` (contactsync.js:58-128), then
  // the post-loop pass-through / Categories / Children blocks
  // (contactsync.js:495-528). Legacy's generated output was tuned over
  // years against multiple servers, so byte-identical order is part of
  // the compatibility contract — keep this list in lockstep with the
  // legacy property map.
  builder.switchpage("Contacts");

  writeFileAs(builder, comp);
  writeDates(builder, comp);
  writeNames(builder, comp);
  // Notes are emitted later by writeNote between the two pages.
  writeEmails(builder, comp);
  writeWeb(builder, comp);
  writeOrganization(builder, comp);
  writeStandardPhones(builder, phoneBuckets);
  writeAddresses(builder, comp, separator);
  emitIf(
    builder,
    "OfficeLocation",
    stringOf(comp.getFirstPropertyValue("x-custom1")),
  );
  writePicture(builder, comp);
  writePrefixedPhonesContacts(builder, phoneBuckets);
  writePassThroughs(builder, comp);
  writeCategories(builder, comp);
  writeChildren(builder, comp);

  // Body emission sits between the Contacts and Contacts2 pages
  // (legacy contactsync.js:530-547). For AS 2.5 the body stays on the
  // Contacts page; for AS >= 12.0 it switches to AirSyncBase and back.
  // writeNote always ends on Contacts2 so the block below can emit
  // directly.
  writeNote(builder, comp, asVersion);

  emitIf(
    builder,
    "NickName",
    stringOf(comp.getFirstPropertyValue("nickname")),
  );
  emitIf(
    builder,
    "CustomerId",
    stringOf(comp.getFirstPropertyValue("x-custom2")),
  );
  emitIf(
    builder,
    "GovernmentId",
    stringOf(comp.getFirstPropertyValue("x-custom3")),
  );
  emitIf(
    builder,
    "AccountName",
    stringOf(comp.getFirstPropertyValue("x-custom4")),
  );
  writeIMs(builder, comp);
  // CompanyMainPhone tag lives in the Contacts2 codepage per
  // MS-ASCNTC; legacy emits it here at the tail of the Contacts2 loop
  // (contactsync.js:141, 549-554). Trying to atag it on the Contacts
  // page would throw because the Contacts codepage table doesn't
  // carry CompanyMainPhone.
  emitIf(
    builder,
    "CompanyMainPhone",
    phoneBuckets.prefixed.CompanyMainPhone,
  );
}

/** Read X-EAS-SERVERID off an existing card so pull / push paths can
 *  look up the matching server record without a separate id map. */
export function readEasServerIdFromVCard(vCard) {
  const comp = parseVCard(vCard);
  if (!comp) return null;
  const v = comp.getFirstPropertyValue(X_EAS_SERVERID);
  return v ? String(v) : null;
}

/** Set / replace X-EAS-SERVERID on an existing card. Used after
 *  push-add when the server returns the canonical ServerId. */
export function stampEasServerId(vCard, serverID) {
  const comp = parseVCard(vCard) ?? newVCard();
  comp.removeAllProperties(X_EAS_SERVERID);
  if (serverID) comp.addPropertyWithValue(X_EAS_SERVERID, String(serverID));
  return comp.toString();
}

/* ── Names ──────────────────────────────────────────────────────────── */

function readNames(adNode, comp) {
  const last = readPathFrom(adNode, ["LastName"]) ?? "";
  const first = readPathFrom(adNode, ["FirstName"]) ?? "";
  const middle = readPathFrom(adNode, ["MiddleName"]) ?? "";
  const title = readPathFrom(adNode, ["Title"]) ?? "";
  const suffix = readPathFrom(adNode, ["Suffix"]) ?? "";
  if (last || first || middle || title || suffix) {
    const prop = new ICAL.Property("n", comp);
    prop.setValue([last, first, middle, title, suffix]);
    comp.addProperty(prop);
  }
}

function writeNames(b, comp) {
  const n = comp.getFirstProperty("n");
  if (!n) return;
  const v = n.getFirstValue();
  const arr = Array.isArray(v) ? v : v ? [v] : [];
  emitIf(b, "LastName", arr[0]);
  emitIf(b, "FirstName", arr[1]);
  emitIf(b, "MiddleName", arr[2]);
  emitIf(b, "Title", arr[3]);
  emitIf(b, "Suffix", arr[4]);
}

/* ── FileAs / formatted name ───────────────────────────────────────── */

function readFileAs(adNode, comp) {
  const v = readPathFrom(adNode, ["FileAs"]);
  if (v) comp.updatePropertyWithValue("fn", v);
}

function writeFileAs(b, comp) {
  const v = stringOf(comp.getFirstPropertyValue("fn"));
  if (v) b.atag("FileAs", v);
}

/* ── Birthday / Anniversary ────────────────────────────────────────── */

function readDates(adNode, comp) {
  const bday = readPathFrom(adNode, ["Birthday"]);
  if (bday) comp.updatePropertyWithValue("bday", isoDateOnly(bday));
  const ann = readPathFrom(adNode, ["Anniversary"]);
  if (ann) comp.updatePropertyWithValue("anniversary", isoDateOnly(ann));
}

function writeDates(b, comp) {
  const bday = stringOf(comp.getFirstPropertyValue("bday"));
  if (bday) b.atag("Birthday", toIsoDateTime(bday));
  const ann = stringOf(comp.getFirstPropertyValue("anniversary"));
  if (ann) b.atag("Anniversary", toIsoDateTime(ann));
}

/* ── Emails ────────────────────────────────────────────────────────── */

// EAS encodes Email{N}Address as RFC 5322 mailboxes (e.g.
// `"John Doe" <john@example.com>`). On read we extract the bare address
// so Thunderbird's address book stores `EMAIL:john@example.com`. On
// write we rebuild the mailbox using the contact's FN as display name.

async function readEmails(adNode, comp) {
  for (let i = 0; i < 3; i++) {
    const v = readPathFrom(adNode, [`Email${i + 1}Address`]);
    if (!v) continue;
    const bare = await extractBareEmail(v);
    comp.addPropertyWithValue("email", bare);
  }
}

async function extractBareEmail(raw) {
  try {
    const parsed = await messenger.messengerUtilities.parseMailboxString(raw);
    const first = Array.isArray(parsed) ? parsed[0] : null;
    if (first?.email) return first.email;
  } catch {
    /* fall through */
  }
  return raw;
}

function writeEmails(b, comp) {
  const fn = stringOf(comp.getFirstPropertyValue("fn")).trim();
  const emails = comp
    .getAllProperties("email")
    .map((p) => stringOf(p.getFirstValue()).trim())
    .filter(Boolean);
  for (let i = 0; i < Math.min(emails.length, 3); i++) {
    b.atag(`Email${i + 1}Address`, formatMailboxForServer(emails[i], fn));
  }
}

// Defensive: skip the wrap if the stored value already looks like a
// mailbox. Avoids "Name <Other Name <addr@x>>" frankenmailboxes when an
// external tool has already formatted the EMAIL property, or when legacy
// data carried the unstripped form forward.
const MAILBOX_RE = /<[^<>@\s]+@[^<>@\s]+>\s*$/;

function formatMailboxForServer(value, displayName) {
  if (MAILBOX_RE.test(value)) return value;
  if (!displayName) return value;
  return buildMailbox(displayName, value);
}

function buildMailbox(name, email) {
  // RFC 5322 §3.4: quote the display name when it contains specials.
  // EAS WBXML payloads are UTF-8 native, so no RFC 2047 encoding needed.
  const needsQuote = /[",;:<>@()[\]\\.]/.test(name);
  if (!needsQuote) return `${name} <${email}>`;
  const escaped = name.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  return `"${escaped}" <${email}>`;
}

/* ── WebPage ───────────────────────────────────────────────────────── */

function readWeb(adNode, comp) {
  const v = readPathFrom(adNode, ["WebPage"]);
  if (v) comp.addPropertyWithValue("url", v);
}

function writeWeb(b, comp) {
  const v = stringOf(comp.getFirstPropertyValue("url"));
  if (v) b.atag("WebPage", v);
}

/* ── Phones ────────────────────────────────────────────────────────── */

function readPhones(adNode, comp) {
  // Standard types: emit single-value lowercase TYPE param matching TB's
  // native vCard 4 export shape (TEL;TYPE=home, TEL;TYPE=fax, …). Legacy
  // did the same; we deliberately omit the redundant `voice` qualifier
  // (vCard 4 default for TEL).
  addPhone(adNode, comp, "MobilePhoneNumber", "cell");
  addPhone(adNode, comp, "PagerNumber", "pager");
  addPhone(adNode, comp, "HomeFaxNumber", "fax");
  addPhone(adNode, comp, "HomePhoneNumber", "home");
  addPhone(adNode, comp, "Home2PhoneNumber", "home");
  addPhone(adNode, comp, "BusinessPhoneNumber", "work");
  addPhone(adNode, comp, "Business2PhoneNumber", "work");
  // Prefixed types: BusinessFaxNumber + Assistant/Car/Radio/CompanyMain.
  // Legacy emits TEL with a PascalCase TYPE and prefixes the value with
  // "<Type>: " so TB's UI (which doesn't render non-standard TYPE values)
  // shows the user what kind of number this is. Mirror that shape so the
  // round-trip is byte-identical to legacy.
  for (const [tag, { type, prefix }] of Object.entries(PREFIXED_PHONES)) {
    const v = readPathFrom(adNode, [tag]);
    if (!v) continue;
    const prop = new ICAL.Property("tel", comp);
    prop.setParameter("type", type);
    prop.setValue(prefix + v);
    comp.addProperty(prop);
  }
}

function addPhone(adNode, comp, tag, type) {
  const v = readPathFrom(adNode, [tag]);
  if (!v) return;
  const prop = new ICAL.Property("tel", comp);
  prop.setParameter("type", type);
  prop.setValue(v);
  comp.addProperty(prop);
}

/** Bucket TEL properties by their TYPE so each EAS field gets the right
 *  value back. Standard types use lowercase single-value TYPE; prefixed
 *  types use PascalCase TYPE with a "<Type>: " value-prefix that we
 *  strip before sending to the server (legacy parity). The buckets are
 *  consumed in two halves: standard + non-Company prefixed phones go
 *  to the Contacts page, CompanyMainPhone goes to Contacts2. */
function bucketPhones(comp) {
  const buckets = { cell: [], pager: [], fax: [], home: [], work: [] };
  const prefixed = {}; // EAS tag → value with prefix stripped

  for (const p of comp.getAllProperties("tel")) {
    const value = stringOf(p.getFirstValue());
    if (!value) continue;
    const types = paramTypes(p); // already lowercased

    // Prefixed types (BusinessFax, Assistant, Car, Radio, Company): match
    // first; their PascalCase TYPE value lowercases to a unique tag like
    // "workfax" / "assistant" / etc.
    let placed = false;
    for (const [tag, { type, prefix }] of Object.entries(PREFIXED_PHONES)) {
      if (types.includes(type.toLowerCase())) {
        prefixed[tag] = value.startsWith(prefix)
          ? value.slice(prefix.length)
          : value;
        placed = true;
        break;
      }
    }
    if (placed) continue;

    if (types.includes("cell")) {
      buckets.cell.push(value);
      continue;
    }
    if (types.includes("pager")) {
      buckets.pager.push(value);
      continue;
    }
    if (types.includes("fax")) {
      buckets.fax.push(value);
      continue;
    }
    if (types.includes("home")) {
      buckets.home.push(value);
      continue;
    }
    if (types.includes("work")) {
      buckets.work.push(value);
      continue;
    }
    // Untagged / unknown TYPE → home (legacy fallback).
    buckets.home.push(value);
  }
  return { buckets, prefixed };
}

/** Emit the seven Contacts-codepage standard phone fields in legacy's
 *  interleaved order: Mobile, Pager, HomeFax, HomePhone, Business,
 *  Home2, Business2 (legacy `Object.keys` order, contactsync.js:89-98). */
function writeStandardPhones(b, { buckets }) {
  emitIf(b, "MobilePhoneNumber", buckets.cell[0]);
  emitIf(b, "PagerNumber", buckets.pager[0]);
  emitIf(b, "HomeFaxNumber", buckets.fax[0]);
  emitIf(b, "HomePhoneNumber", buckets.home[0]);
  emitIf(b, "BusinessPhoneNumber", buckets.work[0]);
  emitIf(b, "Home2PhoneNumber", buckets.home[1]);
  emitIf(b, "Business2PhoneNumber", buckets.work[1]);
}

/** Emit the four Contacts-codepage prefixed phone fields in legacy's
 *  order (contactsync.js:124-127). CompanyMainPhone is excluded — it
 *  lives in the Contacts2 codepage and is emitted from there. */
function writePrefixedPhonesContacts(b, { prefixed }) {
  emitIf(b, "AssistantPhoneNumber", prefixed.AssistantPhoneNumber);
  emitIf(b, "CarPhoneNumber", prefixed.CarPhoneNumber);
  emitIf(b, "RadioPhoneNumber", prefixed.RadioPhoneNumber);
  emitIf(b, "BusinessFaxNumber", prefixed.BusinessFaxNumber);
}

/* ── Addresses (Home / Business / Other) ───────────────────────────── */

const ADDRESS_KINDS = [
  { wbxmlPrefix: "HomeAddress", types: ["home"] },
  { wbxmlPrefix: "BusinessAddress", types: ["work"] },
  { wbxmlPrefix: "OtherAddress", types: [] },
];

function readAddresses(adNode, comp, separator) {
  for (const kind of ADDRESS_KINDS) {
    const street = readPathFrom(adNode, [kind.wbxmlPrefix + "Street"]) ?? "";
    const city = readPathFrom(adNode, [kind.wbxmlPrefix + "City"]) ?? "";
    const state = readPathFrom(adNode, [kind.wbxmlPrefix + "State"]) ?? "";
    const zip = readPathFrom(adNode, [kind.wbxmlPrefix + "PostalCode"]) ?? "";
    const country = readPathFrom(adNode, [kind.wbxmlPrefix + "Country"]) ?? "";
    if (!street && !city && !state && !zip && !country) continue;
    const streetParts = splitStreet(street, separator);
    const prop = new ICAL.Property("adr", comp);
    if (kind.types.length)
      prop.setParameter(
        "type",
        kind.types.length === 1 ? kind.types[0] : kind.types,
      );
    // ADR shape: pobox, ext, street, locality, region, postal, country.
    prop.setValue([
      "",
      "",
      streetParts.length === 1 ? streetParts[0] : streetParts,
      city,
      state,
      zip,
      country,
    ]);
    comp.addProperty(prop);
  }
}

function writeAddresses(b, comp, separator) {
  const matched = new Set();
  for (const kind of ADDRESS_KINDS) {
    let chosen = null;
    for (const p of comp.getAllProperties("adr")) {
      if (matched.has(p)) continue;
      const types = paramTypes(p);
      if (kind.types.length === 0) {
        // Match an ADR with no relevant work/home TYPE.
        if (!types.includes("home") && !types.includes("work")) {
          chosen = p;
          break;
        }
      } else if (kind.types.every((t) => types.includes(t))) {
        chosen = p;
        break;
      }
    }
    if (!chosen) continue;
    matched.add(chosen);
    const v = chosen.getFirstValue();
    const arr = Array.isArray(v) ? v : [];
    const street = Array.isArray(arr[2])
      ? joinStreet(arr[2], separator)
      : (arr[2] ?? "");
    emitIf(b, kind.wbxmlPrefix + "Street", street);
    emitIf(b, kind.wbxmlPrefix + "City", arr[3]);
    emitIf(b, kind.wbxmlPrefix + "State", arr[4]);
    emitIf(b, kind.wbxmlPrefix + "PostalCode", arr[5]);
    emitIf(b, kind.wbxmlPrefix + "Country", arr[6]);
  }
}

function splitStreet(street, separator) {
  const ch = separatorChar(separator);
  if (!ch || !street.includes(ch)) return [street];
  return street.split(ch);
}

function joinStreet(streetArray, separator) {
  // `separator` is always a numeric char-code string by the time we get
  // here - the runner sets it via `String(account.custom?.seperator ??
  // "10")` (sync-runner.mjs:254). No defensive fallback needed.
  return streetArray.filter(Boolean).join(separatorChar(separator));
}

function separatorChar(separator) {
  if (separator == null || separator === "") return null;
  const code = Number(separator);
  if (!Number.isFinite(code)) return null;
  return String.fromCharCode(code);
}

/* ── Organization ──────────────────────────────────────────────────── */

function readOrganization(adNode, comp) {
  const company = readPathFrom(adNode, ["CompanyName"]) ?? "";
  const dept = readPathFrom(adNode, ["Department"]) ?? "";
  const title = readPathFrom(adNode, ["JobTitle"]) ?? "";
  if (company || dept) {
    const prop = new ICAL.Property("org", comp);
    prop.setValue(dept ? [company, dept] : company);
    comp.addProperty(prop);
  }
  if (title) comp.updatePropertyWithValue("title", title);
}

function writeOrganization(b, comp) {
  const org = comp.getFirstProperty("org");
  if (org) {
    const v = org.getFirstValue();
    const arr = Array.isArray(v) ? v : v ? [v] : [];
    emitIf(b, "CompanyName", arr[0]);
    emitIf(b, "Department", arr[1]);
  }
  emitIf(b, "JobTitle", stringOf(comp.getFirstPropertyValue("title")));
}

/* ── Notes / Body (codepage-aware) ─────────────────────────────────── */

function readNote(adNode, comp, asVersion) {
  if (useAirSyncBaseBody(asVersion)) {
    const bodyNode = childByTag(adNode, "Body");
    if (!bodyNode) return;
    const data = readPathFrom(bodyNode, ["Data"]);
    if (data) comp.updatePropertyWithValue("note", data);
  } else {
    const data = readPathFrom(adNode, ["Body"]);
    if (data) comp.updatePropertyWithValue("note", data);
  }
}

function writeNote(b, comp, asVersion) {
  const note = stringOf(comp.getFirstPropertyValue("note"));
  if (note) {
    if (useAirSyncBaseBody(asVersion)) {
      b.switchpage("AirSyncBase");
      b.otag("Body");
      b.atag("Type", "1");
      b.atag("EstimatedDataSize", String(note.length));
      b.atag("Data", note);
      b.ctag();
    } else {
      // AS 2.5 emits <Body> on the current Contacts page.
      b.atag("Body", note);
    }
  }
  // Always end on Contacts2 so the caller can emit Contacts2 fields
  // without an extra switchpage. switchpage is a single SWITCH_PAGE
  // token regardless of the current page, so this is a no-op cost when
  // we never moved off Contacts (empty note, AS 2.5).
  b.switchpage("Contacts2");
}

function useAirSyncBaseBody(asVersion) {
  return asVersion !== "2.5";
}

/* ── Picture ───────────────────────────────────────────────────────── */

function readPicture(adNode, comp) {
  const v = readPathFrom(adNode, ["Picture"]);
  if (!v) return;
  const prop = new ICAL.Property("photo", comp);
  prop.setParameter("value", "uri");
  prop.setValue(`data:image/jpeg;base64,${v}`);
  comp.addProperty(prop);
}

function writePicture(b, comp) {
  const photo = comp.getFirstProperty("photo");
  if (!photo) return;
  const value = stringOf(photo.getFirstValue());
  if (!value) return;
  const m = /^data:image\/[^;]+;base64,(.+)$/i.exec(value);
  if (m) b.atag("Picture", m[1]);
}

/* ── Categories ────────────────────────────────────────────────────── */

function readCategories(adNode, comp) {
  const cats = childByTag(adNode, "Categories");
  if (!cats) return;
  const list = readChildTexts(cats, "Category");
  if (list.length) {
    const prop = new ICAL.Property("categories", comp);
    prop.setValues(list);
    comp.addProperty(prop);
  }
}

function writeCategories(b, comp) {
  const prop = comp.getFirstProperty("categories");
  if (!prop) return;
  const values = prop.getValues().map(stringOf).filter(Boolean);
  if (!values.length) return;
  b.otag("Categories");
  for (const v of values) b.atag("Category", v);
  b.ctag();
}

/* ── NickName + IM (Contacts2 codepage) ────────────────────────────── */

function readNickName(adNode, comp) {
  const v = readPathFrom(adNode, ["NickName"]);
  if (v) comp.updatePropertyWithValue("nickname", v);
}

function readIMs(adNode, comp) {
  for (let i = 0; i < 3; i++) {
    const tag = i === 0 ? "IMAddress" : `IMAddress${i + 1}`;
    const v = readPathFrom(adNode, [tag]);
    if (v) comp.addPropertyWithValue("impp", v);
  }
}

function writeIMs(b, comp) {
  const ims = comp
    .getAllProperties("impp")
    .map((p) => stringOf(p.getFirstValue()))
    .filter(Boolean);
  for (let i = 0; i < Math.min(ims.length, 3); i++) {
    const tag = i === 0 ? "IMAddress" : `IMAddress${i + 1}`;
    b.atag(tag, ims[i]);
  }
}

/* ── Children (Contacts codepage container) ────────────────────────── */

function readChildren(adNode, comp) {
  const node = childByTag(adNode, "Children");
  if (!node) return;
  const names = readChildTexts(node, "Child");
  if (names.length)
    comp.addPropertyWithValue("x-eas-children", JSON.stringify(names));
}

function writeChildren(b, comp) {
  const raw = stringOf(comp.getFirstPropertyValue("x-eas-children"));
  if (!raw) return;
  let names;
  try {
    names = JSON.parse(raw);
  } catch {
    return;
  }
  if (!Array.isArray(names) || !names.length) return;
  b.otag("Children");
  for (const n of names) if (n) b.atag("Child", String(n));
  b.ctag();
}

/* ── Pass-through / opaque fields ──────────────────────────────────── */

function readPassThroughs(adNode, comp) {
  for (const tag of PASS_THROUGH_FIELDS) {
    const v = readPathFrom(adNode, [tag]);
    if (v) comp.addPropertyWithValue(passThroughVcardKey(tag), v);
  }
}

function writePassThroughs(b, comp) {
  for (const tag of PASS_THROUGH_FIELDS) {
    const v = stringOf(comp.getFirstPropertyValue(passThroughVcardKey(tag)));
    if (v) b.atag(tag, v);
  }
}

function passThroughVcardKey(tag) {
  return `x-eas-${tag.toLowerCase()}`;
}

/* ── Custom 1-4 (legacy `x-custom*` slots) ─────────────────────────── */

function readCustomFields(adNode, comp) {
  // Read side is namespace-agnostic - readPathFrom walks the parsed DOM
  // regardless of which WBXML codepage the tag came from.
  for (const [tag, slot] of Object.entries(CUSTOM_FIELD_MAP)) {
    const v = readPathFrom(adNode, [tag]);
    if (v) comp.updatePropertyWithValue(slot, v);
  }
}

// Write side is inlined in `appendApplicationDataFromVCard`: OfficeLocation
// emits on the Contacts page (between addresses and Picture) and the
// other three emit on the Contacts2 page (right after NickName, before
// IMs) — both positions match legacy's iteration order.

/* ── Helpers ───────────────────────────────────────────────────────── */

function newVCard() {
  const comp = new ICAL.Component(["vcard", [], []]);
  comp.updatePropertyWithValue("version", "4.0");
  return comp;
}

function parseVCard(vCard) {
  if (typeof vCard !== "string" || !vCard.trim()) return null;
  try {
    return new ICAL.Component(ICAL.parse(vCard));
  } catch {
    return null;
  }
}

function emitIf(b, tag, value) {
  const s = stringOf(value);
  if (s) b.atag(tag, s);
}

function paramTypes(prop) {
  const raw = prop.getParameter("type");
  if (!raw) return [];
  const arr = Array.isArray(raw) ? raw : [raw];
  return arr.map((t) => String(t).toLowerCase());
}

function stringOf(v) {
  if (v == null) return "";
  if (typeof v === "string") return v;
  if (Array.isArray(v)) return v.filter(Boolean).join(" ");
  return String(v);
}

function isoDateOnly(s) {
  // EAS sends dates as ISO 8601 with a trailing "T00:00:00.000Z"; vCard
  // BDAY / ANNIVERSARY are happiest with the YYYY-MM-DD prefix.
  const m = /^(\d{4}-\d{2}-\d{2})/.exec(String(s));
  return m ? m[1] : String(s);
}

function toIsoDateTime(s) {
  const m = /^(\d{4}-\d{2}-\d{2})/.exec(String(s));
  if (!m) return String(s);
  return `${m[1]}T00:00:00.000Z`;
}

function childByTag(node, tag) {
  if (!node?.children) return null;
  for (const c of node.children) if (c.tagName === tag) return c;
  return null;
}

