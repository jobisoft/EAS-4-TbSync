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
import { readPathFrom } from "./wbxml-helpers.mjs";

/** vCard property name we use to track the EAS server-side serverId so
 *  pull-update / pull-delete can find the right local card without a
 *  separate id map. */
const X_EAS_SERVERID = "x-eas-serverid";

/** Bag of EAS fields with no clean vCard equivalent. We round-trip them
 *  through `X-EAS-<TAG>` properties so the server's value comes back
 *  unchanged on push. */
const PASS_THROUGH_FIELDS = [
  "Alias", "WeightedRank", "YomiCompanyName", "YomiFirstName", "YomiLastName",
  "CompressedRTF", "MMS", "ManagerName", "AssistantName", "Spouse",
  "OfficeLocation", "CustomerId", "GovernmentId", "AccountName",
];

/** EAS phone fields without a standard TYPE: kept as TEL with an
 *  `x-<key>` type so they round-trip cleanly. */
const X_PHONE_TYPES = {
  AssistantPhoneNumber: "x-assistant",
  CarPhoneNumber:       "x-car",
  RadioPhoneNumber:     "x-radio",
  CompanyMainPhone:     "x-company-main",
};

const PHONE_KEYS_BY_TYPE = invertMap({
  cell:     "MobilePhoneNumber",
  pager:    "PagerNumber",
});

/* ── Reader: ApplicationData → vCard ───────────────────────────────── */

/** Build a vCard 4.0 string from an EAS `<ApplicationData>` DOM node.
 *
 *    adNode      DOM node positioned at <ApplicationData>
 *    serverID    EAS server-side identifier; written as X-EAS-SERVERID
 *    asVersion   "2.5" | "14.0" | "14.1" | "16.1" - selects body codepage
 *    separator   ASCII char code (string) for multi-line address streets
 *    uid         optional vCard UID to embed (TB derives the contact id from it) */
export function applicationDataToVCard({ adNode, serverID, asVersion, separator, uid }) {
  const comp = newVCard();
  if (uid) comp.addPropertyWithValue("uid", uid);
  comp.addPropertyWithValue(X_EAS_SERVERID, serverID);

  readNames(adNode, comp);
  readFileAs(adNode, comp);
  readDates(adNode, comp);
  readEmails(adNode, comp);
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
  return comp.toString();
}

/* ── Writer: vCard → ApplicationData WBXML ─────────────────────────── */

/** Append Contacts / Contacts2 / AirSyncBase tags onto an open
 *  `<ApplicationData>` builder. Caller is responsible for switching the
 *  codepage back afterwards if more sibling commands follow. */
export function appendApplicationDataFromVCard({ builder, vCard, asVersion, separator }) {
  const comp = parseVCard(vCard);
  if (!comp) return;

  builder.switchpage("Contacts");
  writeNames(builder, comp);
  writeFileAs(builder, comp);
  writeDates(builder, comp);
  writeEmails(builder, comp);
  writeWeb(builder, comp);
  writePhones(builder, comp);
  writeAddresses(builder, comp, separator);
  writeOrganization(builder, comp);
  writePicture(builder, comp);
  writeCategories(builder, comp);
  writeChildren(builder, comp);
  writePassThroughs(builder, comp);

  builder.switchpage("Contacts2");
  writeNickName(builder, comp);
  writeIMs(builder, comp);

  writeNote(builder, comp, asVersion);
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
  const last   = readPathFrom(adNode, ["LastName"])   ?? "";
  const first  = readPathFrom(adNode, ["FirstName"])  ?? "";
  const middle = readPathFrom(adNode, ["MiddleName"]) ?? "";
  const title  = readPathFrom(adNode, ["Title"])      ?? "";
  const suffix = readPathFrom(adNode, ["Suffix"])     ?? "";
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
  const arr = Array.isArray(v) ? v : (v ? [v] : []);
  emitIf(b, "LastName",   arr[0]);
  emitIf(b, "FirstName",  arr[1]);
  emitIf(b, "MiddleName", arr[2]);
  emitIf(b, "Title",      arr[3]);
  emitIf(b, "Suffix",     arr[4]);
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

function readEmails(adNode, comp) {
  for (let i = 0; i < 3; i++) {
    const v = readPathFrom(adNode, [`Email${i + 1}Address`]);
    if (v) comp.addPropertyWithValue("email", v);
  }
}

function writeEmails(b, comp) {
  const emails = comp.getAllProperties("email").map(p => stringOf(p.getFirstValue())).filter(Boolean);
  for (let i = 0; i < Math.min(emails.length, 3); i++) {
    b.atag(`Email${i + 1}Address`, emails[i]);
  }
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
  // Single-occurrence + the second slots (Home2, Business2). Each EAS
  // tag becomes a separate TEL property with appropriate TYPE params.
  addPhone(adNode, comp, "MobilePhoneNumber",   ["cell"]);
  addPhone(adNode, comp, "PagerNumber",         ["pager"]);
  addPhone(adNode, comp, "HomeFaxNumber",       ["home", "fax"]);
  addPhone(adNode, comp, "HomePhoneNumber",     ["home", "voice"]);
  addPhone(adNode, comp, "Home2PhoneNumber",    ["home", "voice"]);
  addPhone(adNode, comp, "BusinessPhoneNumber", ["work", "voice"]);
  addPhone(adNode, comp, "Business2PhoneNumber",["work", "voice"]);
  addPhone(adNode, comp, "BusinessFaxNumber",   ["work", "fax"]);
  for (const [tag, xtype] of Object.entries(X_PHONE_TYPES)) {
    addPhone(adNode, comp, tag, [xtype]);
  }
}

function addPhone(adNode, comp, tag, types) {
  const v = readPathFrom(adNode, [tag]);
  if (!v) return;
  const prop = new ICAL.Property("tel", comp);
  prop.setParameter("type", types.length === 1 ? types[0] : types);
  prop.setValue(v);
  comp.addProperty(prop);
}

function writePhones(b, comp) {
  // Bucket TEL properties by their type set. We emit at most two for
  // home/work to cover Home2/Business2.
  const buckets = {
    cell: [], pager: [], homeFax: [], workFax: [],
    home: [], work: [],
  };
  const xPhones = {}; // x-assistant/x-car/etc. → first value wins
  for (const p of comp.getAllProperties("tel")) {
    const value = stringOf(p.getFirstValue());
    if (!value) continue;
    const types = paramTypes(p);
    const has = (t) => types.includes(t);
    let placed = false;
    for (const [xtype, easTag] of Object.entries(X_PHONE_TYPES).map(([t, x]) => [x, t])) {
      if (has(xtype)) { xPhones[easTag] = value; placed = true; break; }
    }
    if (placed) continue;
    if (has("cell"))  { buckets.cell.push(value);    continue; }
    if (has("pager")) { buckets.pager.push(value);   continue; }
    if (has("fax") && has("home"))   { buckets.homeFax.push(value); continue; }
    if (has("fax") && has("work"))   { buckets.workFax.push(value); continue; }
    if (has("home")) { buckets.home.push(value);    continue; }
    if (has("work")) { buckets.work.push(value);    continue; }
    // Untagged or unknown → drop into home as legacy did.
    buckets.home.push(value);
  }
  emitIf(b, "MobilePhoneNumber",    buckets.cell[0]);
  emitIf(b, "PagerNumber",          buckets.pager[0]);
  emitIf(b, "HomeFaxNumber",        buckets.homeFax[0]);
  emitIf(b, "BusinessFaxNumber",    buckets.workFax[0]);
  emitIf(b, "HomePhoneNumber",      buckets.home[0]);
  emitIf(b, "Home2PhoneNumber",     buckets.home[1]);
  emitIf(b, "BusinessPhoneNumber",  buckets.work[0]);
  emitIf(b, "Business2PhoneNumber", buckets.work[1]);
  for (const [tag, value] of Object.entries(xPhones)) emitIf(b, tag, value);
}

/* ── Addresses (Home / Business / Other) ───────────────────────────── */

const ADDRESS_KINDS = [
  { wbxmlPrefix: "HomeAddress",     types: ["home"] },
  { wbxmlPrefix: "BusinessAddress", types: ["work"] },
  { wbxmlPrefix: "OtherAddress",    types: [] },
];

function readAddresses(adNode, comp, separator) {
  for (const kind of ADDRESS_KINDS) {
    const street = readPathFrom(adNode, [kind.wbxmlPrefix + "Street"]) ?? "";
    const city   = readPathFrom(adNode, [kind.wbxmlPrefix + "City"])   ?? "";
    const state  = readPathFrom(adNode, [kind.wbxmlPrefix + "State"])  ?? "";
    const zip    = readPathFrom(adNode, [kind.wbxmlPrefix + "PostalCode"]) ?? "";
    const country = readPathFrom(adNode, [kind.wbxmlPrefix + "Country"]) ?? "";
    if (!street && !city && !state && !zip && !country) continue;
    const streetParts = splitStreet(street, separator);
    const prop = new ICAL.Property("adr", comp);
    if (kind.types.length) prop.setParameter("type", kind.types.length === 1 ? kind.types[0] : kind.types);
    // ADR shape: pobox, ext, street, locality, region, postal, country.
    prop.setValue(["", "", streetParts.length === 1 ? streetParts[0] : streetParts, city, state, zip, country]);
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
        if (!types.includes("home") && !types.includes("work")) { chosen = p; break; }
      } else if (kind.types.every(t => types.includes(t))) {
        chosen = p; break;
      }
    }
    if (!chosen) continue;
    matched.add(chosen);
    const v = chosen.getFirstValue();
    const arr = Array.isArray(v) ? v : [];
    const street = Array.isArray(arr[2]) ? joinStreet(arr[2], separator) : (arr[2] ?? "");
    emitIf(b, kind.wbxmlPrefix + "Street",     street);
    emitIf(b, kind.wbxmlPrefix + "City",       arr[3]);
    emitIf(b, kind.wbxmlPrefix + "State",      arr[4]);
    emitIf(b, kind.wbxmlPrefix + "PostalCode", arr[5]);
    emitIf(b, kind.wbxmlPrefix + "Country",    arr[6]);
  }
}

function splitStreet(street, separator) {
  const ch = separatorChar(separator);
  if (!ch || !street.includes(ch)) return [street];
  return street.split(ch);
}

function joinStreet(streetArray, separator) {
  return streetArray.filter(Boolean).join(separatorChar(separator) ?? "\n");
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
  const dept    = readPathFrom(adNode, ["Department"]) ?? "";
  const title   = readPathFrom(adNode, ["JobTitle"]) ?? "";
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
    const arr = Array.isArray(v) ? v : (v ? [v] : []);
    emitIf(b, "CompanyName", arr[0]);
    emitIf(b, "Department",  arr[1]);
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
  if (!note) return;
  if (useAirSyncBaseBody(asVersion)) {
    b.switchpage("AirSyncBase");
    b.otag("Body");
      b.atag("Type", "1");
      b.atag("EstimatedDataSize", String(note.length));
      b.atag("Data", note);
    b.ctag();
    b.switchpage("Contacts2");
  } else {
    b.atag("Body", note);
  }
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
  const list = [];
  for (const c of cats.children) {
    if (c.tagName === "Category") {
      const t = decodeIfNeeded(c.textContent);
      if (t) list.push(t);
    }
  }
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

function writeNickName(b, comp) {
  emitIf(b, "NickName", stringOf(comp.getFirstPropertyValue("nickname")));
}

function readIMs(adNode, comp) {
  for (let i = 0; i < 3; i++) {
    const tag = i === 0 ? "IMAddress" : `IMAddress${i + 1}`;
    const v = readPathFrom(adNode, [tag]);
    if (v) comp.addPropertyWithValue("impp", v);
  }
}

function writeIMs(b, comp) {
  const ims = comp.getAllProperties("impp").map(p => stringOf(p.getFirstValue())).filter(Boolean);
  for (let i = 0; i < Math.min(ims.length, 3); i++) {
    const tag = i === 0 ? "IMAddress" : `IMAddress${i + 1}`;
    b.atag(tag, ims[i]);
  }
}

/* ── Children (Contacts codepage container) ────────────────────────── */

function readChildren(adNode, comp) {
  const node = childByTag(adNode, "Children");
  if (!node) return;
  const names = [];
  for (const c of node.children) {
    if (c.tagName === "Child") {
      const t = decodeIfNeeded(c.textContent);
      if (t) names.push(t);
    }
  }
  if (names.length) comp.addPropertyWithValue("x-eas-children", JSON.stringify(names));
}

function writeChildren(b, comp) {
  const raw = stringOf(comp.getFirstPropertyValue("x-eas-children"));
  if (!raw) return;
  let names;
  try { names = JSON.parse(raw); } catch { return; }
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

/* ── Helpers ───────────────────────────────────────────────────────── */

function newVCard() {
  const comp = new ICAL.Component(["vcard", [], []]);
  comp.updatePropertyWithValue("version", "4.0");
  return comp;
}

function parseVCard(vCard) {
  if (typeof vCard !== "string" || !vCard.trim()) return null;
  try { return new ICAL.Component(ICAL.parse(vCard)); }
  catch { return null; }
}

function emitIf(b, tag, value) {
  const s = stringOf(value);
  if (s) b.atag(tag, s);
}

function paramTypes(prop) {
  const raw = prop.getParameter("type");
  if (!raw) return [];
  const arr = Array.isArray(raw) ? raw : [raw];
  return arr.map(t => String(t).toLowerCase());
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

function decodeIfNeeded(text) {
  if (text == null) return "";
  try { return decodeURIComponent(text); }
  catch { return text; }
}

function invertMap(obj) {
  const out = {};
  for (const [k, v] of Object.entries(obj)) out[v] = k;
  return out;
}
