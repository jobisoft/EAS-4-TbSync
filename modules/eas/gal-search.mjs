/**
 * EAS Search command (codepage 15) for the Global Address List.
 *
 * Wire shape (after WBXML decode):
 *
 *   <Search>
 *     <Response>
 *       <Store>
 *         <Status>1</Status>
 *         <Result>
 *           <Properties>
 *             <DisplayName>…</DisplayName>
 *             <FirstName>…</FirstName>
 *             <LastName>…</LastName>
 *             <EmailAddress>…</EmailAddress>
 *             <MobilePhone>…</MobilePhone>
 *             <HomePhone>…</HomePhone>
 *             <Phone>…</Phone>
 *             <Title>…</Title>
 *             <Office>…</Office>
 *           </Properties>
 *         </Result>
 *         …
 *       </Store>
 *     </Response>
 *   </Search>
 *
 * Returned values match the `ContactProperties` shape that the
 * `addressBooks.provider.onSearchRequest` listener is expected to yield.
 */

import { createWBXML } from "../wbxml.mjs";
import { easRequest } from "../network.mjs";
import { readPathFrom } from "./wbxml-helpers.mjs";

const RANGE = "0-99";

function buildSearchBody(query) {
  const w = createWBXML();
  w.switchpage("Search");
  w.otag("Search");
  w.otag("Store");
  w.atag("Name", "GAL");
  w.atag("Query", query);
  w.otag("Options");
  // Range is required by Z-Push and harmless to Exchange.
  w.atag("Range", RANGE);
  w.ctag();
  w.ctag();
  w.ctag();
  return w.getBytes();
}

/** Run a GAL Search request and return mapped contact properties.
 *  Caller must have verified that the account's `allowedEasCommands`
 *  include `Search`. Returns an empty list on a non-success Status or
 *  when the response carries no Result nodes. */
export async function runGalSearch({ account, asVersion, query, companyName }) {
  const body = buildSearchBody(query);
  const { doc } = await easRequest({
    account,
    command: "Search",
    body,
    asVersion,
  });
  if (!doc) return [];

  const results = [];
  for (const result of doc.getElementsByTagName("Result")) {
    const props = readProperties(result, companyName);
    if (props) results.push(props);
  }
  return results;
}

function readProperties(resultNode, companyName) {
  // The Properties wrapper is mandatory for a real result row; entries
  // without it are filler (e.g. a Stores summary node) and are skipped.
  const r = (tag) => readPathFrom(resultNode, ["Properties", tag]);

  const out = {};
  const firstName = r("FirstName");
  const lastName = r("LastName");
  const displayName = r("DisplayName");
  const email = r("EmailAddress");
  const mobile = r("MobilePhone");
  const home = r("HomePhone");
  const work = r("Phone");
  // Properties.Title is the person's job title; Properties.Office is the
  // department / office location. (The legacy add-on shipped these
  // swapped against the EAS schema; the corrected mapping is here.)
  const jobTitle = r("Title");
  const department = r("Office");

  if (!firstName && !lastName && !displayName && !email) return null;

  if (firstName) out.FirstName = firstName;
  if (lastName) out.LastName = lastName;
  if (displayName) out.DisplayName = displayName;
  if (email) out.PrimaryEmail = email;
  if (mobile) out.CellularNumber = mobile;
  if (home) out.HomePhone = home;
  if (work) out.WorkPhone = work;
  if (jobTitle) out.JobTitle = jobTitle;
  if (department) out.Department = department;
  if (companyName) out.Company = companyName;
  return out;
}
