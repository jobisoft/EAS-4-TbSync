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
import { EasHttpError, NET_ERR, easRequest } from "../network.mjs";
import { readPathFrom } from "./wbxml-helpers.mjs";

const RANGE = "0-99";

const PROVISION_REQUIRED_STATUSES = new Set(["141", "142", "143", "144"]);

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

/** Run a GAL Search request.
 *
 *  Returns `{ results, total }`: the mapped contact properties, and what
 *  the server says it found. [MS-ASCMD] has the server return as many
 *  entries as `<Range>` asks for - 100 by default, which is what we ask -
 *  and it "MUST also indicate the total number of entries that are found"
 *  in `<Total>`. So `total > results.length` means the answer is a
 *  truncated view of a larger match set, which is the only reliable way
 *  to know: counting rows cannot tell a GAL holding exactly 100 matches
 *  from one holding 400.
 *
 *  `total` is null when the server states none, which is what an empty
 *  result set looks like on some servers. The caller treats that as
 *  complete - there is nothing being hidden behind an absent count.
 *
 *  Caller must have verified that the account's `allowedEasCommands`
 *  include `Search`. Yields no results on a non-success Status or when
 *  the response carries no Result nodes. */
export async function runGalSearch({ account, asVersion, query, companyName }) {
  const body = buildSearchBody(query);
  const { doc } = await easRequest({
    account,
    command: "Search",
    body,
    asVersion,
  });
  if (!doc) return { results: [], total: null };

  const topStatus =
    doc.getElementsByTagName("Status")[0]?.textContent ?? null;
  if (topStatus && PROVISION_REQUIRED_STATUSES.has(topStatus)) {
    // Server demands re-Provision before honouring Search. Throw with
    // the shared transport-level shape so the caller's provision-
    // recovery wrapper can re-acquire the policy and retry.
    throw new EasHttpError(NET_ERR.PROVISION_REQUIRED, 0, {
      message: `Search rejected (Status=${topStatus}); server demands re-Provision`,
    });
  }

  const rows = doc.getElementsByTagName("Result");
  const results = [];
  for (const result of rows) {
    const props = readProperties(result, companyName);
    if (props) results.push(props);
  }
  // One <Total> per response, inside <Store> beside <Range>. Read flat
  // like the Result nodes above; a Search response carries no other.
  const totalText = doc.getElementsByTagName("Total")[0]?.textContent;
  const total = Number.isFinite(Number(totalText)) && totalText !== ""
    ? Number(totalText)
    : null;
  // `delivered` is what the server sent, `results` what we could map -
  // a row carrying neither a name nor an address is dropped by
  // `readProperties`. Completeness is a statement about the server's
  // match set, so it has to be judged against what it delivered;
  // comparing `total` with `results.length` would report a complete
  // answer as truncated whenever a single row was unusable.
  return { results, total, delivered: rows.length };
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
