/**
 * The Windows FILETIME zero, read as "no value".
 *
 * FILETIME counts 100-nanosecond intervals from 1601-01-01T00:00:00Z, so a
 * field a server never set can serialise as that instant instead of being
 * omitted. Kerio Connect does exactly that: an Anniversary nobody set
 * arrives as 1601-01-01, and a Birthday the user has just cleared comes
 * back the same way - so the birthday could never be cleared, and every
 * contact gained an anniversary it never had.
 *
 * ActiveSync has no sentinel for absence - an unset element is omitted -
 * so a value that can only be a null is read as one.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();
import { isFiletimeZero } from "../../src/modules/eas/wbxml-helpers.mjs";

test("the epoch is recognised in both wire spellings", () => {
  assert.equal(isFiletimeZero("1601-01-01T00:00:00.000Z"), true);
  assert.equal(isFiletimeZero("16010101T000000Z"), true);
});

test("the epoch rendered in a server's own zone is the same non-value", () => {
  // A server east of UTC renders local midnight as the previous evening,
  // one west of it as the small hours. Both are the epoch, not a date.
  assert.equal(isFiletimeZero("1600-12-31T23:00:00.000Z"), true);
  assert.equal(isFiletimeZero("1601-01-01T05:00:00.000Z"), true);
});

test("a real date is never mistaken for it", () => {
  assert.equal(isFiletimeZero("1980-02-29T00:00:00.000Z"), false);
  assert.equal(isFiletimeZero("19800229T000000Z"), false);
  // Adjacent years are dates, not sentinels.
  assert.equal(isFiletimeZero("1601-01-02T00:00:00.000Z"), false);
  assert.equal(isFiletimeZero("1602-01-01T00:00:00.000Z"), false);
  assert.equal(isFiletimeZero("1600-12-30T00:00:00.000Z"), false);
});

test("nothing at all is not the epoch", () => {
  assert.equal(isFiletimeZero(""), false);
  assert.equal(isFiletimeZero(null), false);
  assert.equal(isFiletimeZero(undefined), false);
  assert.equal(isFiletimeZero("not a date"), false);
});

test("a contact's epoch birthday and anniversary are not stored", async () => {
  const { applicationDataToVCard } = await import(
    "../../src/modules/eas/contact-codec.mjs"
  );
  const { el } = await import("./support/ad-node.mjs");
  const ical = await applicationDataToVCard({
    adNode: el("ApplicationData", [
      el("FileAs", "Probe"),
      el("Birthday", "1601-01-01T00:00:00.000Z"),
      el("Anniversary", "1601-01-01T00:00:00.000Z"),
    ]),
    existingVCard: null,
    serverID: "srv-1",
    asVersion: "14.1",
  });
  assert.doesNotMatch(String(ical), /^BDAY/m, "an epoch birthday was stored");
  assert.doesNotMatch(
    String(ical),
    /^ANNIVERSARY/m,
    "an epoch anniversary was stored",
  );
});
