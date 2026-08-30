/**
 * Unit tests for what happens when a server will not answer OPTIONS.
 *
 * One hosted Exchange server answers OPTIONS with HTTP 500. That was
 * terminal and stayed terminal: the probe runs whenever an account has
 * no `asversion`, so the failure that stopped one being stored is the
 * same failure that runs again on every attempt, and picking a version
 * by hand could not break the loop because the pin lives in a different
 * field.
 *
 * What this pins is the shape of the failure, not a cure for it. A
 * refused credential and a redirect stay exactly what they are, so a
 * wrong password never turns into advice about protocol versions. Every
 * other failure becomes one the connect can act on, because a server that
 * will not answer this request may serve every other one perfectly well.
 *
 * Run with `npm run test:unit` (node --test).
 */

import { test, mock } from "node:test";
import assert from "node:assert/strict";
import { installWebextEnv } from "./support/webext-env.mjs";
installWebextEnv();

const NETWORK = new URL("../../src/modules/network.mjs", import.meta.url);

/** How many probes the negotiation made. */
let probes = 0;
/** What the next one answers. */
let answer = null;

const real = await import(NETWORK.href);
mock.module(NETWORK.href, {
  namedExports: {
    ...real,
    easOptions: async () => {
      probes += 1;
      if (answer instanceof Error) throw answer;
      return answer;
    },
  },
});
const { negotiateAsVersion, isNoOptionsAnswer } =
  await import("../../src/modules/eas/connect.mjs");

const httpError = (code, status) => {
  const e = new Error(`HTTP ${status}`);
  e.code = code;
  e.status = status;
  return e;
};

function expecting(next) {
  probes = 0;
  answer = next;
  return { accountId: "1", custom: { server: "https://x.invalid", user: "u" } };
}

test("a server that answers the probe is negotiated normally", async () => {
  const account = expecting({ versions: ["14.1", "16.1"], commands: ["Sync"] });
  const rv = await negotiateAsVersion({ account });
  assert.equal(rv.asVersion, "16.1");
  assert.deepEqual(rv.allowedAsVersions, ["14.1", "16.1"]);
  assert.equal(probes, 1);
});

test("a refused probe is one the connect can act on", async () => {
  const account = expecting(httpError("E:HTTP", 500));
  await assert.rejects(negotiateAsVersion({ account }), (err) => {
    assert.ok(isNoOptionsAnswer(err), "the caller has to be able to tell");
    return true;
  });
  assert.equal(probes, 1, "asked once, not retried");
});

test("a rejected credential stays a rejected credential", async () => {
  // Otherwise a wrong password ends up telling the user to choose a
  // protocol version, which is advice about the wrong problem.
  const account = expecting(httpError("E:AUTH", 401));
  await assert.rejects(negotiateAsVersion({ account }), (err) => {
    assert.equal(err.code, "E:AUTH");
    assert.ok(!isNoOptionsAnswer(err));
    return true;
  });
});

test("a redirect stays a redirect", async () => {
  const account = expecting(httpError("E:HOST_REDIRECT", 451));
  await assert.rejects(negotiateAsVersion({ account }), (err) => {
    assert.equal(err.code, "E:HOST_REDIRECT");
    assert.ok(!isNoOptionsAnswer(err));
    return true;
  });
});

test("a probe that answers with no versions is a probe failure", async () => {
  const account = expecting({ versions: [], commands: [] });
  await assert.rejects(negotiateAsVersion({ account }), (err) => {
    assert.ok(isNoOptionsAnswer(err));
    return true;
  });
});

test("a server advertising only versions we cannot speak is not one", async () => {
  // It answered. There is simply nothing in common, which is a different
  // problem with a different remedy, so it must not be mistaken for one.
  const account = expecting({ versions: ["12.0", "12.1"], commands: [] });
  await assert.rejects(negotiateAsVersion({ account }), (err) => {
    assert.ok(!isNoOptionsAnswer(err));
    assert.match(err.message, /12\.0/);
    return true;
  });
});

/* ── which version an account speaks ──────────────────────────────────── */

// Three values decide it, and only one of them is the user's: the config
// says a version or `auto`, the server advertises a list, and the overlap
// of that list with what we speak is the suggestion `auto` follows. The
// version the account is actually running on is settled once, when it
// connects, and nothing may move it afterwards - a protocol switch
// changes item-identity semantics between 14.x and 16.1.

const { suggestedAsVersion } = await import(
  new URL("../../src/modules/eas/connect.mjs", import.meta.url).href
);

/** What a connect decides, given a stored config and an advertised list.
 *  The provider spells this out in `#doConnectAndDiscover`; kept here as
 *  the rule rather than the code, so a test failure names the rule. An
 *  account that has never had the dropdown touched stores nothing, and
 *  nothing means the same as `auto`. */
const versionOnConnect = (config, advertised) =>
  (config || "auto") === "auto" ? suggestedAsVersion(advertised) : config;

test("the suggestion is the best overlap, in our preference order", () => {
  assert.equal(
    suggestedAsVersion(["2.5", "12.0", "14.0", "14.1", "16.0", "16.1"]),
    "16.1",
  );
  // 16.0 is not one we speak, so a server offering only it and older
  // gets the newest we both know.
  assert.equal(suggestedAsVersion(["2.5", "14.0", "14.1", "16.0"]), "14.1");
  assert.equal(suggestedAsVersion(["2.5"]), "2.5");
  // Nothing in common, and nothing said at all.
  assert.equal(suggestedAsVersion(["12.0", "12.1"]), null);
  assert.equal(suggestedAsVersion([]), null);
  assert.equal(suggestedAsVersion(undefined), null);
});

test("auto follows the server, a chosen version does not", () => {
  const advertised = ["2.5", "14.0", "14.1", "16.1"];
  assert.equal(versionOnConnect("auto", advertised), "16.1");
  // The case that started this: the server suggests 16.1 and the user
  // asked for 14.1. The user wins, or the setting means nothing.
  assert.equal(versionOnConnect("14.1", advertised), "14.1");
  assert.equal(versionOnConnect("2.5", advertised), "2.5");
});

test("an account with no choice stored follows the server", async () => {
  // Every account made before the dropdown worked has this unset, and a
  // new one is created holding "auto" outright.
  const advertised = ["2.5", "14.0", "14.1", "16.1"];
  assert.equal(versionOnConnect(undefined, advertised), "16.1");
  assert.equal(versionOnConnect("", advertised), "16.1");
});

test("a new account is created choosing auto", async () => {
  const { readFileSync } = await import("node:fs");
  const src = readFileSync(
    new URL("../../src/modules/eas-provider.mjs", import.meta.url),
    "utf8",
  );
  // One per flavour of account: OAuth, auto-discovered, and custom.
  assert.equal(
    src.match(/asversionselected: "auto",/g)?.length,
    3,
    "an account flavour is created without a stored version choice",
  );
  assert.match(
    src,
    /custom\?\.asversionselected \|\| "auto"/,
    "the connect no longer reads an unset choice as auto",
  );
});

test("a connect is the only thing that decides it", async () => {
  // Read from the provider itself rather than restated here: the rule is
  // that the version is written under `connecting` and nowhere else, so
  // what is asserted is that no other path writes it.
  const { readFileSync } = await import("node:fs");
  const src = readFileSync(
    new URL("../../src/modules/eas-provider.mjs", import.meta.url),
    "utf8",
  );
  const writes = [...src.matchAll(/asversion:\s*([^,\n}]+)/g)].map((m) =>
    m[1].trim(),
  );
  // The three empty ones are account-creation defaults; the live value is
  // written in exactly one place, and that place is guarded by
  // `connecting`.
  assert.deepEqual(
    writes.filter((w) => w !== '""'),
    ["asVersion"],
    "something other than the connect writes the account's EAS version",
  );
  assert.match(
    src,
    /if \(connecting && asVersion !== ctx\.account\.custom\?\.asversion\)/,
    "the one writer is no longer guarded by `connecting`",
  );
});
