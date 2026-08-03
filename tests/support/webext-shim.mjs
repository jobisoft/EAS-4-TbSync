/**
 * Minimal stand-in for the `browser`/`messenger` WebExtension globals,
 * just enough for `timezone-mapping.mjs`'s `ensureLoaded()` to complete
 * outside Thunderbird.
 *
 * `ensureLoaded()` needs two real files (WindowsTimezone.csv/Aliases.csv,
 * read via `browser.runtime.getURL` + `fetch`, both fully vendored in
 * this repo so no faking there) and a Thunderbird-calendar-provided list
 * of known IANA zones with full VTIMEZONE definitions
 * (`messenger.calendar.timezones.*`), which we do NOT have outside a
 * running Thunderbird. We don't fake that data - we just don't offer any:
 * `timezoneIds: []` and `currentZone: "UTC"`. Per timezone-mapping.mjs's
 * own `loadTzInfo`, "UTC" (and "floating") short-circuit before ever
 * calling `getDefinition`, and UTC is unconditionally pinned into the
 * resolver's tables regardless of what `timezoneIds` contains - so this
 * is enough for `ensureLoaded()` to succeed and for every test that only
 * needs UTC-resolution (the common case: EAS 16.1 sends UTC times with
 * no per-event TimeZone blob at all).
 *
 * Tests that need a *specific* non-UTC IANA zone resolved are out of
 * scope for this shim - see TEST-PLAN.md's "known gaps" section.
 */

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const SRC_ROOT = fileURLToPath(new URL("../../src/", import.meta.url));

globalThis.messenger = {
  calendar: {
    timezones: {
      currentZone: "UTC",
      timezoneIds: [],
      async getDefinition() {
        return null;
      },
    },
  },
  // contact-codec.mjs's readEmails() calls this to pull a bare address
  // out of a "Display Name <addr@example.com>" mailbox string. Real
  // Thunderbird's parser handles far more (quoted names, multiple
  // addresses, malformed input); this is only enough for the plain
  // "name <addr>" / bare-address fixtures the contact-codec tests use.
  messengerUtilities: {
    async parseMailboxString(raw) {
      const m = /<([^>]+)>/.exec(raw);
      return [{ email: m ? m[1] : raw }];
    },
  },
};

globalThis.browser = {
  runtime: {
    getURL(relativePath) {
      return `webext-fixture:${relativePath}`;
    },
  },
};

const realFetch = globalThis.fetch?.bind(globalThis);

globalThis.fetch = async (url) => {
  const s = String(url);
  if (s.startsWith("webext-fixture:")) {
    const relativePath = s.slice("webext-fixture:".length);
    const text = readFileSync(SRC_ROOT + relativePath, "utf8");
    return { text: async () => text };
  }
  if (!realFetch) {
    throw new Error(`webext-shim: no real fetch available for ${s}`);
  }
  return realFetch(url);
};
