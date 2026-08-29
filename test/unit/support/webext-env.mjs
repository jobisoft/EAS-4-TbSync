/**
 * Stand-in for the `browser` / `messenger` globals - the HOST
 * ENVIRONMENT, never the code under test - just enough for the codec
 * modules to load and run under node:test.
 *
 * The honest boundary: `timezone-mapping.ensureLoaded()` needs the two
 * CSV files (real, read from src/) and Thunderbird's calendar timezone
 * service (VTIMEZONE definitions - which only exist inside a running
 * Thunderbird). We do not fake zone data; we offer none: `currentZone`
 * "UTC", no ids, no definitions. timezone-mapping pins UTC into its
 * tables unconditionally, so every UTC-shaped test resolves - and any
 * test that needs a real named zone belongs to the LIVE suite (section
 * 4 runs those against actual servers). If such a test lands here by
 * mistake it fails on resolution rather than passing against fabricated
 * zone data.
 *
 * Install explicitly from a test file:
 *
 *   import { installWebextEnv } from "./support/webext-env.mjs";
 *   installWebextEnv();
 *
 * Idempotent; node:test runs each file in its own process, so files
 * stay independent. ESM hoists static imports, so the call runs AFTER
 * the codec modules load - that is fine and relied upon: they touch
 * `messenger`/`browser`/`fetch` lazily, never at module top level. Idea from PR #345 (tomaskovacik); reimplemented.
 */

import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const SRC_ROOT = fileURLToPath(new URL("../../../src/", import.meta.url));
const FIXTURE_SCHEME = "webext-fixture:";

export function installWebextEnv() {
  if (globalThis.messenger?.__webextEnv) return;

  globalThis.messenger = {
    __webextEnv: true,
    calendar: {
      timezones: {
        currentZone: "UTC",
        timezoneIds: [],
        async getDefinition() {
          return null;
        },
      },
    },
    // contact-codec's readEmails() pulls a bare address out of a
    // "Display Name <addr@example.com>" mailbox string. Thunderbird's
    // real parser handles far more (quoting, lists, malformed input);
    // this covers exactly the plain fixtures the unit layer uses.
    messengerUtilities: {
      async parseMailboxString(raw) {
        const m = /<([^>]+)>/.exec(raw);
        return [{ email: m ? m[1] : String(raw).trim() }];
      },
      // Stand-in for TB's HTML→text: strips tags, turns block breaks into
      // newlines, decodes the handful of entities the fixtures use, trims.
      // The real converter is Thunderbird's; the codec only needs *a*
      // deterministic plaintext so the DESCRIPTION value assertion means
      // something. Named limits, not a general HTML parser.
      async convertToPlainText(body) {
        return String(body ?? "")
          .replace(/<\s*br\s*\/?>/gi, "\n")
          .replace(/<\/\s*(p|div|li|h[1-6])\s*>/gi, "\n")
          .replace(/<[^>]+>/g, "")
          .replace(/&nbsp;/g, " ")
          .replace(/&lt;/g, "<")
          .replace(/&gt;/g, ">")
          .replace(/&quot;/g, '"')
          .replace(/&#39;/g, "'")
          .replace(/&amp;/g, "&")
          .replace(/[ \t]+\n/g, "\n")
          .replace(/\n{3,}/g, "\n\n")
          .trim();
      },
    },
  };

  globalThis.browser = {
    runtime: {
      getURL: (relativePath) => FIXTURE_SCHEME + relativePath,
    },
    // Every key answers "", which is what an unknown one does in the real
    // extension - so code that falls back to a literal when a message is
    // missing takes that path here, and code that does not shows up as an
    // empty message rather than a TypeError from a missing API.
    i18n: {
      getMessage: () => "",
    },
    // storage is deliberately absent: no codec module may touch it, and
    // a test that trips over this is a layering bug worth hearing about.
  };

  const realFetch = globalThis.fetch?.bind(globalThis);
  globalThis.fetch = async (url, ...rest) => {
    const s = String(url);
    if (s.startsWith(FIXTURE_SCHEME)) {
      const text = readFileSync(
        SRC_ROOT + s.slice(FIXTURE_SCHEME.length),
        "utf8",
      );
      return { text: async () => text };
    }
    if (!realFetch) throw new Error(`webext-env: no fetch for ${s}`);
    return realFetch(url, ...rest);
  };
}
