#!/usr/bin/env python3
"""Bridge test suite for EAS-4-TbSync.

    npm test                 every test that applies to the granted account
    npm test -- 3            section 3
    npm test -- 3.4          one step
    npm test -- 2 5          several
    npm test -- --list       what would run, and what gates each test

These drive a live account through TbSync's bridge, so they need Thunderbird
running with the bridge switched on and pointed at an EAS account granting
contacts, events and tasks. Anything else stops at preflight with a message
saying what to change.

The steps that need a person instead - install lockstep, authentication
failure, setup flow - stayed behind in `docs/manual-test-plan.html`.
"""

import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
# `test/vendor/` holds the two files TbSync owns - the bridge client and the
# run loop. Both directories go on the path so a test module's `import
# harness` reads the same regardless of which side of the boundary it is on.
sys.path.insert(0, os.path.join(_HERE, "vendor"))
sys.path.insert(0, _HERE)

import probes
import session as session_mod
from harness import REGISTRY, run, select

# One module per section, and every section stands alone: each clears and
# builds whatever it needs, so `npm test -- 5` is a complete run. Steps
# within a section chain on purpose - re-establishing the fixture per step
# would mean several more full syncs against a server that throttles.
# Run order, which is not section order. Section numbers are shared
# vocabulary - cited in code comments and in docs/manual-test-plan.html - so
# a new section takes the next free number rather than renumbering the rest.
# Where it *runs* is a separate question: 8 is cheap and version-independent,
# and the heavy sections (5 onwards) have twice drawn a 503 out of a
# throttled account and stopped the run before reaching it.
MODULES = [
    "test_1_handshake",
    "test_2_round_trip",
    "test_8_refresh",
    "test_3_exceptions",
    "test_4_timezones",
    "test_5_digest",
    "test_6_provider_calendar",
    "test_7_task_recurrence",
    "test_10_contacts",
    "test_11_regressions",
    "test_12_capability",
    "test_14_body_format",
    # After everything that reads the folder, because it re-pulls the whole
    # thing: a section running behind it pays for a full download it did not
    # ask for.
    "test_13_resync",
    # Last, and not by number: disconnecting clears the account's folder
    # records, and a provider that mints fresh folder ids on reconnect (as
    # google does) leaves the bridge's folder-scoped grant pointing at ids
    # that no longer exist. Anything after this would run against a stale
    # grant.
    "test_9_disconnect",
]


MODULE_BY_SECTION = {m.split("_")[1]: m for m in MODULES}


def main(argv):
    selectors = [a for a in argv if not a.startswith("-")]
    listing = "--list" in argv

    for name in MODULES:
        __import__(name)

    # Bind only the resources the selected sections actually use.
    needed = set()
    for name in MODULES:
        mod = sys.modules[name]
        section = name.split("_")[1]
        if not selectors or any(
            sel == section or sel.startswith(section + ".") for sel in selectors
        ):
            needed.update(getattr(mod, "NEEDS", ("events",)))

    tests = select(selectors)
    if not tests:
        known = sorted({t["section"] for t in REGISTRY})
        print(f"nothing matches {selectors!r}. Sections: {', '.join(known)}")
        return 2

    if listing:
        for t in tests:
            gate = "/".join(t["versions"]) + ".x" if t["versions"] else "any"
            print(f"  {t['id']:<5} [{gate:>7}]  {t['description']}")
        print(f"\n  {len(tests)} test(s)")
        return 0

    try:
        first_needs = getattr(
            sys.modules[MODULE_BY_SECTION[tests[0]["section"]]], "NEEDS", ("events",)
        )
        s = session_mod.preflight(
            require=tuple(sorted(needed)) or ("events",), bind=first_needs
        )
        # Sections that test recurrence need the account to sync it. A
        # module says so with NEEDS_RECURRENCE; the check is here rather
        # than per-section so the run refuses before touching anything.
        recurring = sorted(
            {
                t["section"]
                for t in tests
                if getattr(
                    sys.modules[MODULE_BY_SECTION[t["section"]]],
                    "NEEDS_RECURRENCE",
                    False,
                )
            },
            key=int,
        )
        if recurring:
            session_mod.require_recurrence(s, recurring)
    except session_mod.PreflightError as e:
        print(f"\n  Cannot run: {e}\n")
        return 2

    print(f"\n  account  {s.account['accountName']}  (AS {s.version})")
    # What is actually selected, not what is merely granted - the two used
    # to be conflated here, which hid the fact that every sync was touching
    # all three resources.
    print("  syncing  " + ", ".join(f"{k}={s.folders[k]}" for k in s.active))
    print()

    def prepare(section):
        """Per-section preflight: the account is put into the state this
        section needs, rather than inheriting the last one's.

        Three things, all of which have bitten us. Selecting only what the
        section uses, because `syncAccount` syncs every selected folder;
        clearing leftover fixtures, because a crashed or throttled run leaves
        them behind and the next run then matches whichever copy comes first;
        and marking where the log stands, so the wire assertions read this
        section rather than the whole run. Errors need no marking - every
        bridge call checks for them.
        """
        needs = getattr(sys.modules[MODULE_BY_SECTION[section]], "NEEDS", ("events",))
        print(f"  -- section {section}")
        s.mark()
        session_mod.select_resources(s, needs, indent="       ")
        probes.reset(s, needs)

    return run(tests, s, prepare=prepare)


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
