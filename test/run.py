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
failure, setup flow - are not covered here and are still done by hand.
"""

import json
import os
import sys
import time

_HERE = os.path.dirname(os.path.abspath(__file__))
# `test/vendor/` holds the two files TbSync owns - the bridge client and the
# run loop. Both directories go on the path so a test module's `import
# harness` reads the same regardless of which side of the boundary it is on.
sys.path.insert(0, os.path.join(_HERE, "vendor"))
sys.path.insert(0, _HERE)

import pathlib

import probes
import session as session_mod
from harness import REGISTRY, run, select

# One module per section, and every section stands alone: each clears and
# builds whatever it needs, so `npm test -- 5` is a complete run. Steps
# within a section chain on purpose - re-establishing the fixture per step
# would mean several more full syncs against a server that throttles.
# Run order, which is not section order. Section numbers are shared
# vocabulary - cited in code comments across the suite - so
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
    "test_15_task_identity",
    # Split out of section 3: the move is where item 47 lives, so it fails
    # without costing the import tests or the all-day binding their run.
    "test_16_exception_move",
    # Version-agnostic, and the only exception coverage a 14.x account gets.
    "test_17_allday_exceptions",
    # Also version-agnostic, and for the same reason: sections 3, 5 and 16
    # assert on a wire shape only 16.x produces, so the <=14.x path where
    # exceptions ride embedded in the master had no coverage at all.
    "test_18_exception_outcomes",
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
        s = session_mod.preflight(require=tuple(sorted(needed)) or ("events",))
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
    # Nothing is bound yet - each section binds what it needs and drops it
    # again - so this names what the run will touch, not what is selected
    # right now.
    print("  will use " + ", ".join(f"{k}={s.folders[k]}" for k in sorted(needed)))
    print(f"  pace     {s.sync_gap:.0f}s between syncs")
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
        # Disconnect everything, then bind back only what this section needs.
        # One action purges the queue, the identity map and the sync key,
        # because the provider clears them when a resource is disabled - so
        # the section starts from the server's version of the folder and
        # inherits nothing from the last one.
        session_mod.isolate(s, needs, indent="       ")
        probes.reset(s, needs)
        # `reset` deletes leftovers, which queues deletes of its own; it syncs
        # when it removed something, but a delete the server refuses stays
        # owed and nothing there looks again.
        session_mod.drain_queues(s, needs, indent="       ")

    def finish(section):
        """The other half of `prepare`: a section says what it leaves behind.

        The setup purges, so anything a section still owes at this point is
        about to be discarded unseen - including an edit that never reached
        the server, which is a real failure wearing the disguise of a clean
        start. Draining first distinguishes the two: work that can still be
        pushed is pushed, and only what refuses to leave fails the section.
        """
        needs = getattr(sys.modules[MODULE_BY_SECTION[section]], "NEEDS", ("events",))
        session_mod.drain_queues(s, needs, indent="       ")
        for kind in needs:
            try:
                owed = s.changelog(kind)
            except AssertionError:
                continue  # not bound, so nothing of ours can be owed
            if owed:
                statuses = sorted({e.get("status") for e in owed})
                raise AssertionError(
                    f"{len(owed)} {kind} edit(s) still owed after the section "
                    f"finished and the queue was drained ({', '.join(statuses)})"
                    f" - the section left work that will not push"
                )

    rc = run(tests, s, prepare=prepare, finish=finish)
    save_wire(s, selectors, rc)
    return rc


def save_wire(session, selectors, rc):
    """Keep the event log of every run, and say where it went.

    The wire is the only evidence there is for an intermittent failure, and
    preflight clears the buffer at the start of each run - so without this,
    every run destroyed the previous one's. Two days of chasing section 3
    were spent on failures whose wire had already been thrown away, and the
    one capture that survived long enough to be read is what finally showed
    the server sending an <Exceptions> block with an exception missing.

    Written on success as well as failure: the interesting comparison is a
    failing run against a passing one, and which is which is not known until
    afterwards.
    """
    from bridge import ok as bridge_ok

    try:
        entries = bridge_ok("getEventLog")["entries"]
    except Exception as e:  # noqa: BLE001 - never fail a run over its own record
        print(f"\n  (could not save the event log: {e})")
        return
    out = pathlib.Path(__file__).resolve().parent / "wire"
    out.mkdir(exist_ok=True)
    name = (
        f"{time.strftime('%Y%m%d-%H%M%S')}"
        f"-{session.account['accountName'].split('@')[0]}"
        f"-{'+'.join(selectors) or 'all'}"
        f"-{'pass' if rc == 0 else 'FAIL'}.json"
    )
    path = out / name
    path.write_text(json.dumps(entries, indent=1))
    print(f"  wire     {path} ({len(entries)} entries)")


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
