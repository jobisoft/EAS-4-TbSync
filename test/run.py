#!/usr/bin/env python3
"""Bridge test suite for EAS-4-TbSync.

    npm test                 every test that applies to the granted account
    npm test -- 4            section 4
    npm test -- 4.3          one step
    npm test -- 2 5          several
    npm test -- --list       what would run, and what gates each test

A run is started in the background and writes its report to
`test/runs/<timestamp>-<what>.log` as it goes. Starting one prints that
path and the pid and returns at once - a run takes tens of minutes, so
watching it is `tail -f` on the file, and stopping it is `kill <pid>`,
which unwinds and puts the account's settings back. `--list` is instant
and stays in the foreground.

These drive a live account through TbSync's bridge, so they need Thunderbird
running with the bridge switched on and pointed at an EAS account granting
contacts, events and tasks. Anything missing there stops at preflight with a
message saying what to change. The handful of account settings a run needs
are not asked for - preflight sets them and puts them back afterwards.

The steps that need a person instead - install lockstep, authentication
failure, setup flow - are not covered here and are still done by hand.
"""

import json
import os
import re
import signal
import subprocess
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
    "test_3_refresh",
    "test_4_exceptions",
    "test_5_timezones",
    "test_6_digest",
    "test_7_provider_calendar",
    "test_8_task_recurrence",
    "test_9_contacts",
    # The contact photo, kept out of section 9 so a server that mishandles
    # one contact field cannot cost it its coverage.
    "test_10_contact_photo",
    "test_11_regressions",
    "test_12_capability",
    "test_13_body_format",
    "test_14_task_identity",
    # Split out of section 4: the move is where item 47 lives, so it fails
    # without costing the import tests or the all-day binding their run.
    "test_15_exception_move",
    # Version-agnostic, and the only exception coverage a 14.x account gets.
    "test_16_allday_exceptions",
    # Also version-agnostic, and for the same reason: sections 4, 6 and 15
    # assert on a wire shape only 16.x produces, so the <=14.x path where
    # exceptions ride embedded in the master had no coverage at all.
    "test_17_exception_outcomes",
    # The import path: a recurrence one item cannot carry, split and
    # restated. Version-agnostic - everything it sends is an ordinary
    # series - and it re-pulls the folder, so it sits with the sections
    # that do.
    "test_18_recurrence_shapes",
    # After everything that reads the folder, because it re-pulls the whole
    # thing: a section running behind it pays for a full download it did not
    # ask for.
    "test_19_resync",
    # Last, and not by number: disconnecting clears the account's folder
    # records, and a provider that mints fresh folder ids on reconnect (as
    # google does) leaves the bridge's folder-scoped grant pointing at ids
    # that no longer exist. Anything after this would run against a stale
    # grant.
    "test_20_disconnect",
]


MODULE_BY_SECTION = {m.split("_")[1]: m for m in MODULES}


def _start_in_background(argv, watch=True):
    """Re-exec this run detached, and hand the caller its report.

    A run takes tens of minutes and says nothing useful until it ends, so
    it has no business holding a terminal. It is started, it names where it
    is writing, and reading that file is the caller's job.

    Unbuffered on purpose: a buffered stdout redirected to a file stays
    empty until the process exits, which is exactly how an hour-long run
    once produced no visible progress at all.

    Everything goes to the file, preflight included. The run-level preflight
    can refuse - a server that will not connect, an account missing a
    resource - and that refusal is a result like any other, written where
    the rest of the run is written.
    """
    runs = os.path.join(_HERE, "runs")
    os.makedirs(runs, exist_ok=True)
    what = "+".join(a for a in argv if not a.startswith("-")) or "all"
    log = os.path.join(runs, f"{time.strftime('%Y%m%d-%H%M%S')}-{what}.log")
    fh = open(log, "w", encoding="utf-8")
    proc = subprocess.Popen(
        [sys.executable, "-u", os.path.abspath(__file__), "--worker", *argv],
        stdout=fh,
        stderr=subprocess.STDOUT,
        cwd=_HERE,
        # Its own session, so it outlives the shell that started it and can
        # be signalled as a group.
        start_new_session=True,
    )
    print(f"\n  started in the background")
    print(f"  log    {log}")
    print(f"  pid    {proc.pid}")
    print(f"  watch  tail -f {log}")
    print(f"  stop   kill {proc.pid}   (cleans up; kill -9 does not)\n")
    return _watch(log, proc.pid) if watch else 0


# Colours only when someone is looking: a redirected or piped watch stays
# plain, so grepping a captured watch is not an exercise in escape codes.
_PAINT = {"PASS": "\033[32m", "FAIL": "\033[31m", "ERROR": "\033[31m",
          "SKIP": "\033[33m"}


def _paint(line):
    colour = _PAINT.get(line.strip().split(" ", 1)[0]) if line.strip() else None
    if not colour or not sys.stdout.isatty():
        return line
    return f"{colour}{line.rstrip(chr(10))}\033[0m\n"


def _pid_of_run(path):
    """The pid the run recorded in its own first line, if it is there."""
    try:
        with open(path, "r", encoding="utf-8") as fh:
            m = re.match(r"\s*pid (\d+)\b", fh.readline())
    except OSError:
        return None
    return int(m.group(1)) if m else None


def _alive(pid):
    if pid is None:
        return False
    try:
        os.kill(pid, 0)
    except (ProcessLookupError, PermissionError):
        return False
    return True


def _watch(path, pid=None):
    """Follow a run's report until it ends, or until the reader gives up.

    Opens with everything already written, so attaching to a run an hour in
    reads the same as having watched it from the start.

    Interrupting this stops the watching and nothing else: the run is in
    its own session, so it neither sees the signal nor cares. That is the
    point of watching rather than waiting - a run can be looked in on and
    left alone again, and no Ctrl-C aimed at the watching can cost an hour
    of syncing.
    """
    if pid is None:
        pid = _pid_of_run(path)
    print(f"  watching {path}")
    print("  Ctrl-C stops watching - the run carries on\n")
    try:
        with open(path, "r", encoding="utf-8") as fh:
            for line in fh.readlines():
                sys.stdout.write(_paint(line))
            sys.stdout.flush()
            while True:
                line = fh.readline()
                if line:
                    sys.stdout.write(_paint(line))
                    sys.stdout.flush()
                    continue
                # Nothing more to read: done once the writer has gone, but
                # only after one last look, or its closing lines are lost
                # to the race between its exit and this check.
                #
                # Only when there is a writer to ask about. A log with no
                # pid line - one written before runs recorded it - is
                # followed until the reader stops, never abandoned on the
                # strength of a question that could not be asked.
                if pid is not None and not _alive(pid):
                    line = fh.readline()
                    if not line:
                        break
                    sys.stdout.write(_paint(line))
                    continue
                time.sleep(0.3)
    except KeyboardInterrupt:
        tip = f", stop it with: kill {pid}" if pid else ""
        print(f"\n  detached - the run is still going{tip}\n")
        return 0
    print()
    return 0


def _stop_cleanly(_signum, _frame):
    """SIGTERM, unwound rather than obeyed on the spot.

    Python's default terminates without unwinding, so the `finally` that
    puts the account's settings back would never run and a stopped run
    would leave the log level, sync-after-change and device state as it
    found them mid-flight. Raising instead lets that teardown happen, which
    is what makes `kill <pid>` a clean stop.
    """
    raise SystemExit("\n  stopped on SIGTERM\n")


def main(argv):
    # `--list` is instant and exists to be read, so it stays in the
    # foreground. Anything that actually syncs is backgrounded.
    if "--watch" in argv:
        rest = [a for a in argv if a != "--watch"]
        if not rest:
            print("\n  --watch needs the log file to follow\n")
            return 2
        return _watch(rest[0])

    if "--worker" in argv:
        argv = [a for a in argv if a != "--worker"]
        signal.signal(signal.SIGTERM, _stop_cleanly)
        # Written by the run itself, so a log found later says which
        # process wrote it - whether to check that it is still alive, or
        # to stop it without hunting for the id the launcher printed.
        print(
            f"  pid {os.getpid()}  started {time.strftime('%Y-%m-%d %H:%M:%S')}"
            f"  -  stop with: kill {os.getpid()}"
        )
    elif "--list" not in argv:
        return _start_in_background(
            [a for a in argv if a != "--no-watch"], watch="--no-watch" not in argv
        )

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

    # From here on the run owns settings it does not own outright: preflight
    # changes what it needs and registers each undo. The `finally` covers
    # every way out - a refusal in preflight, a failing run, Ctrl-C - because
    # any of them would otherwise hand back an account configured differently
    # from how it was found.
    try:
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
                session_mod.ensure_recurrence(s, recurring)
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
    finally:
        session_mod.restore_overrides()


def save_wire(session, selectors, rc):
    """Keep the event log of every run, and say where it went.

    The wire is the only evidence there is for an intermittent failure, and
    preflight clears the buffer at the start of each run - so without this,
    every run destroyed the previous one's. Two days of chasing section 4
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
