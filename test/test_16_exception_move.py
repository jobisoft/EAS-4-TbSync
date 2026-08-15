"""16. Moving a recurrence exception - AS 16.x only.

Split out of section 3, which chained six tests through one fixture: when the
move failed the harness stopped the section, so the two tests after it never
ran at all. 3.5 had therefore never executed once, and 3.6 only twice. A test
that fails should cost its own coverage and nobody else's.

What is under test is the *second* edit to a series that already has
exceptions. Importing them works (section 3); moving one afterwards is where
an override and the master's EXDATE have been seen to vanish, on all three
16.x servers and intermittently - the same account has failed, passed 6/6 and
failed again inside twenty minutes with no code change. That is item 47, and
it is open: this section is where it will be caught.

Imports its own fixture rather than inheriting section 3's, because sections
no longer share state.
"""

import re
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import harness
import probes
from bridge import ok
from harness import test

NEEDS = ("events",)

# Matched on the SUMMARY, never the UID: the clean pull mints fresh UIDs, so
# a UID-keyed lookup reports an intact series as missing.
MARK = "SUMMARY:TZ6 weekly"
VERSIONS = ("16",)


# The occurrence the fixture moves, as the instant it denotes.
MOVED = datetime(2026, 9, 9, 13, 0, tzinfo=timezone.utc)


def _rid_instant(line):
    """The UTC instant a RECURRENCE-ID line denotes, or None.

    Matched by instant rather than by text on purpose. The same occurrence
    is written `TZID=America/New_York:20260909T090000` or
    `20260909T130000Z` depending on which side last rendered it, and this
    test used to search for the UTC spelling alone - so an override that was
    present and correct read as missing, which is how a test bug spent a day
    wearing item 47's clothes.
    """
    m = re.match(r"RECURRENCE-ID(?:;TZID=([^:;\r\n]+))?[^:]*:(\d{8})T?(\d{6})?Z?", line)
    if not m:
        return None
    tzid, day, hms = m.group(1), m.group(2), m.group(3) or "000000"
    naive = datetime.strptime(day + hms, "%Y%m%d%H%M%S")
    if line.rstrip().endswith("Z") or not tzid:
        return naive.replace(tzinfo=timezone.utc)
    try:
        return naive.replace(tzinfo=ZoneInfo(tzid)).astimezone(timezone.utc)
    except Exception:
        return None


def _override_block(body, instant=MOVED):
    """The override VEVENT for `instant`, in whatever zone it is written."""
    for block in re.findall(
        r"BEGIN:VEVENT(?:(?!BEGIN:VEVENT)[\s\S])*?END:VEVENT", body
    ):
        for line in block.splitlines():
            if not line.startswith("RECURRENCE-ID"):
                continue
            if _rid_instant(line) == instant:
                return block
    return None


def _series(s):
    for item in s.items("events", "event"):
        if MARK in (item.get("item") or ""):
            return item
    return None


@test("16.1", "import the series with its two exceptions", VERSIONS)
def t_16_1(s):
    # Section 3 asserts what this sends; here it is only the fixture the
    # move needs, so it checks just that the series arrived intact.
    s.mark()
    ok("items.create", type="event", ical=probes.fixture("tz-test-exdate.ics"))
    s.sync()
    harness.true(_series(s) is not None, "the series did not reach the calendar")
    harness.eq(s.changelog("events"), [], "changelog drained")


@test("16.2", "move the override - exactly one <Change>, no <Delete>", VERSIONS)
def t_16_2(s):
    item = _series(s)
    body = item["item"]
    # The override's DTSTART is representation-fragile: the fixture wrote
    # America/New_York 13:00, but as soon as any server echo rebuilds the
    # item (Exchange re-sends the master with its Exceptions shortly after
    # a push, and the post-push pull may pick that up), the codec renders
    # the same instant in the default timezone - Europe/Berlin 19:00. So
    # find the override COMPONENT by its RECURRENCE-ID and move whatever
    # DTSTART it carries one hour later; both forms land on the same UTC
    # instant, which is what 16.3 verifies after the clean pull.
    block = _override_block(body)
    harness.true(
        block is not None,
        "the 9 Sep override is not in the local item - it did not survive "
        "16.1's import. Item 47. Section 3 tests the import on its own: if "
        "that passes and this does not, the loss is in the echo after it, "
        "not in what was sent",
    )
    dt = re.search(r"DTSTART;TZID=([^:;\r\n]+):20260909T(\d{2})(\d{4})", block)
    harness.true(dt is not None, "the override carries no TZID DTSTART to move")
    hour = int(dt.group(2)) + 1
    moved_block = block.replace(
        dt.group(0), f"DTSTART;TZID={dt.group(1)}:20260909T{hour:02d}{dt.group(3)}"
    )
    # The end moves with the start. Moving only DTSTART walked the start
    # onto the untouched end and made the occurrence zero-length, which is
    # not what "move an occurrence" means anywhere - Thunderbird shifts
    # both - and which section 12's timing gate now (correctly) refuses to
    # push. The old fixture passed only because a degenerate event used to
    # go out unquestioned.
    de = re.search(r"DTEND;TZID=([^:;\r\n]+):20260909T(\d{2})(\d{4})", block)
    harness.true(de is not None, "the override carries no TZID DTEND to move")
    end_hour = int(de.group(2)) + 1
    moved_block = moved_block.replace(
        de.group(0), f"DTEND;TZID={de.group(1)}:20260909T{end_hour:02d}{de.group(3)}"
    )
    s.mark()
    ok("items.update", id=item["id"], ical=body.replace(block, moved_block))
    s.sync()
    cmds = s.instance_commands()
    harness.eq([c[0] for c in cmds], ["Change"], f"instance commands sent: {cmds}")


@test("16.3", "clean resync - one EXDATE, the move and the cancellation intact", VERSIONS)
def t_16_3(s):
    s.rebind("events")
    item = _series(s)
    harness.true(item is not None, "the series did not survive the clean pull")
    body = item["item"]

    # Count EXDATE *lines*. A substring count also matches the fixture's own
    # description - "16 Sep is cancelled by the EXDATE and must not appear at
    # all" - and reports one cancellation as two. That false positive
    # reproduced on two different 16.x servers before it was spotted, which
    # is exactly how a test bug earns a reputation as a product bug.
    exdates = [l for l in body.splitlines() if l.startswith("EXDATE")]
    harness.eq(len(exdates), 1, f"expected one EXDATE property, got {exdates}")
    harness.contains(exdates[0], "20260916", "the cancelled date")
    harness.contains(body, "RECURRENCE-ID", "the moved occurrence must come back")
    probes.reset(s)


