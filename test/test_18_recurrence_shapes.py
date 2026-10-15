"""18. An import whose recurrence one item cannot carry.

Thunderbird stores shapes ActiveSync cannot: an .ics can hand it two
RRULEs, an RRULE beside RDATEs, or a bare list of dates. None survives a
push - a 14.1 server keeps only the last of two rules, a 16.1 server
refuses them, and a list of occurrences is stored without the dates.

So the provider holds the calendar to shapes that can be carried: one rule
per item, and a set of dates restated as a rule with the occurrences it
cannot place moved onto their dates by an override. What this section
walks is the path a user actually takes to reach that - an import - and
then whether the pieces survive a round trip.

One of the listed dates carries an override of its own, because a date is
allowed to say no more than that an occurrence exists: anything about its
content is an override sitting on it. Restating the dates as a rule has to
move that override with the occurrence it describes rather than mint a
second one beside it.

Version-agnostic on purpose. Everything that leaves here is an ordinary
series with modified occurrences, which both wire forms carry: <=14.x
embeds the exceptions in the master's payload, 16.x sends them as
per-instance commands, and the stored outcome must be the same either way.

A task is restated the same way and then cannot be sent at all, [MS-ASTASK]
declaring no exception element at any version, so 18.3 covers the refusal
rather than a round trip.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

NEEDS = ("events", "tasks")
NEEDS_RECURRENCE = True

MARK = "PROBE combined series"

# The subject the fixture puts on one listed occurrence and no other.
OWN_SUBJECT = "PROBE combined series own subject"

# What the fixture states: three from the rule, three listed.
WANTED = sorted(
    probes.line_instant(f"DTSTART:{d}T090000Z")
    for d in ("20261106", "20261107", "20261108", "20261119", "20261225", "20270103")
)


def _master(body):
    """The master VEVENT of a blob.

    Scoped deliberately: a zoned item carries a VTIMEZONE whose own
    definitions are full of RRULE and RDATE lines, so a whole-blob grep
    reads every zoned event as a ruled one.
    """
    blocks = re.findall(r"BEGIN:VEVENT(?:(?!BEGIN:VEVENT)[\s\S])*?END:VEVENT", body)
    for block in blocks:
        if not re.search(r"^RECURRENCE-ID", block, re.M):
            return block
    return blocks[0] if blocks else ""


def _overrides(items):
    """Every override across the pieces, as (DTSTART instant, block)."""
    out = []
    for item in items:
        body = item.get("item") or ""
        for block in re.findall(
            r"BEGIN:VEVENT(?:(?!BEGIN:VEVENT)[\s\S])*?END:VEVENT", body
        ):
            if not re.search(r"^RECURRENCE-ID", block, re.M):
                continue
            line = re.search(r"^DTSTART[^\r\n]*", block, re.M)
            out.append((probes.line_instant(line.group(0)) if line else None, block))
    return out


def _carried(items):
    """The override the fixture put on a listed date, wherever it landed.

    Found by when it happens, not by its RECURRENCE-ID: restating the
    dates as a rule is what moves that key onto an instant the rule
    produces, so the key is the thing under test.
    """
    at = probes.line_instant("DTSTART:20261225T090000Z")
    return [b for when, b in _overrides(items) if when == at]


def _ours(s):
    return [i for i in s.items("events", "event") if MARK in (i.get("item") or "")]


def _instants(s):
    """Every occurrence the calendar expands our pieces to, as instants."""
    got = ok(
        "items.query",
        rangeStart="20261001T000000Z",
        rangeEnd="20270301T000000Z",
        expand=True,
    )
    out = []
    for item in got:
        body = item.get("item") or ""
        if MARK not in body:
            continue
        line = re.search(r"^DTSTART[^\r\n]*", _master(body), re.M)
        if line:
            out.append(probes.line_instant(line.group(0)))
    return sorted(out)


@test("18.1", "an import one item cannot carry is split, and its dates restated")
def t_19_1(s):
    probes.reset(s, ("events",))
    ok("items.create", ical=probes.fixture("combined-series.ics"))

    items = _ours(s)
    harness.eq(len(items), 2, "the import did not become two items")

    ruled = [i for i in items if re.search(r"^RRULE:", _master(i["item"]), re.M)]
    harness.eq(len(ruled), 2, "a piece came out without a rule of its own")
    # No piece may still state occurrences as a list: nothing can carry it.
    for item in items:
        harness.true(
            not re.search(r"^RDATE", _master(item["item"]), re.M),
            "a piece still states its occurrences as a list of dates",
        )
    # The occurrences are the fixture's, whatever shape they now take.
    harness.eq(_instants(s), WANTED, "the split changed which occurrences exist")

    # Two listed dates need moving onto; the third is where the new rule
    # starts. A third override here means one was minted beside the one
    # the fixture already had, which is two overrides for one occurrence.
    harness.eq(len(_overrides(items)), 2, "an occurrence was overridden twice")
    carried = _carried(items)
    harness.eq(len(carried), 1, "the override on a listed date did not survive")
    harness.true(
        OWN_SUBJECT in carried[0],
        "the override was replaced by a minted one, losing what it said",
    )


@test("18.2", "both pieces survive a clean pull with the same occurrences")
def t_19_2(s):
    s.sync()
    harness.eq(s.status("events"), "success", "the folder did not accept the pieces")

    # The only reading that distinguishes what the server kept from what we
    # still hold: the local copies are deleted and pulled down again.
    s.rebind("events")

    items = _ours(s)
    harness.eq(len(items), 2, "a piece did not survive the round trip")
    harness.eq(_instants(s), WANTED, "the server kept different occurrences")
    carried = _carried(items)
    harness.eq(len(carried), 1, "the moved override did not survive the round trip")
    harness.true(
        OWN_SUBJECT in carried[0], "the server kept the occurrence without its subject"
    )


@test("18.3", "a task that moves an occurrence is refused, not sent short")
def t_19_3(s):
    probes.reset(s, ("tasks",))
    ok("items.create", ical=probes.fixture("moved-task.ics"), resource="tasks")
    s.mark()
    s.sync()

    # [MS-ASTASK] declares no exception element at any version, so sending
    # this would drop the override in silence and put the occurrence back
    # on the rule's own instant.
    harness.true(
        s.warnings("only as a rule"),
        "a task moving an occurrence was not refused",
    )
    probes.reset(s, ("events", "tasks"))
    s.sync()
