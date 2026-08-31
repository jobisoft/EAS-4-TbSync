"""17. Exception outcomes, on every protocol version.

Sections 3, 5 and 16 assert on the wire - "exactly one <Change>, keyed by
this InstanceId" - which is inherently 16.x, so all three are gated to it.
That leaves the <=14.x path with no coverage at all, and it is not a variant
of the same code:

  - 16.0/16.1 MUST NOT send <Exceptions> to change an exception. Each one
    travels as its own request keyed by airsyncbase:InstanceId.
  - <=14.x embeds them in the master's payload, and `followUpPhase` re-sends
    the master as one full <Change> once the Add has been acked and there is
    a ServerId to hang them on.

`followUpPhase` fires only when an *Add* carrying overrides is acked, which
no other section does on 14.x - 17.1 creates a bare series and adds the
exceptions in a later edit, which is an ordinary modify. Measured: a full
72-test Z-Push run contains no <Add> carrying <Exceptions> anywhere in
its wire.

So this section asserts the *outcome* - what the calendar holds after a
clean pull - which is identical on both wire forms and is the thing a user
actually has. It runs everywhere, and a divergence between the two paths is
what fails it.

Everything is matched by instant, never by spelling: 16.x sends an
InstanceId in UTC and <=14.x embeds the exception in the zone it was
authored in, so the same occurrence comes back written two different ways.
"""

import harness
import probes
from bridge import ok
from harness import test

NEEDS = ("events",)

# Matched on the SUMMARY, never the UID: a clean pull mints fresh UIDs.
MARK = "SUMMARY:TZ6 weekly"

# What the fixture says, as instants rather than as text.
CANCELLED = probes.line_instant("EXDATE;TZID=America/New_York:20260916T090000")
OVERRIDE = probes.line_instant("RECURRENCE-ID;TZID=America/New_York:20260909T090000")
OVERRIDE_START = probes.line_instant("DTSTART;TZID=America/New_York:20260909T130000")


def _series(s):
    for item in s.items("events", "event"):
        if MARK in (item.get("item") or ""):
            return item
    return None


def _override_block(body):
    """The override VEVENT, found by the instant its RECURRENCE-ID denotes."""
    import re

    for block in re.findall(
        r"BEGIN:VEVENT(?:(?!BEGIN:VEVENT)[\s\S])*?END:VEVENT", body
    ):
        for line in block.splitlines():
            if line.startswith("RECURRENCE-ID"):
                if probes.line_instant(line) == OVERRIDE:
                    return block
    return None


def _override_start(body):
    block = _override_block(body)
    if block is None:
        return None
    for line in block.splitlines():
        if line.startswith("DTSTART"):
            return probes.line_instant(line)
    return None


@test("17.1", "import a series that already carries its exceptions")
def t_18_1(s):
    # One create, exceptions included - the shape that reaches
    # `followUpPhase` on <=14.x and the instance phase on 16.x. Every other
    # section builds the series first and adds the exceptions afterwards,
    # which is a different path on both.
    def attempt():
        for it in s.items("events", "event"):
            if MARK in (it.get("item") or ""):
                ok("items.remove", id=it["id"])
        s.mark()
        ok("items.create", type="event", ical=probes.fixture("tz-test-exdate.ics"))
        s.sync()

    s.conflict_retry(attempt)
    harness.true(_series(s) is not None, "the series did not reach the calendar")
    harness.eq(s.changelog("events"), [], "changelog drained")

    # Read back from the server, not from what we wrote: a local copy proves
    # only that Thunderbird accepted the item.
    s.rebind("events")
    item = _series(s)
    harness.true(item is not None, "the series did not survive the clean pull")
    body = item["item"]

    harness.contains(body, "RRULE", "the recurrence must survive the round trip")

    exdates = probes.vevent_lines(body, "EXDATE")
    harness.eq(len(exdates), 1, f"expected one EXDATE property, got {exdates}")
    harness.eq(
        probes.line_instant(exdates[0]),
        CANCELLED,
        f"the cancelled occurrence is not the one that was cancelled: {exdates}",
    )

    rids = probes.vevent_lines(body, "RECURRENCE-ID")
    harness.eq(len(rids), 1, f"expected one override, got {rids}")
    harness.eq(
        probes.line_instant(rids[0]),
        OVERRIDE,
        f"the override anchors to the wrong occurrence: {rids}",
    )
    harness.eq(
        _override_start(body),
        OVERRIDE_START,
        "the override came back at the wrong time - it was moved to 13:00 "
        "New York and must return as that instant, however it is spelled",
    )
    harness.contains(body, "TZ6 moved occurrence", "the override's own content")


@test("17.2", "move the override - the new time survives a clean pull")
def t_18_2(s):
    # The second edit to a series that already has exceptions, which is
    # where an override has been seen to vanish. Asserted on the stored
    # result so it reads the same on both wire forms.
    import re

    moved_to = None

    def shift(body):
        # Idempotent: the target time is computed from the fixture's own
        # value, not from whatever the body currently holds, so applying
        # this twice lands on the same instant.
        nonlocal moved_to
        block = _override_block(body)
        harness.true(block is not None, "the override is not in the item to move")
        dt = re.search(r"^DTSTART([^:\r\n]*):(\d{8}T\d{6})Z?", block, re.M)
        harness.true(dt is not None, "the override carries no DTSTART to move")
        de = re.search(r"^DTEND([^:\r\n]*):(\d{8}T\d{6})Z?", block, re.M)
        harness.true(de is not None, "the override carries no DTEND to move")

        def plus_one_hour(stamp):
            from datetime import datetime, timedelta

            return (
                datetime.strptime(stamp, "%Y%m%dT%H%M%S") + timedelta(hours=1)
            ).strftime("%Y%m%dT%H%M%S")

        # The end moves with the start. Moving only DTSTART walks the start
        # onto the untouched end and makes the occurrence zero-length, which
        # section 12's timing gate correctly refuses to push.
        new_block = block.replace(
            dt.group(0), f"DTSTART{dt.group(1)}:{plus_one_hour(dt.group(2))}"
        ).replace(de.group(0), f"DTEND{de.group(1)}:{plus_one_hour(de.group(2))}")
        moved_to = probes.line_instant(
            f"DTSTART{dt.group(1)}:{plus_one_hour(dt.group(2))}"
        )
        return body.replace(block, new_block)

    s.edit(lambda: _series(s), shift, missing="17.1 must have left the series")

    s.rebind("events")
    item = _series(s)
    harness.true(item is not None, "the series did not survive the clean pull")
    body = item["item"]

    # The anchor does not move when the occurrence does - RECURRENCE-ID names
    # which occurrence this is, DTSTART says when it now happens.
    rids = probes.vevent_lines(body, "RECURRENCE-ID")
    harness.eq(len(rids), 1, f"expected one override after the move, got {rids}")
    harness.eq(
        probes.line_instant(rids[0]),
        OVERRIDE,
        "the override stopped anchoring to the occurrence it overrides",
    )
    harness.eq(
        _override_start(body),
        moved_to,
        "the move did not survive - the override came back at its old time, "
        "so the change never reached the server or was overwritten by it",
    )

    exdates = probes.vevent_lines(body, "EXDATE")
    harness.eq(
        [probes.line_instant(e) for e in exdates],
        [CANCELLED],
        f"moving the override disturbed the cancellation: {exdates}",
    )

    probes.reset(s)
