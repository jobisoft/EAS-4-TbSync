"""3. Recurrence exceptions - AS 16.x only.

One series, one cancelled occurrence and one moved one, from
`fixtures/tz-test-exdate.ics`. On 16.x each exception travels as its own
top-level command carrying an `InstanceId`, where <=14.x embeds
`<Exceptions>` inside the master's payload - so this section is gated, and
the same behaviour on 14.1 is correct while looking completely different on
the wire.

3.3 is the regression test for the re-assertion bug: before the exception
fingerprint landed, touching the master re-sent every occurrence and Exchange
rejected each one it already had.

Self-contained - 3.1 clears and imports - so `npm test -- 3` is a complete
run. The steps within it deliberately chain, because re-importing the series
per step would mean four more full syncs against a server that throttles.
"""

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches. Preflight binds only what the selected
# sections need - binding one is a full download, and the suite has no
# reason to pull an address book it never reads.
NEEDS = ("events",)

# Matched on the SUMMARY, never the UID: 3.5's clean pull mints fresh UIDs,
# so a UID-keyed lookup reports an intact series as missing.
MARK = "SUMMARY:TZ6 weekly"
VERSIONS = ("16",)


def _series(s):
    for item in s.items("events", "event"):
        if MARK in (item.get("item") or ""):
            return item
    return None


@test("3.1", "import - Delete 20260916T130000Z and Change 20260909T130000Z", VERSIONS)
def t_3_1(s):
    s.mark()
    ok("items.create", type="event", ical=probes.fixture("tz-test-exdate.ics"))
    s.sync()

    cmds = s.instance_commands()
    harness.true(cmds, "no instance commands were sent for the exceptions")
    harness.contains(cmds, ("Delete", "20260916T130000Z"), "cancelled occurrence")
    harness.contains(cmds, ("Change", "20260909T130000Z"), "moved occurrence")
    harness.eq(s.changelog("events"), [], "changelog drained")


@test("3.2", "sync again with no edit - no instance commands", VERSIONS)
def t_3_2(s):
    s.mark()
    s.sync()
    harness.eq(s.instance_commands(), [], "an unchanged series must send nothing")


@test("3.3", "edit only the master's title - no instance commands", VERSIONS)
def t_3_3(s):
    item = _series(s)
    harness.true(item is not None, "3.1 must have left the series in place")
    s.mark()
    ok(
        "items.update",
        id=item["id"],
        ical=item["item"].replace("SUMMARY:TZ6 weekly", "SUMMARY:TZ6 weekly edited"),
    )
    s.sync()
    harness.eq(
        s.instance_commands(),
        [],
        "editing the master re-asserted its exceptions - each one the server "
        "already holds, and rejects",
    )


@test("3.4", "move the override - exactly one <Change>, no <Delete>", VERSIONS)
def t_3_4(s):
    item = _series(s)
    s.mark()
    ok(
        "items.update",
        id=item["id"],
        ical=item["item"].replace(
            "DTSTART;TZID=America/New_York:20260909T130000",
            "DTSTART;TZID=America/New_York:20260909T140000",
        ),
    )
    s.sync()
    cmds = s.instance_commands()
    harness.eq([c[0] for c in cmds], ["Change"], f"instance commands sent: {cmds}")


@test("3.5", "clean resync - one EXDATE, the move and the cancellation intact", VERSIONS)
def t_3_5(s):
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
