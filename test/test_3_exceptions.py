"""3. Importing recurrence exceptions - AS 16.x only.

One series, one cancelled occurrence and one moved one, from
`fixtures/tz-test-exdate.ics`. On 16.x each exception travels as its own
top-level command carrying an `InstanceId`, where <=14.x embeds
`<Exceptions>` inside the master's payload - so this section is gated, and
the same behaviour on 14.1 is correct while looking completely different on
the wire.

Only the import and the two quiet syncs after it live here. Moving an
exception is section 16 and the all-day binding is section 17: both used to
chain off this fixture, so a failure in the move stopped the section and
cost the tests behind it their run.

3.3 is the regression test for the re-assertion bug: before the exception
fingerprint landed, touching the master re-sent every occurrence and Exchange
rejected each one it already had.

Self-contained - 3.1 imports - so `npm test -- 3` is a complete run. The
three steps chain, because re-importing the series per step would mean two
more full syncs against a server that throttles.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches. Preflight binds only what the selected
# sections need - binding one is a full download, and the suite has no
# reason to pull an address book it never reads.
NEEDS = ("events",)

# Matched on the SUMMARY, never the UID: a clean pull mints fresh UIDs, so a
# UID-keyed lookup reports an intact series as missing.
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

    probes.reset(s)
