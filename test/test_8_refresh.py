"""8. The calendar's own refresh path.

Thunderbird arms a refresh timer for every calendar that says it can refresh,
and ours does. When it fires it calls `provider.onSync`, which we forward to
the host as a sync request - a sync nobody in this suite asked for, on a
schedule nobody here controls.

That path had never been asserted. It only ever showed up in a log as its own
failure, `[target] sync request for <uuid> failed: unknown folder`, which is
what a refresh looks like when it reaches a calendar the host no longer has a
folder for. The success case is what these tests pin down.

`calendars.synchronize` is the only way to reach `onSync` from a script: it
runs `calCachedCalendar.refresh()`, the same entry point as the Reload button
and the platform's timer.

Kept out of sections 6 and 7 deliberately - both are heavy enough to draw a
503 out of a throttled account, and this needs to be able to run on its own.
"""

import time

import harness
from bridge import ok, rpc
from harness import test

# Resources this section touches.
NEEDS = ("events",)

# How long to wait for a refresh to reach the wire. It is asynchronous: the
# verb returns as soon as Thunderbird has asked, not when the sync is over.
SETTLE_S = 15

_before = {}


def _net_lines(s):
    """Wire traffic for the account under test, since the mark.

    Filtered by account on purpose: another account syncing in the
    background would otherwise satisfy this, and the test would pass without
    the refresh having reached anything.
    """
    return [
        e["message"]
        for e in s.log()
        if (e.get("message") or "").startswith("[eas:net]")
        and e.get("accountId") == s.account_id
    ]


def _failed_requests(s):
    return [
        e["message"]
        for e in s.log()
        if "[target] sync request" in (e.get("message") or "")
    ]


@test("8.1", "calendars.synchronize - the refresh reaches the provider and syncs")
def t_8_1(s):
    _before["items"] = len(s.items("events"))
    s.mark()
    ok("calendars.synchronize")

    # Poll rather than sleep the maximum: a quiet account answers in about a
    # second, and the section should not cost fifteen for a passing case.
    waited = 0
    while waited < SETTLE_S and not _net_lines(s):
        time.sleep(1)
        waited += 1

    harness.true(
        _net_lines(s),
        f"nothing reached the server within {SETTLE_S}s of the refresh - the "
        f"calendar asked, and either onSync did not fire or the host refused "
        f"the request",
    )
    harness.eq(
        _failed_requests(s),
        [],
        "the calendar asked for a sync and the host turned it down - this is "
        "the 'unknown folder' case, a calendar still alive and still armed "
        "with a refresh timer after its binding was dropped",
    )


@test("8.2", "the refresh leaves the folder green and the store unchanged")
def t_8_2(s):
    # Settle first: the refresh is still running when 8.1 returns, and a
    # folder mid-sync is not a verdict on anything.
    s.sync()
    harness.eq(s.status("events"), "success", "folder status after a refresh")
    harness.eq(s.changelog("events"), [], "changelog drained")
    # A refresh is a pull, so on a quiet account it must not add or drop
    # anything. If this ever fails on its own, check whether the account
    # really was quiet before reading it as a bug.
    harness.eq(
        len(s.items("events")),
        _before["items"],
        "the refresh changed the item count",
    )
