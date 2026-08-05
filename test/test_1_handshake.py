"""1. Handshake and baseline.

The cheapest possible statement that the account is usable: the provider is
answering, a sync completes, and what came down is not already corrupt. Every
later section assumes all three, so they run first and the numbering says so.
"""

import harness
import probes
from harness import test

# Resources this section touches. Preflight binds only what the selected
# sections need - binding one is a full download, and the suite has no
# reason to pull an address book it never reads.
NEEDS = ("events",)


@test("1.1", "getState - the provider reports active, and the account is clean")
def t_1_1(s):
    # Section 1 is where the account's state is established for everything
    # after it, so the sweep belongs here as much as the handshake does: a
    # run that begins on top of a previous run's litter is not a baseline.
    # Preflight already found the account this way, so what is left to check
    # is that the provider behind it is connected and not sitting on an
    # error - an account can be granted and still be broken.
    harness.eq(s.account["provider"], "eas", "provider")
    harness.true(s.account["enabled"], "account is enabled")
    harness.eq(s.account["error"], None, "account error")


@test("1.2", "syncAccount - folder status success")
def t_1_2(s):
    s.sync()
    # Only what this section actually selected. Iterating every *granted*
    # resource asserts against folders preflight has deliberately
    # deselected, which fails on a resource the section never wanted.
    for kind in s.active:
        harness.eq(s.status(kind), "success", f"{kind} folder status")
        harness.eq(s.folder(kind)["error"], None, f"{kind} folder error")


@test("1.3", "items.query - item count and distinct UID count agree, no duplicates")
def t_1_3(s):
    # A resync that re-adds instead of matching shows up here first: the
    # counts diverge before anything else looks wrong.
    for kind in s.active:
        type_ = "task" if kind == "tasks" else "event"
        if kind == "contacts":
            continue
        items = s.items(kind, type_)
        uids = set()
        for item in items:
            for line in (item.get("item") or "").splitlines():
                if line.startswith("UID:"):
                    uids.add(line[4:])
                    break
        harness.eq(
            len(uids),
            len(items),
            f"{kind}: {len(items)} item(s) but {len(uids)} distinct UID(s)",
        )
