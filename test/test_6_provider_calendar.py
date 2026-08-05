"""6. Provider-backed calendar.

The calendars are ours: our own type, our own item hooks, and the host no
longer watches them. That buys a structural separation between a user edit
and our own sync write - but it also means several things that used to be
someone else's problem are now ours to get wrong, and each step here guards
one of them.

6.2 is the one with history. The platform announces a removal for every one
of our calendars whenever our type unregisters, which happens on each reload,
update or disable - so something has to tell that apart from the user
actually deleting one, or every reload silently deselects the folder.

6.4 guards the other side: a pre-tag exists only to stop our own observer
logging a write we are about to make. For a calendar we supply there is
nothing observing, so a tag left behind is never consumed and sits in the
changelog forever, one per synced item.
"""

import harness
import probes
from bridge import ok, rpc
from harness import test

# Resources this section touches. Preflight binds only what the selected
# sections need - binding one is a full download, and the suite has no
# reason to pull an address book it never reads.
NEEDS = ("events",)

PROVIDER_TYPE = "ext-eas4tbsync@jobisoft.de"
SLUG = "provider-cal"


@test("6.1", "deselect and reselect - a new calendar of our type, all items pulled")
def t_6_1(s):
    before = len(s.items("events", "event"))
    s.rebind("events")

    row = s.folder("events")
    harness.true(row["targetID"], "the folder rebound to a calendar")
    cal = ok("calendars.get")
    harness.eq(cal["type"], PROVIDER_TYPE, "calendar type")
    harness.eq(
        len(s.items("events", "event")), before, "every item came back after the re-pull"
    )


@test("6.2", "reloadProvider - folder stays selected; target and item count unchanged")
def t_6_2(s):
    import time

    row = s.folder("events")
    target, count = row["targetID"], len(s.items("events", "event"))
    synckey = (row.get("custom") or {}).get("synckey")

    ok("reloadProvider", accountId=s.account_id)
    # The provider needs a moment to come back, and the bridge with it.
    for _ in range(20):
        time.sleep(3)
        try:
            rpc("getState", timeout=10)
            break
        except Exception:
            continue
    time.sleep(3)

    row = s.folder("events")
    harness.true(
        row["selected"],
        "the folder was deselected by a reload - the platform's removal "
        "announcement for our own unregistering type was taken for a deletion",
    )
    harness.eq(row["targetID"], target, "targetID survived the reload")
    harness.eq((row.get("custom") or {}).get("synckey"), synckey, "sync key survived")
    harness.eq(len(s.items("events", "event")), count, "item count survived")


@test("6.3", "edit an item - one changelog entry, modified_by_user, with detail")
def t_6_3(s):
    ok(
        "items.create",
        type="event",
        ical=probes.event(
            SLUG,
            lines=[
                "DTSTART;TZID=Europe/Berlin:20261117T090000",
                "DTEND;TZID=Europe/Berlin:20261117T100000",
                "RRULE:FREQ=WEEKLY;COUNT=3",
            ],
            timezone=True,
        ),
    )
    s.sync()
    item = s.find("events", f"{probes.MARKER} {SLUG}", "event")
    harness.true(item is not None, "the probe series was not created")

    ok(
        "items.update",
        id=item["id"],
        ical=item["item"].replace(f"{probes.MARKER} {SLUG}", f"{probes.MARKER} {SLUG} v2"),
    )
    entries = s.changelog("events")
    mine = [e for e in entries if e["itemId"] == item["id"]]
    harness.eq(len(mine), 1, f"expected exactly one entry for the edit, got {entries}")
    harness.eq(mine[0]["status"], "modified_by_user", "entry status")
    # A recurring item carries the pre-edit exception fingerprint, which is
    # the only record of what the series looked like before the write.
    harness.true(
        "detail" in mine[0],
        "a recurring item's entry must carry detail.exceptions - nothing can "
        "reconstruct the previous shape once the new version is written",
    )
    s.settle("events")


@test("6.4", "full resync - no *_by_server entries left behind")
def t_6_4(s):
    s.rebind("events")
    left = [e for e in s.changelog("events") if str(e.get("status", "")).endswith("_by_server")]
    harness.eq(
        left,
        [],
        "pre-tags survived a full resync - for a calendar we supply nothing "
        "observes, so a tag is never consumed and accumulates one per item",
    )


@test("6.5", "calendars.rename - targetName follows the new name")
def t_6_5(s):
    was = s.folder("events").get("targetName")
    ok("calendars.rename", name="Renamed by test 4.5")
    import time

    time.sleep(2)
    harness.eq(s.folder("events").get("targetName"), "Renamed by test 4.5", "targetName")
    ok("calendars.rename", name=was or "Calendar")
    time.sleep(2)


@test("6.6", "calendars.remove - folder unselected, target cleared, sync state reset")
def t_6_6(s):
    import time

    ok("calendars.remove")
    time.sleep(3)
    s.sync()
    row = s.folder("events")
    harness.eq(row["targetID"], None, "targetID cleared after the calendar was deleted")
    custom = row.get("custom") or {}
    harness.eq(str(custom.get("synckey", "0")), "0", "sync key reset")
    harness.eq(custom.get("indexMap") or [], [], "index map emptied")


@test("6.7", "re-select and sync - all items return")
def t_6_7(s):
    import time

    ok(
        "setFolderSelected",
        accountId=s.account_id,
        folderId=s.folders["events"],
        selected=True,
    )
    time.sleep(2)
    s.sync()
    row = s.folder("events")
    harness.true(row["targetID"], "the folder rebound")
    # The refill only happens because 6.6 reset the sync key - with a stale
    # key the server answers "nothing has changed" and the calendar stays
    # empty.
    harness.true(
        len(s.items("events", "event")) > 0,
        "the calendar came back empty - the sync key was not reset, so the "
        "server reported no changes and there was nothing to refill from",
    )
    probes.reset(s, ("events",))
