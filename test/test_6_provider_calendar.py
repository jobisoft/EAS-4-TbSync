"""6. Provider-backed calendar.

The calendars are ours: our own type, our own item hooks, and the host does
not watch them. That buys a structural separation between a user edit and
our own sync write, and hands us a set of failures that are ours alone.
Each step guards one.

6.2: the platform announces a removal for every one of our calendars
whenever our type unregisters, which happens on each reload, update and
disable. Something has to tell that apart from the user deleting one, or a
reload silently deselects the folder.

6.4: a pre-tag exists only to stop an observer logging a write we are about
to make. Nothing observes a calendar we supply, so a tag left behind is
never consumed and sits in the queue forever, one per synced item.

6.8: the queue is ours and lives outside the folder row, so what ties it to
a folder is the session id naming the current binding. If that stops moving,
edits outlive the calendar they were made against.
"""

import harness
import probes
from bridge import ok, rpc
from harness import test

# Resources this section touches. Preflight binds only what the selected
# sections need - binding one is a full download, and the suite has no
# reason to pull an address book it never reads.
NEEDS = ("events",)
# Needs the account to sync recurrence - the probe is a weekly series.
NEEDS_RECURRENCE = True

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


@test("6.8", "a binding that ends takes its queued edits with it")
def t_6_8(s):
    """The queue for one of our calendars lives in the provider, and the only
    thing tying it to a folder is the session id the host mints for the
    current binding. Deselecting ends that binding - the calendar is deleted
    - so an edit still queued against it is an edit to something that no
    longer exists. Carrying it into the next binding would re-apply it to a
    freshly downloaded calendar the user never touched.

    This is also the check that the host is minting sessions at all. If it
    stopped, everything else here would still pass: the queue would simply
    persist, and only this test would notice.
    """
    slug = f"{SLUG}-session"
    summary = f"SUMMARY:{probes.MARKER} {slug}"
    ok(
        "items.create",
        type="event",
        ical=probes.event(slug, lines=["DTSTART:20261201T090000Z", "DTEND:20261201T100000Z"]),
    )
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained after the create")
    item = s.find("events", f"{probes.MARKER} {slug}", "event")
    harness.true(item is not None, "the probe was not created")
    before = s.folder("events")["sessionId"]
    harness.true(before, "the folder row carries no session id")

    # An edit the sync never gets to see. Only the SUMMARY line is touched -
    # rewriting the whole slug would rewrite the UID with it and re-key the
    # item, which is a different test and a confusing failure. Asserted
    # rather than assumed: a replace that matched nothing would write the
    # item back unchanged, and "no entry was queued" would look like the
    # bug this test is here to catch.
    harness.contains(item["item"], summary, "the summary line to edit")
    ok("items.update", id=item["id"], ical=item["item"].replace(summary, f"{summary} v2"))
    harness.true(s.changelog("events"), "the edit was not queued at all")

    s.rebind("events")

    after = s.folder("events")["sessionId"]
    harness.true(
        after and after != before,
        "the host must mint a new session for the new binding - without that "
        "the provider cannot tell the old queue apart from a live one",
    )
    harness.eq(
        s.changelog("events"),
        [],
        "the old binding's queue followed it into the new one",
    )

    # And the edit really never reached the server: what came back is the
    # version from before it.
    fresh = s.find("events", f"{probes.MARKER} {slug}", "event")
    harness.true(fresh is not None, "the item did not come back from the server")
    harness.true(
        f"{slug} v2" not in fresh["item"],
        "the queued edit was pushed after all - it belonged to a calendar "
        "that had already been deleted",
    )

    # The new binding queues and pushes like any other.
    harness.contains(fresh["item"], summary, "the summary line to edit")
    ok("items.update", id=fresh["id"], ical=fresh["item"].replace(summary, f"{summary} v3"))
    harness.true(s.changelog("events"), "an edit under the new session was not queued")
    s.sync()
    harness.eq(s.changelog("events"), [], "changelog drained")

    done = s.find("events", f"{probes.MARKER} {slug}", "event")
    ok("items.remove", id=done["id"])
    s.sync()
