"""13. Server-initiated resync - the identity map.

A folder's `synckey` is the server's record of what we have. When it decides
ours is stale - its own state rebuilt, a mailbox moved, a long absence - it
answers Sync with **Status 3**, and the client must start again from
`SyncKey=0`. The server then sends the whole folder as `<Add>` commands,
because from its side we hold nothing at all.

Every one of those Adds names an item we already have. Taking them at face
value would create a second copy of everything in the folder, which is the
failure this section exists to catch.

What prevents it is `findExistingByServerId`: the incoming ServerId is
looked up in `custom.indexMap`, and when it resolves to a local item the Add
becomes an update. Which is why a Status 3 resets the sync key **and
nothing else** - our record of which local item each server id names is not
what the server refused, and the entire folder is about to arrive needing
exactly that answer.

So this section is the proof that the map survives the reset and is what
does the matching. Nothing else covers it: section 1 checks for duplicates
but never provokes a resync, and the deselect/reselect other sections use
looks similar while testing the opposite thing - it deletes the local
calendar first, so every Add really is new and nothing has to be matched.

Events have a second net underneath, and 13.4 exists to stop it hiding a
broken map: an EAS calendar payload carries the item's `UID`, so `applyAdd`
would find the local twin and adopt it even if the map answered nothing.
Contacts have no UID on the wire and no such net, which is why "the map
survived" has to be asserted rather than inferred from the item count.

Kept in its own section because it is expensive - the server re-sends the
whole folder - and because it is the only one that writes to host storage
directly. It builds its own fixture and clears it again, so it touches no
other section's state.

If a run dies half way, the folder is left holding a key the server rejects.
That is self-healing: the next sync resyncs again and recovers, which is the
behaviour under test.
"""

import re

import harness
import probes
from bridge import ok
from harness import test

# Resources this section touches.
NEEDS = ("events",)

# The host key holding every account's folder rows.
FOLDERS_KEY = "tbsync.folders"

# A syntactically plausible key the server will not recognise. Deliberately
# not "0" - that is a bootstrap, a different path, and would prove nothing.
BOGUS_SYNCKEY = "999999999"

# Enough items that a folder-wide re-add is unmistakable, few enough that the
# re-pull costs one small round trip.
FIXTURE_ITEMS = 5

# What the folder looked like before the resync, filled in by 13.1.
_before = {}


def _uids(s):
    """UIDs of the events currently in the local calendar, in order.

    Same extraction as 1.3. Identity across a resync is the ServerId, never
    the UID - but locally a *matched* item is updated in place and keeps the
    UID it had, while a *re-created* one arrives with a fresh one. So the UID
    set is exactly the signal that separates the two outcomes.
    """
    out = []
    for item in s.items("events", "event"):
        for line in (item.get("item") or "").splitlines():
            if line.startswith("UID:"):
                out.append(line[4:])
                break
    return out


def _details(s):
    """Every log entry's `details` since the last mark, whitespace collapsed.

    The decoded WBXML of each request and response lands here, which is the
    only place a status code is visible - the messages themselves say only
    that a Sync was sent.
    """
    return [re.sub(r"\s+", " ", e.get("details") or "") for e in s.log()]


def _custom(s):
    return s.folder("events").get("custom") or {}


def _need_baseline():
    if not _before:
        raise harness.Skip("13.1 did not run, so there is no baseline")


@test("13.1", "a populated folder, with a sync key and an index to lose")
def t_13_1(s):
    # An empty folder would satisfy every assertion below without exercising
    # anything, and a test account's folder can hold anything from nothing to
    # hundreds - so the section supplies its own items rather than hoping.
    for i in range(FIXTURE_ITEMS):
        ok(
            "items.create",
            type="event",
            ical=probes.event(
                f"resync-{i}",
                lines=[
                    f"DTSTART:202609{10 + i:02d}T090000Z",
                    f"DTEND:202609{10 + i:02d}T100000Z",
                ],
            ),
        )
    # They have to reach the server to matter: an item the server never saw
    # is not one it can re-send as an <Add>.
    s.sync()

    uids = _uids(s)
    harness.true(
        len(uids) >= FIXTURE_ITEMS,
        f"created {FIXTURE_ITEMS} events but the folder holds {len(uids)}",
    )
    harness.eq(len(set(uids)), len(uids), "the folder already holds duplicates")

    key = str(_custom(s).get("synckey") or "")
    harness.true(
        key not in ("", "0"),
        "the folder has no sync key to invalidate - it has never synced",
    )
    index = _custom(s).get("indexMap") or []
    harness.true(
        len(index) >= len(uids),
        f"index holds {len(index)} entries for {len(uids)} items before the "
        f"resync - there is no map here to survive one",
    )
    _before["uids"] = set(uids)
    _before["count"] = len(uids)
    _before["index_uids"] = {e.get("uid") for e in index}


@test("13.2", "an unrecognised sync key: the server refuses it and we start over")
def t_13_2(s):
    _need_baseline()

    # Written straight into host storage. There is no verb for this and
    # there should not be - a sync key nobody issued is not a state the host
    # can be asked to enter. `storage.restore` merges what it is handed, so
    # only the folder rows move, and it refuses the bridge's own keys.
    snap = ok("storage.snapshot")
    folders = snap[FOLDERS_KEY]
    folders[s.account_id][s.folders["events"]]["custom"]["synckey"] = BOGUS_SYNCKEY
    ok("storage.restore", data={FOLDERS_KEY: folders})
    harness.eq(
        str(_custom(s).get("synckey")),
        BOGUS_SYNCKEY,
        "the corrupted key did not reach storage",
    )

    s.mark()
    s.sync()

    # Status 3 gets no log line of its own - the runner matches the constant
    # and resets - so the proof is the response body itself. Without this the
    # section could pass against a server that quietly accepted the key, and
    # every assertion after it would be about an ordinary sync.
    seen = [d for d in _details(s) if re.search(r"<Status>\s*3\s*</Status>", d)]
    harness.true(
        bool(seen),
        f"no <Status>3</Status> in the sync response - the server did not "
        f"refuse {BOGUS_SYNCKEY}, so nothing below is testing the resync "
        f"path. Needs the log level at 3",
    )

    # And the folder came back down as <Add>s. Without this the section
    # would also pass against a server that refused the key and then sent
    # nothing - no Adds means nothing to match, and 13.3 would be counting
    # items no re-import ever touched.
    adds = sum(line.count("Add") for line in s.wire() if line.startswith("RECV"))
    harness.true(
        adds >= FIXTURE_ITEMS,
        f"the resync received {adds} <Add>(s) for {FIXTURE_ITEMS} known "
        f"items - the server refused the key without re-sending the folder, "
        f"so there was nothing for the fallback to match",
    )

    # And it recovered: a fresh key means it went back to zero and was
    # issued a new one, rather than stopping at the refusal.
    after = str(_custom(s).get("synckey") or "")
    harness.true(
        after not in (BOGUS_SYNCKEY, "0", ""),
        f"the server refused the key but no new one was stored (got "
        f"{after!r}) - the resync stopped at the refusal",
    )


@test("13.3", "every item was re-matched, not re-created")
def t_13_3(s):
    _need_baseline()

    # The point of the section. The server re-sent the whole folder as Adds
    # and every one of them had to resolve to the item it already named.
    uids = _uids(s)
    harness.eq(
        len(uids),
        _before["count"],
        f"item count moved across the resync: {_before['count']} -> "
        f"{len(uids)}. Every server item arrived as an <Add>, and at least "
        f"one was created instead of matched",
    )
    harness.eq(
        len(set(uids)),
        len(uids),
        f"{len(uids) - len(set(uids))} duplicate UID(s) after the resync",
    )
    harness.eq(
        set(uids),
        _before["uids"],
        "the folder holds different items than before - same count, so one "
        "was re-created while another went missing",
    )


@test("13.4", "and it was the map that matched them, not the UID fallback")
def t_13_4(s):
    _need_baseline()

    index = _custom(s).get("indexMap") or []
    harness.true(
        len(index) >= _before["count"],
        f"index holds {len(index)} entries for {_before['count']} items - "
        f"the resync lost mappings instead of updating them",
    )
    harness.true(
        _before["index_uids"] <= {e.get("uid") for e in index},
        "the index no longer covers the items it did before the resync - "
        "entries were dropped and re-made rather than kept",
    )

    # 13.3 would pass on events even with the map gone entirely, because
    # `applyAdd` adopts a local twin by UID before creating anything. That
    # adopt logs, and here it must not have: a burst of it means the map
    # answered nothing and only the calendar-only net saved the folder.
    # Contacts have no such net.
    adopted = [
        e.get("message")
        for e in s.log()
        if "adopting the local copy" in (e.get("message") or "")
    ]
    harness.eq(
        adopted,
        [],
        f"{len(adopted)} item(s) were matched by their UID instead of by the "
        f"index - the map did not survive the resync, and an address book "
        f"would have been duplicated here",
    )

    harness.eq(s.changelog("events"), [], "the resync left entries queued")


@test("13.5", "and the sync after it is an ordinary incremental one")
def t_13_5(s):
    _need_baseline()

    s.mark()
    s.sync()
    harness.eq(s.status("events"), "success", "folder status after the resync")
    harness.eq(len(_uids(s)), _before["count"], "the follow-up sync changed the count")

    # Cleared here rather than left for the next section's reset: this is
    # the one section whose fixtures get re-downloaded in full, so leaving
    # them behind makes the *next* run's resync more expensive than it
    # needs to be.
    probes.reset(s, ("events",), report=False)
