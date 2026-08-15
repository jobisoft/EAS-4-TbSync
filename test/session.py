"""Preflight, and the helpers every test uses.

Preflight decides whether the suite may run at all. It refuses loudly rather
than testing the wrong thing: an EAS account with all three resources granted
is the contract, and anything else stops the run with a message saying what to
fix in the Bridge tab.

Nothing here hard-codes an account or folder id. Both change whenever an
account is reconfigured - during one afternoon of this suite's development
the Google account's folder id changed twice - so the granted target is
always discovered by probing scope.
"""

import os
import re
import time

import bridge
import harness
from bridge import rpc, ok

# The three resource kinds an EAS account must grant, as (label, targetType,
# the query used to probe the grant).
# Everything an EAS account can grant. Preflight requires only what the
# selected tests need, which is never all three: no section reads contacts.
KINDS = ("contacts", "events", "tasks")

# Smallest gap between two syncs of the same account, in seconds - the run's
# pace, set once in preflight and carried on the Session so a run can be
# slowed without editing anything.
#
# Pacing the syncs matters more than pacing the tests: one test can sync four
# or more times, a rebind is two, and `settle` retries up to three. And the
# throttle these servers apply is a rate over a short window rather than a
# budget that refills - measured 15 Aug 2026, when an account rested for two
# hours and one rested for 94 minutes each managed the same three or four
# minutes of work before their first 503. Resting does not buy sections;
# going slower is the only thing that can.
#   TBSYNC_TEST_SYNC_GAP=5 npm test      # the old pace, for a quick local run
DEFAULT_SYNC_GAP_S = float(os.environ.get("TBSYNC_TEST_SYNC_GAP", "20"))


class PreflightError(Exception):
    """The bridge is not pointed at something this suite can test."""


class Session:
    """One run against one account. Holds the discovered ids and the small
    vocabulary the tests share."""

    def __init__(self, account, folders, version):
        self.account = account
        self.account_id = account["accountId"]
        self.folders = folders  # kind -> folderId
        self.version = version
        self.family = version_family(version)
        self._last_sync = 0.0
        # The run's pace. Preflight sets it; every sync waits it out.
        self.sync_gap = DEFAULT_SYNC_GAP_S
        # Where the current section began in the event log. Errors are read
        # from here, not from the last sync, so one logged in a gap between
        # two syncs still fails the section it happened in.
        self.section_seq = 0
        # Resources actually selected for this run - set by preflight, and
        # what `syncAccount` will therefore touch. Not the same as `folders`,
        # which is everything the bridge grants.
        self.active = ()

    # ── the guard ───────────────────────────────────────────────────────

    def _active(self, kind):
        """Refuse to touch a resource this run has not selected.

        `syncAccount` only syncs selected folders, so an unselected one has
        no local target and every verb against it fails with the platform's
        "not bound to a calendar yet" - which reads like a product fault and
        is really a test asking for something its section declared it did not
        need.

        Making that impossible rather than merely unusual: three separate
        helpers reached past their section's NEEDS when selection narrowing
        first landed, and each was found by a confusing failure rather than
        by reading the code.
        """
        if kind in self.active:
            return kind
        raise AssertionError(
            f"this section did not select {kind!r} - it is running with "
            f"{list(self.active) or 'nothing'} selected. Either add {kind!r} "
            f"to the module's NEEDS, or do not touch it here."
        )

    # ── folders and status ──────────────────────────────────────────────

    def folder(self, kind):
        """The current folder row. Always re-read: status, changelog and
        targetID all move under us during a sync."""
        rows = ok("getFolders", accountId=self.account_id)["folders"]
        wanted = self.folders[kind]
        for row in rows:
            if row["folderId"] == wanted:
                return row
        raise AssertionError(f"folder {wanted} ({kind}) has vanished")

    def changelog(self, kind):
        """Pending entries for a folder, wherever they are kept.

        For contacts that is the host's folder row. For events and tasks it
        is the provider's own storage - it supplies those calendars, so it
        is handed the user's edits directly and queues them itself, and the
        folder row's changelog stays permanently empty. Reading the row for
        those would report "nothing pending" for a folder holding a dozen
        unpushed edits, and every assertion about a drained queue would pass
        without ever being tested.

        `getChangelog` asks the right side and raises when it cannot. It must
        never answer [] on failure: that is indistinguishable from an empty
        queue, and would turn the one breakage these assertions exist to
        catch into a pass.
        """
        self._active(kind)
        return ok(
            "getChangelog",
            accountId=self.account_id,
            folderId=self.folders[kind],
        )["entries"]

    def status(self, kind):
        self._active(kind)
        return self.folder(kind)["status"]

    # ── syncing ─────────────────────────────────────────────────────────

    def sync(self, allow_errors=False):
        """Sync the account, then fail if the sync itself reported an error.

        Checking is the point. A test asserts on specific values, so a sync
        that threw can still leave every assertion satisfied - the local
        store is unchanged, and "the item is still there" reads as success.
        Section 6 passed 7/7 while the log held a TypeError from the folder
        sync, and a 503 has twice produced failure text that looked like a
        product bug because the sync behind it never ran.

        So the transport and the provider get the same standing as an
        assertion: any error logged while this sync ran fails the test, with
        the server's own words.

        Every bridge call checks the log, so an error normally raises from
        `rpc` itself. This waits out the tail first: a sync resolves, and a
        moment later the last of what it logged arrives.

        Waits out the run's `sync_gap` first, so back-to-back syncs inside
        one test do not arrive as a burst.
        """
        waited = time.time() - self._last_sync
        if waited < self.sync_gap:
            time.sleep(self.sync_gap - waited)
        ok("syncAccount", accountId=self.account_id)
        self._last_sync = time.time()
        time.sleep(2)
        try:
            bridge.audit()
        except bridge.LoggedError:
            if not allow_errors:
                raise

    def settle(self, kind, tries=3):
        """Sync until the folder's changelog drains, or give up.

        Bounded on purpose: a queue that will not drain is a result, and a
        loop that waits forever turns it into a hang.
        """
        self._active(kind)
        for _ in range(tries):
            if not self.changelog(kind):
                return True
            self.sync()
        return not self.changelog(kind)

    def rebind(self, kind):
        """Deselect and reselect a resource: deletes the local target and
        pulls it down again from scratch.

        The only way to see what the server actually stored rather than what
        we still hold locally. Slow, so tests use it only where local state
        could otherwise pass for a successful push.
        """
        self._active(kind)
        fid = self.folders[kind]
        ok("setFolderSelected", accountId=self.account_id, folderId=fid, selected=False)
        time.sleep(2)
        self.sync()
        ok("setFolderSelected", accountId=self.account_id, folderId=fid, selected=True)
        time.sleep(2)
        self.sync()

    # ── items ───────────────────────────────────────────────────────────

    def items(self, kind, type_=None):
        self._active(kind)
        args = {"resource": "tasks"} if kind == "tasks" else {}
        if type_:
            args["type"] = type_
        return ok("items.query", **args)

    def cards(self):
        self._active("contacts")
        return ok("contacts.query")

    def find(self, kind, marker, type_=None):
        """First item whose body contains `marker`, else None.

        Anchor on something the server does not rebuild - a SUMMARY, or an
        email address. Never on FN or the UID: a clean pull rebuilds FN from
        the name components ("Testkarte Google" came back as "Dr. Testkarte Q
        Google") and mints fresh UIDs, so matching on either reports a
        perfectly healthy round trip as a missing item.
        """
        for item in self.items(kind, type_):
            if marker in (item.get("item") or ""):
                return item
        return None

    def find_card(self, marker):
        for card in self.cards():
            if marker in ((card.get("properties") or {}).get("vCard") or ""):
                return card
        return None

    # ── the wire ────────────────────────────────────────────────────────

    def mark(self):
        """Remember where the log stands, so `log` and `wire` below report on
        what happens next rather than on the whole run.

        A mark rather than a clear: the log is the record of what the add-on
        did and is worth keeping whole, since a section that fails is read
        afterwards.
        """
        self.section_seq = ok("getEventLog")["lastSeq"]

    def log(self):
        return ok("getEventLog", sinceSeq=self.section_seq)["entries"]

    def wire(self):
        """Sync commands actually sent and received, as ["SEND Add", ...].

        What distinguishes "the local store looks right" from "we told the
        server" - the difference the round-trip tests exist to catch.
        """
        out = []
        for entry in self.log():
            details = re.sub(r"\s+", " ", entry.get("details") or "")
            if "<Commands>" not in details:
                continue
            message = (entry.get("message") or "").lower()
            tag = "SEND" if "send" in message else "RECV" if "receive" in message else None
            if tag:
                verbs = re.findall(r"<(Add|Change|Delete)>", details)
                out.append(f"{tag} {','.join(verbs)}")
        return out

    def instance_commands(self):
        """(verb, InstanceId) for every command carrying one, in order sent.

        On 16.x an exception travels as its own top-level command identified
        by InstanceId, so this is how a test says *which* occurrence moved
        rather than merely that something was sent.
        """
        out = []
        for entry in self.log():
            details = re.sub(r"\s+", " ", entry.get("details") or "")
            if "send" not in (entry.get("message") or "").lower():
                continue
            for m in re.finditer(r"<(Change|Delete)>(.*?)</\1>", details):
                iid = re.search(r"<InstanceId[^>]*>(.*?)</InstanceId>", m.group(2))
                if iid:
                    out.append((m.group(1), iid.group(1)))
        return out

    def server_won(self):
        """True when the server overruled a pushed edit in this window.

        Every push declares `<Conflict>1</Conflict>`, so the server is
        entitled to answer Status 7: it refuses the edit and keeps its own
        copy, which the pull at the end of the same sync brings down over
        whatever the test wrote. Read from the wire rather than from a log
        line, because the status is the fact and the wording around it is
        not.

        Scoped to `<Responses>`, which is where a verdict on something *we*
        sent appears. A 7 elsewhere in the reply is not about our push.
        """
        for entry in self.log():
            if "receive" not in (entry.get("message") or "").lower():
                continue
            details = re.sub(r"\s+", " ", entry.get("details") or "")
            for block in re.findall(r"<Responses>(.*?)</Responses>", details):
                if re.search(r"<Status[^>]*>7</Status>", block):
                    return True
        return False

    def edit(self, find, mutate, resource=None, after_write=None,
             missing="the item to edit is not there"):
        """Change a synced item and push it, absorbing a server-wins answer.

        The guarded form of "read an item, change its body, sync" - which is
        most of what this suite does to an item the server already holds, and
        every one of those is a place the server may answer Status 7. The
        provider then does the right thing and drops the edit, which leaves
        the test without the state its later assertions need.

        So the whole cycle is the unit that repeats: `find()` is called again
        each time, because after a rejection the body it returned no longer
        exists, and `mutate(body)` is applied to whatever the server imposed.
        `mutate` MUST therefore be idempotent - written so that applying it
        to its own output changes nothing further.

        `after_write` runs between the write and the sync, for the few tests
        that assert on what the local write did before anything is pushed.

        `resource` is the one `items.update` takes - "tasks", or omitted for
        a calendar. Not the resource name `find` uses, which is the suite's
        own ("events"); passing that through is refused by the bridge.
        """
        def attempt():
            item = find()
            harness.true(item is not None, missing)
            args = {"resource": resource} if resource else {}
            self.mark()
            ok("items.update", id=item["id"], ical=mutate(item["item"]), **args)
            if after_write:
                after_write()
            self.sync()

        self.conflict_retry(attempt)

    def conflict_retry(self, attempt, tries=3):
        """Run `attempt` until the server stops overruling it.

        A Status 7 is not a failure - it is the declared policy working -
        but it leaves the change the test needs undone and the server's own
        copy in its place. So the test makes the change again, against the
        state the server has just imposed.

        `attempt` performs one whole try: mark the wire, read the item as it
        now stands, edit it, sync. It is re-run from the top rather than
        resumed, because the body it edited no longer exists. Marking inside
        it also means the assertions after this call see only the attempt
        that stuck, not the discarded ones.

        No sleep of its own: `sync()` already waits out the run's pace, so a
        retry arrives no sooner than any other sync would.
        """
        for _ in range(tries):
            attempt()
            if not self.server_won():
                return
        raise AssertionError(
            f"the server overruled this edit {tries} times running "
            f"(Status 7, server-wins); it never took the change under test"
        )

    def warnings(self, needle=""):
        return [
            e["message"]
            for e in self.log()
            if e.get("level") == "warning" and needle in (e.get("message") or "")
        ]


def version_family(version):
    """Normalise an AS version to the family used for gating: 2.5, 14, 16."""
    v = str(version or "")
    if v.startswith("2."):
        return "2.5"
    if v.startswith("14"):
        return "14"
    if v.startswith("16"):
        return "16"
    return v or "unknown"


def preflight(provider="eas", require=KINDS, sync_gap=None):
    """Find the granted account, bind what the run starts with, return a
    Session.

    `require` is what the run needs to be *granted*, checked by probing
    without changing anything. `bind` is what to actually select now -
    normally just the first section's needs, since each section selects its
    own. Binding the union instead selected every resource the run would
    ever touch and synced them all before the first test, which is the cost
    per-section selection exists to avoid.

    Raises PreflightError with something actionable - every failure here is
    something the user fixes in the Bridge tab, not a bug in the add-on.
    """
    if not bridge.is_up():
        raise PreflightError(
            f"the bridge is not answering on 127.0.0.1:{bridge.PORT}.\n"
            f"  Start Thunderbird and switch the bridge on in TbSync's "
            f"Bridge tab."
        )

    # From here on every bridge call checks the event log; anything logged
    # before this point belongs to an earlier run.
    bridge.arm()

    accounts = ok("getState")["accounts"]
    granted = _granted_account(accounts)
    if not granted:
        raise PreflightError(
            "no account is granted to the bridge.\n"
            "  Pick one in TbSync's Bridge tab."
        )
    if granted["provider"] != provider:
        raise PreflightError(
            f"the bridge is pointed at a {granted['provider']!r} account "
            f"({granted['accountName']}), but this suite tests "
            f"{provider!r}.\n  Re-point it in the Bridge tab."
        )

    folders = _granted_folders(granted["accountId"])
    missing = [k for k in require if k not in folders]
    if missing:
        raise PreflightError(
            f"{granted['accountName']} does not grant: {', '.join(missing)}.\n"
            f"  This suite needs {', '.join(require)} - grant all of them in "
            f"the Bridge tab."
        )

    _require_debug_logging()
    _require_no_sync_after_change(granted)
    _require_server_wins(granted)
    _widen_and_clear_event_log()

    # Nothing is bound here. Every section now disconnects and re-binds what
    # it needs, so binding a guess at start-up would only be torn down again
    # - and it made a run of one section pay a full pull it never used.
    #
    # The cost is that a server this account cannot reach is no longer
    # reported before the first test; it surfaces as that section's own
    # setup failing, which names the same problem one line later.
    session = Session(granted, folders, (granted.get("custom") or {}).get("asversion"))
    session.sync_gap = sync_gap if sync_gap is not None else DEFAULT_SYNC_GAP_S
    return session


EVENT_LOG_MAX_WANTED = 5000


def _widen_and_clear_event_log():
    """Give the run a buffer big enough to hold a whole section, and start it
    empty.

    The log is a ring buffer, 500 entries by default. Every wire assertion
    reads it - `wire()` reconstructs what was sent from the logged requests -
    so once it rolls mid-section, a command that *was* sent is simply no
    longer there, and the assertion reports "the edit never reached the
    server". That is indistinguishable from the defect it exists to catch.

    Set, not checked, unlike the log level and sync-after-change: those are
    the account owner's choices about their own data, while this is a debug
    buffer size with no effect on what syncs. It is restored to the stored
    default by the host on the next Thunderbird restart, since the buffer
    itself is session-scoped.

    Cleared as well, so a section's reads cannot reach entries from a
    previous run - and so the whole buffer belongs to this one. The size is
    chosen for headroom rather than measured need: the cheapest section uses
    15 entries, the heaviest run for minutes, and the cost of guessing low is
    an assertion that reports a command as never sent when it was only no
    longer recorded.
    """
    snap = ok("storage.snapshot")
    settings = dict(snap.get("tbsync.settings") or {})
    settings["eventLogMax"] = EVENT_LOG_MAX_WANTED
    ok("storage.restore", data={"tbsync.settings": settings})
    ok("clearEventLog")


def require_recurrence(session, sections):
    """Refuse to run sections whose subject is recurrence when the account
    has recurrence sync switched off.

    With `syncrecurrence` false the codec emits no recurrence at all -
    deliberately, and the client-side rejection for sub-daily rules is
    gated on the same flag, because nothing can be misrepresented if
    nothing is sent. Every assertion about a series then fails, or worse
    passes vacuously: section 12's "the hourly event must not be pushed"
    reads as a code regression when it is an account setting.

    Checked, not set - like the log level. The flag is the user's
    choice, and an account created by an older release carries it off.
    """
    if (session.account.get("custom") or {}).get("syncrecurrence"):
        return
    raise PreflightError(
        f"{session.account['accountName']} has recurrence sync switched "
        f"off, and section(s) {', '.join(sections)} test recurrence.\n"
        f"  Switch 'Synchronize recurring events' on for this account - "
        f"without it no series reaches the wire, so these tests would "
        f"report a code fault that is not there."
    )


def _require_server_wins(account):
    """Refuse to run unless the account declares the server-wins policy.

    Every push states `<Conflict>` from this setting. On server-wins a
    rejected push is answered with Status 7, the server's own copy is taken
    and the edit is dropped - which `conflict_retry` is written to absorb.
    On device-wins the server accepts the push instead, so that answer never
    arrives and a test written around it would pass without ever exercising
    what it names.

    Checked, not set, like the two beside it: the value is the account
    owner's choice. Absent counts as server-wins, because that is what the
    provider does with it.
    """
    value = (account.get("custom") or {}).get("conflict")
    if value != "0":
        return
    raise PreflightError(
        f"{account['accountName']} is set to device-wins, and these tests "
        f"expect server-wins.\n"
        f"  Set 'Conflict resolution' to the server's copy winning, in this "
        f"account's TbSync configuration - on device-wins the server accepts "
        f"every push, so the conflict handling these tests cover is never "
        f"reached."
    )


def _require_no_sync_after_change(account):
    """Refuse to run while the account syncs itself after every change.

    Unlike the two checks around it, this one is not about assertions going
    vacuous - it is about the suite no longer being the only thing driving
    the account. With the option on, every item this suite writes arms a
    timer in the provider, and seconds later a sync it never asked for
    starts: mid-section, between an edit and the assertion about the queue
    that holds it, or on top of a sync already running.

    Seen before it was checked for: preflight's own writes armed the timer,
    preflight then rebound the calendar, and the timer fired at a target
    that no longer existed.

    Checked, not set - like the log level, and for the same reason. The
    value is the account owner's choice, and it is the suite's job to say
    what it needs rather than to hand back a profile configured differently
    from how it was found. Absent counts as on, because that is what the
    provider does with it.
    """
    value = (account.get("custom") or {}).get("syncOnChange")
    if value == "0":
        return
    shown = "the default" if value is None else f"{value} seconds"
    raise PreflightError(
        f"{account['accountName']} syncs a calendar after every change "
        f"({shown}), so this suite would not be the only thing syncing it.\n"
        f"  Set 'Synchronize a calendar after a change' to Off for this "
        f"account, in TbSync's account settings under Calendar - a sync "
        f"arriving in the middle of a section fails tests that are about "
        f"something else entirely."
    )


def _require_debug_logging():
    """Refuse to run unless the Event Log is capturing `debug`.

    The decoded WBXML of every request and response is logged at that level
    and nowhere else, so below it `wire()` returns an empty list. Assertions
    that something *was* sent then fail with a misleading message, and - far
    worse - the ones checking something was *not* sent pass without looking
    at anything.

    Checked, not set: the log level is the user's setting, and a suite that
    quietly rewrote it would hand back a profile in a state its owner did
    not choose.
    """
    level = (ok("storage.snapshot").get("tbsync.settings") or {}).get("logLevel")
    if level != 3:
        raise PreflightError(
            f"the Event Log level is {level}, so the wire is not being "
            f"captured.\n  Set it to 3 (Debug) in TbSync's Event Log tab - "
            f"below that, tests that assert nothing was sent pass without "
            f"testing anything."
        )


def _granted_account(accounts):
    """Which account the bridge may touch.

    Found by probing scope: rewriting an account's autosync interval with the
    value it already has is a no-op that succeeds only for the granted
    account. There is no read-only verb that reports the grant.
    """
    for account in accounts:
        reply = rpc(
            "setAutoSyncInterval",
            accountId=account["accountId"],
            minutes=account.get("autoSyncIntervalMinutes", 0),
        )
        if reply.get("ok"):
            return account
    return None


def _granted_folders(account_id):
    """kind -> folderId for the resources this account grants.

    Same trick: selecting a folder is refused unless it is one of the
    bridge's targets, and re-selecting an already-selected folder changes
    nothing.
    """
    found = {}
    for row in ok("getFolders", accountId=account_id)["folders"]:
        target_type = row.get("targetType")
        kind = {"contacts": "contacts", "calendars": "events", "tasks": "tasks"}.get(
            target_type
        )
        if not kind or kind in found:
            continue
        reply = rpc(
            "setFolderSelected",
            accountId=account_id,
            folderId=row["folderId"],
            selected=row["selected"],
        )
        if reply.get("ok"):
            found[kind] = row["folderId"]
    return found


def rediscover(session, tries=10):
    """Refresh `session.folders` from the host.

    Needed after a disconnect: the host forgets its folder records, and the
    provider re-announces them on connect. The ids have been stable in
    practice, but a run that assumes so would fail as "folder has vanished"
    rather than saying what happened.
    """
    for _ in range(tries):
        found = _granted_folders(session.account_id)
        if found:
            session.folders = found
            return found
        time.sleep(2)
    raise AssertionError(
        "the account announced no folders after reconnecting - the provider "
        "never re-ran its folder discovery"
    )


def select_resources(session, kinds, indent="  "):
    """Select exactly `kinds`; deselect every other granted resource.

    `syncAccount` is account-wide: it syncs every *selected* folder, so a
    resource left selected is pulled on every sync of every section whether
    the tests touch it or not. Section 2 needs only events, and was silently
    syncing contacts and tasks as well - three folders' work per sync, which
    is the load the delays were trying and failing to compensate for.

    So the state of every granted resource is asserted rather than assumed,
    and asserted again before each section, since sections need different
    ones.

    Deselecting deletes the local calendar or address book, which is why this
    belongs in a test suite and nowhere near production code. The cost is a
    fresh download when a later run wants that resource back; the saving is
    every sync in between.
    """
    changed = []
    for kind, folder_id in session.folders.items():
        row = session.folder(kind)
        want = kind in kinds
        if row["selected"] == want:
            continue
        ok(
            "setFolderSelected",
            accountId=session.account_id,
            folderId=folder_id,
            selected=want,
        )
        changed.append(f"{'+' if want else '-'}{kind}")
    if changed:
        print(f"{indent}selection {' '.join(changed)}")
        time.sleep(2)
    session.active = tuple(kinds)
    # Selecting a resource does not bind it - the target is created by the
    # first sync. Without this, the section that newly selects one queries a
    # folder that answers "not bound to a calendar yet", which reads like a
    # product fault and is really the suite going too fast.
    if any(c.startswith("+") for c in changed):
        session.sync()
        ensure_bound(session, kinds)
    return changed


def isolate(session, kinds, indent="  "):
    """Give the section a resource with no history: disconnect everything,
    then bind back only what it needs.

    A section is meant to be a complete statement. It was not: it inherited
    whatever the last one left - fixtures, queued edits, a sync key, an
    identity map - and every one of those has produced a failure that looked
    like a defect in the add-on and was not. The 2.1 wire assertion reading
    `SEND Add,Change`; section 3.4 passing or failing by history rather than
    by code; a whole night's tally counted against the wrong cause.

    Disconnecting is what makes it cheap to state: the provider clears
    `{synckey: "0", indexMap: []}` when a resource is disabled and the local
    store goes with it, so one action purges the queue, the map and the sync
    position together. Re-binding then pulls the folder fresh from the
    server, which is the only starting point the suite can describe.

    The cost is real - a bind, a bootstrap and a full pull per resource per
    section - and is the reason the runner rests on a 503 rather than
    hurrying.
    """
    for kind, folder_id in session.folders.items():
        row = session.folder(kind)
        if row["selected"]:
            ok(
                "setFolderSelected",
                accountId=session.account_id,
                folderId=folder_id,
                selected=False,
            )
    print(f"{indent}disconnected every resource")
    time.sleep(2)
    session.active = ()
    select_resources(session, kinds, indent=indent)


def drain_queues(session, kinds, tries=4, indent="  "):
    """Start every section from a queue it can name, rather than one it
    inherited.

    A section states what it creates. It does not state what it *starts
    from* - and the wire assertions depend on that silently, because
    `wire()` names every verb in a request. One edit left owed by an earlier
    section rides into the fixture's own push, the entry reads
    "SEND Add,Change" where the test asked for "SEND Add", and the failure
    reads as a regression in the code under test. It is not; it is state.

    Draining here rather than loosening the assertion is deliberate. A
    tolerant "some request mentioning Add" would also stop noticing a
    `<Change>` nobody asked for, which is exactly what those assertions
    exist to catch.

    Measured before it was written, as section 16: with one unrelated edit
    deliberately left queued the mixed request happens on every server and
    every attempt, and draining first makes the strict assertion hold on
    every server and every attempt.
    """
    for kind in kinds:
        owed = []
        for _ in range(tries):
            try:
                owed = session.changelog(kind)
            except AssertionError:
                # Not bound yet. The caller binds and syncs, which drains it.
                return
            if not owed:
                break
            print(f"{indent}draining {len(owed)} inherited {kind} edit(s)")
            session.sync()
        if owed:
            # Not fatal here: the section's own assertions are what say what
            # went wrong. But they will read as a defect in the add-on, so
            # the real reason belongs above them in the log.
            print(
                f"{indent}WARNING: {kind} still owes {len(owed)} edit(s) after "
                f"{tries} syncs - wire assertions below may see commands this "
                f"section did not send"
            )


def ensure_bound(session, kinds):
    for kind in kinds:
        row = session.folder(kind)
        if not row.get("targetID"):
            raise PreflightError(
                f"{kind} ({row['displayName']}) did not bind after a sync - "
                f"status {row['status']!r}, error {row.get('error')!r}."
            )


