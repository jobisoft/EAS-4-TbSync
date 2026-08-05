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
from bridge import rpc, ok

# The three resource kinds an EAS account must grant, as (label, targetType,
# the query used to probe the grant).
# Everything an EAS account can grant. Preflight requires only what the
# selected tests need, which is never all three: no section reads contacts.
KINDS = ("contacts", "events", "tasks")

# Smallest gap between two syncs of the same account. A single test can sync
# four or more times - a rebind is two, and `settle` retries up to three - so
# pacing the syncs matters more than pacing the tests. Exchange answers 503
# when pushed, and a 503 mid-resync has wiped a folder list before now.
#   TBSYNC_TEST_SYNC_GAP=0 npm test
MIN_SYNC_GAP_S = float(os.environ.get("TBSYNC_TEST_SYNC_GAP", "5"))


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
        """Pending entries for a folder.

        Read from the folder row, which is where the changelog lives. There
        is no `getChangelog` verb - an earlier helper called one anyway and
        returned [] when it failed, which is indistinguishable from an empty
        queue and quietly passed every assertion about a drained changelog.
        """
        self._active(kind)
        return self.folder(kind).get("changelog") or []

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

        Waits out `MIN_SYNC_GAP_S` first, so back-to-back syncs inside one
        test do not arrive as a burst.
        """
        waited = time.time() - self._last_sync
        if waited < MIN_SYNC_GAP_S:
            time.sleep(MIN_SYNC_GAP_S - waited)
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

        This replaces clearing. The log is the record of what the add-on did
        and is worth keeping whole - a section that fails is read afterwards,
        and a clear would have thrown that away.
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


def preflight(provider="eas", require=KINDS, bind=None):
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

    session = Session(granted, folders, (granted.get("custom") or {}).get("asversion"))
    session.mark()
    try:
        _bind(session, tuple(bind) if bind else require)
    except AssertionError as e:
        raise PreflightError(f"the initial sync failed.\n  {e}") from None
    return session


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


def ensure_bound(session, kinds):
    for kind in kinds:
        row = session.folder(kind)
        if not row.get("targetID"):
            raise PreflightError(
                f"{kind} ({row['displayName']}) did not bind after a sync - "
                f"status {row['status']!r}, error {row.get('error')!r}."
            )


def _bind(session, kinds):
    """Bring the account to the state this run needs, once at start-up."""
    select_resources(session, kinds)
    session.sync()
    ensure_bound(session, kinds)
