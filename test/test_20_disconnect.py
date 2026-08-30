"""20. Disconnect as a recovery path.

The manager's Disconnect button is gated on nothing: not on the provider
being alive, not on a sync being in flight. It is what a user reaches for
when something is wedged, so gating it on either would grey it out exactly
when it is needed - `SYNC_FOLDER` has no timeout, and an unanswered sync
would otherwise hold the account until the host add-on is reloaded.

What is asserted here is the whole contract, in order: a sync running,
disconnect while it runs, and then - the part that matters - connect again
and sync. A disconnect that left the account unusable would satisfy every
other check in this file.

The sync is started from a second thread because `syncAccount` over the
bridge waits for the sync to finish, which is exactly what is being
interrupted. And the interruption is asserted, not assumed: the first
version of 20.1 slept a fixed two seconds and passed - while the log showed
the sync had finished in 1.5 and the disconnect had aborted nothing.
"""

import threading
import time

import re

import harness
import session as session_mod
from bridge import ok, rpc
from harness import test

# Resources this section touches.
NEEDS = ("events",)



def _account_row(s):
    for a in ok("getState")["accounts"]:
        if a["accountId"] == s.account_id:
            return a
    raise AssertionError("the account has vanished")


def _sync_in_background(s):
    """Start a sync and hand back the thread, so a test can act while it runs.

    Failures are recorded rather than raised: a cancelled sync is *meant* to
    end badly for its caller, and what the account looks like afterwards is
    the assertion that matters.
    """
    outcome = {}

    def run():
        try:
            outcome["reply"] = rpc("syncAccount", accountId=s.account_id)
        except Exception as err:  # noqa: BLE001 - recorded on purpose
            outcome["error"] = err

    thread = threading.Thread(target=run, daemon=True)
    thread.start()
    return thread, outcome


def _syncing_now(s):
    transient = ok("getState").get("transient") or {}
    return s.account_id in transient.get("syncingAccounts", [])


@test("20.1", "disconnect while a sync is running - the sync stops with it")
def t_9_1(s):
    # Read while the account is still connected: 20.4 compares against it.
    # The value differs legitimately between accounts - one that
    # provisions never acquires it by the Settings route at all - so the
    # test is that the disconnect left it alone, not what it holds.
    s.device_ack_before = (_account_row(s).get("custom") or {}).get(
        "deviceInfoAcked"
    )

    thread, _outcome = _sync_in_background(s)

    # Catch the sync in the act rather than sleeping a fixed guess: this
    # account syncs an unchanged calendar in under two seconds, and a
    # disconnect that lands after the finish proves teardown, not abort.
    deadline = time.time() + 15
    while time.time() < deadline and thread.is_alive():
        if _syncing_now(s):
            break
        time.sleep(0.05)
    harness.true(
        thread.is_alive() and _syncing_now(s),
        "the sync was never observed running - it finished before the "
        "disconnect could interrupt it, so this run would prove nothing",
    )

    ok("setAccountEnabled", accountId=s.account_id, enabled=False)
    thread.join(timeout=60)
    harness.true(
        not thread.is_alive(),
        "the sync was still running a minute after the disconnect - its "
        "in-flight command was never settled, which is the lock this whole "
        "section exists to prevent",
    )
    harness.eq(_account_row(s)["enabled"], False, "account enabled")

    # The abort's own receipt. `abortAccountSync` writes this line only when
    # the account was still syncing when the abort ran - so its presence is
    # what says the disconnect interrupted a live sync rather than tidying
    # up after a finished one.
    messages = [e.get("message") or "" for e in s.log()]
    harness.true(
        any("Sync cancelled." in m for m in messages),
        "no 'Sync cancelled.' in the log - the disconnect ran after the "
        "sync had already finished, and the abort path was never exercised",
    )


@test("20.2", "nothing is left mid-sync, and a cancel is not an error")
def t_9_2(s):
    harness.eq(
        _account_row(s)["error"],
        None,
        "the account carries an error after being disconnected on purpose",
    )
    rows = ok("getFolders", accountId=s.account_id)["folders"]
    stuck = [f["folderId"] for f in rows if f.get("status") == "pending"]
    harness.eq(stuck, [], "folders left mid-sync")


@test("20.3", "connect and sync again - the account is usable, so nothing is stuck")
def t_9_3(s):
    s.mark()
    ok("setAccountEnabled", accountId=s.account_id, enabled=True)
    time.sleep(3)
    # Disconnecting cleared the folder records; the provider re-announces
    # them on connect, and the resource has to be selected again before it
    # can sync.
    session_mod.rediscover(s)
    session_mod.select_resources(s, NEEDS, indent="       ")
    s.sync()
    harness.eq(s.status("events"), "success", "folder status after reconnect")
    harness.eq(_account_row(s)["error"], None, "account error after reconnect")


@test("20.4", "the device stays introduced across a disconnect")
def t_9_4(s):
    """Disconnecting is a client-side act, and the partnership it leaves
    alone lives on the server.

    So the record of it survives, and reconnecting must not re-announce
    the same device: that would be a request telling the server something
    it already knows, on every reconnect, forever. This reads 20.3's
    window, where the account was re-enabled and synced.

    Neither carrier may do it - the standalone Settings command, nor a
    Provision body, which embeds the same element on 14.1 and above.
    """
    announced = []
    for entry in s.log():
        details = re.sub(r"\s+", " ", entry.get("details") or "")
        message = (entry.get("message") or "").lower()
        if "send" not in message:
            continue
        if "<DeviceInformation" in details and "<Set" in details:
            announced.append(message)

    harness.eq(announced, [], "device details re-announced after reconnect")
    harness.eq(
        (_account_row(s).get("custom") or {}).get("deviceInfoAcked"),
        s.device_ack_before,
        "the record of the acknowledgement changed across the disconnect",
    )
