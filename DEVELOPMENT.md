# Development

Building the add-on, and running its two test suites.

## Build

```
npm run build
```

Writes three xpis into `dist/`.

**Install `dist/dev.xpi` as a temporary add-on, not the `src/` folder.**
Thunderbird will happily load an unpacked directory, but then you are
running `src/` alone - without the `beta/` overlay. It also resolves paths
differently from a packed add-on, which the vendored calendar Experiment
does not survive. Using `dev.xpi` as a temporary install decouples the
running add-on from the local source folder and gets both right.

To refresh the running add-on, build and then reload it. The bridge can do
the reload on its own, so an external agent driving the bridge can apply
and test code changes.

| artifact | what is in it | what it is for |
| --- | --- | --- |
| `<name>_<version>_atn.xpi` | `src/` and nothing else | the release on addons.thunderbird.net |
| `<name>_<version>_beta.xpi` | `src/` plus the `beta/` overlay | the GitHub beta channel |
| `dev.xpi` | the beta build, under a name that never changes | local development |

### ATN versus beta

**Each build updates from where it came from.** An ATN install pulls its
updates from ATN; a beta install pulls them from GitHub. 

The two differ only by the `beta/` overlay, and at minimum that is
`beta/manifest.json`, which does two things:

- adds `browser_specific_settings.gecko.update_url`, pointing at the
  `updates.json` served from GitHub. That is what makes a beta install
  self-updating from GitHub. The ATN build carries no `update_url` at all,
  which is what leaves it updating from ATN - and it is not optional, since
  ATN **rejects** a manifest that carries one.
- overrides the add-on's name, so a beta install is distinguishable from a
  release one once installed.

The split is deliberately one-way: the ATN build never looks anywhere but
`src/`, so it has no exclude list that could be forgotten. **A beta-only
feature lives entirely in `beta/` and cannot reach ATN by accident.**

### Why `dev.xpi` exists

Its filename carries no version, so the path never moves and the add-on can
be reloaded at any time, including across a version bump. The other two
carry the version in their filenames, so a bump moves the file and any
fixed install path stops resolving.

Its add-on *name* carries the build time - `EAS-4-TbSync Beta (dev
2026-08-30 09:12)` - because the version alone cannot say which build is
loaded. Two builds minutes apart share a version, and an add-on that failed
to reload looks exactly like one that did. The Add-ons Manager shows that
name, so the answer is visible without unpacking anything.

`build.js` is mirrored byte-for-byte into TbSync, EAS-4-TbSync and
google-4-tbsync. Change one and re-copy it to the others:

```
diff -q TbSync/build.js EAS-4-TbSync/build.js
diff -q TbSync/build.js google-4-tbsync/build.js
```

## Unit tests

```
npm run test:unit
```

Node's own runner over `test/unit/*.test.mjs`. No Thunderbird, no network,
no account: these cover the codecs, the recurrence guard, the wire readers
and the version model against fixtures. Fast enough to run on every edit.

## The live suite

```
npm test                    every test that applies to the granted account
npm test -- 7               one section
npm test -- 7.3             one step
npm test -- 2 5             several
npm test -- --list          what would run, and what gates each  (instant)
npm test -- --watch <log>   attach to a run: last 20 lines, then follow
npm test -- --no-watch      start it and return, without attaching
```

This one drives a real account against a real server. It is the only thing
that can tell you what a server actually stored, and it is slow - a full
run is roughly an hour, most of it deliberate pacing.

### What it needs

Thunderbird running, TbSync's **Bridge** switched on, and an EAS account
granted to it. Preflight enforces the rest: it checks the bridge is
answering, that an account is granted, that it is an EAS one, and that it
grants the resources the selected sections need - refusing with what to fix
in the Bridge tab rather than testing half of what you asked for.

The one thing it cannot check for you: **use a test account, not a real
mailbox.** The suite creates, edits and deletes items in the granted
folders.

### What a run changes, and puts back

Preflight takes over a handful of account settings for the duration and
registers an undo for each: the event log goes to debug, sync-after-change
is switched off, conflict handling is set to server-wins, and the device
acknowledgement and policy key are cleared so the run exercises a real
connect. All of it is restored when the run ends - including when it is
stopped, see below.

Sections that push recurrence need the account's own recurrence option on;
the run says so and refuses up front rather than failing halfway.

### Watching, and stopping

A run is started in the background and writes its report as it goes:

```
  started in the background
  log    /path/to/test/runs/20260830-095945-all.log
  pid    1550474
  watch  tail -f /path/to/test/runs/20260830-095945-all.log
  stop   kill 1550474   (cleans up; kill -9 does not)
```

`npm test` then attaches to that log. **Ctrl-C stops the watching, not the
run** - the run is in its own session and never sees the signal. Come back
to it with `npm test -- --watch <log>`, which prints the last 20 lines
before following, so attaching an hour in tells you where it has got to.

`kill <pid>` is the clean stop: the run unwinds, puts the account's
settings back, and says so in the log. `kill -9` cannot be caught, so it
skips all of that and leaves the account as the run had it.

Each run also writes an event-log capture to `test/wire/` - the wire for a
failing run is usually the only thing that says what really happened.
Neither `test/runs/` nor `test/wire/` is tracked.

### How it is organised

One module per section, `test/test_<n>_<name>.py`, and **the number is the
position it runs in**. Sections do not chain: each clears and builds what it
needs, so a single section is a complete run. Steps within a section do
chain, so a failing test abandons the rest of *its* section and the run
carries on with the next one. A failing preflight stops the run outright -
it states the conditions the tests are judged against.

Tests can be gated on the protocol version (`VERSIONS = ("16",)`), and skip
with their reason on an account that cannot run them. A skip is always
printed; "0 failed" has to mean the same thing every run.
