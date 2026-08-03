# Automated tests

An automated test harness for the pure-logic parts of this add-on: EAS
WBXML `<ApplicationData>` ↔ iCal/vCard conversion, the WBXML request/
response helpers, and the push/pull orchestration in `sync-runner.mjs`.
It does not (and cannot, from Node) drive Thunderbird itself - see
"Known gap" below.

## How to run

```
npm ci
npm test
```

`vitest run`, zero config beyond `package.json`'s `test` script. CI runs
it on a Node 20.x/22.x matrix via `.github/workflows/test.yml`.
`Dockerfile.test` reproduces the same run locally:

```
docker build -f Dockerfile.test -t eas-4-tbsync-tests .
docker run --rm eas-4-tbsync-tests
```

## What's covered

- `calendar-codec.mjs`: inbound `<ApplicationData>` → iCal (single event,
  whole-series RRULE, 16.1 per-occurrence `InstanceId` edits/deletes,
  14.x embedded `<Exceptions>`) and the outbound reverse, round-tripped
  through the real WBXML encoder/decoder (`wbxml.mjs`), not a stub.
- `task-codec.mjs` / `contact-codec.mjs`: field mapping and the
  merge-aware "a delta without this tag leaves the existing value
  untouched" behavior each codec relies on.
- `wbxml-helpers.mjs` / `timezone-blob.mjs`: the small parsing utilities
  everything else depends on.
- `sync-runner.mjs`'s push orchestration: `appendCommands` and
  `applyResponses` are exported test-only (every production caller still
  reaches them through the normal `pushPhase`/`runOneSync` flow) so two
  request/response shapes can be driven directly without mocking the
  network.

Two of those double as regression evidence for filed issues:

- `sync-runner.bundled-instance-change.test.mjs` demonstrates that a
  modified recurring item's master `<Change>` and its per-occurrence
  `InstanceId` `<Change>` get sent under the same `ServerId` in one
  `<Sync>` request, and documents today's retry/changelog behavior when
  a server rejects both (see #334).
- `calendar-codec.inbound-exceptions.test.mjs` documents that a second,
  status-only touch of an already-synced recurrence exception drops its
  Subject/Start/End locally, because the exception is rebuilt from an
  empty component on every update with no inheritance from the prior
  state (see #342).

Both are written as characterization tests - they assert today's actual
behavior on purpose, so a fix has to touch the test file rather than
land unverified.

## Tooling choices

Real WBXML round-trips and real captured wire fixtures wherever
practical, instead of hand-written XML approximations - `createWBXML`/
`decodeWBXML` are exercised directly. [vitest](https://vitest.dev) is
the only dependency this adds; `vi.mock`/`vi.hoisted` cover the one seam
that needs mocking (the network call in the `sync-runner` push test).

## Known gap: the WebExtension host boundary

`timezone-mapping.mjs`'s `ensureLoaded()` needs
`messenger.calendar.timezones.*` (Thunderbird's own calendar backend),
unavailable outside Thunderbird. `tests/support/webext-shim.mjs` shims
just enough for the common UTC-only path - `"UTC"` short-circuits before
ever needing real per-zone data. Tests requiring a specific non-UTC IANA
zone resolution are out of scope here and stay manual-test territory.

## Scope of this branch

This covers only code paths that exist on `master` as of this branch.
MeetingResponse (Accept/Decline/Tentative, #339) isn't merged yet, so
tests for it aren't part of this PR; they'll follow as a separate PR
once #339 lands.
