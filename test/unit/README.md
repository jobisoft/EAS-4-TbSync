# The unit layer

`npm run test:unit` — plain `node --test`, no dependencies, no network,
no Thunderbird.

Two layers share `test/`, with one boundary:

- **This layer** owns pure logic: the codecs (calendar, task, contact),
  the WBXML helpers and encoder round-trip, the TimeZone blob packing.
  Anything provable from inputs alone.
- **The live suite** (`python3 run.py`, see the section docstrings) owns
  everything behavioral: real servers, the bridge, folder state,
  migrations. If a case needs a server's opinion, it goes there.

Fixture policy, two idioms only:

- `el(tag, value)` builds synthetic nodes where no capture exists.
- `parseAdNode(xml)` turns a wire capture **pasted verbatim from the
  Event Log's decoded WBXML** into codec input. Prefer it: a user's log
  becomes a regression fixture unchanged.

`support/webext-env.mjs` stands in for the host environment (never for
code under test) with UTC-only timezone resolution. A test needing a
real named zone belongs to the live suite (section 4). Note the codec
does not always *fail* without the zone — `writeDateProp` deliberately
falls back to UTC — so an assertion that pins a converted wall-clock
value under this environment is pinning the UTC-fallback form, and must
say so where it does.

Much of this layer's coverage was contributed in PR #345 (tomaskovacik)
and ported from vitest to `node --test`; live-server captures in the
fixtures are his.
