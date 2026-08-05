# EAS-4-TbSync working notes

Notes for the Exchange ActiveSync provider — the codec, the sync engine, and
the calendars it supplies. Written down as they are found, during audits,
protocol reading and debugging.

| | |
| --- | --- |
| [**How things work**](descriptions.html) | Descriptions of the provider's mechanisms, and why each is built the way it is. |
| [**Todos**](todos.html) | What is outstanding, with what is understood about each so far. |
| [**Manual test plan**](manual-test-plan.html) | What a person has to check before a release. The bridge-runnable half is `test/`, run with `npm test`. |

The host's own notes, including the bridge these tests are driven from, are at
[TbSync](https://jobisoft.github.io/TbSync/).
