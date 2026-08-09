# EAS-4-TbSync developer notes

The Exchange ActiveSync provider for
[TbSync](https://jobisoft.github.io/TbSync/). It speaks EAS 2.5, 14.0, 14.1
and 16.1 to Exchange, Office 365, Z-Push, Kopano and Grommunio, and supplies
Thunderbird with the calendars, task lists and address books it syncs.

These pages are for people working on the add-on. For user documentation see
the [wiki](https://github.com/jobisoft/EAS-4-TbSync/wiki/About:-Provider-for-Exchange-ActiveSync).

| | |
| --- | --- |
| [**How things work**](descriptions.html) | The mechanisms worth understanding before changing the sync engine or the codecs. |
| [**Manual test plan**](manual-test-plan.html) | What a person has to check before a release. The automated half is `test/`. |

The host's own mechanisms — the provider handshake, sessions, the E:AUTH
rule — are in the [TbSync notes](https://jobisoft.github.io/TbSync/).
