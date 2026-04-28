# Vendored Files

This file lists files that were not created by this project and are maintained upstream elsewhere.

---

## calendar Experiments API

- **Files** : `/experiments/calendar/**` (subset; `calendar.calendars`, `calendar.items`, `calendar.timezones`, `calendar.provider`)
- **Source** : https://github.com/thunderbird/webext-experiments/tree/main/calendar/experiments/calendar
- **Commit** : b7f7cb3e76807903a785a03784d6e7df7b213f21
- **License** : MPL 2.0
- **Note** : Mirror byte-for-byte with the host copy at `tbsync-new/experiments/calendar/`. `calendarItemAction` / `calendarItemDetails` are intentionally not vendored. We don't author a custom calendar type (no `calendar_provider` entry in the manifest's top-level), but the `calendar_provider` experiment is still registered because its `onStartup` is what sets up the `resource://experiments-calendar-${uuid}/` substitution that the other parent scripts use to import `ext-calendar-utils.sys.mjs`.

---

## ical.min.js

- **File** : `/vendor/ical.min.js`
- **Source** : https://github.com/kewisch/ical.js/releases/download/v2.2.1/ical.min.js
- **Version** : v2.2.1
- **License** : MPL 2.0 (see header of [ical.min.js](./vendor/ical.min.js))

---

## i18n.mjs

- **File** : `/vendor/i18n/i18n.mjs`
- **Source** : https://github.com/thunderbird/webext-support/blob/6bbbf8ac2105d04c1b59083e8bd52e0046448ec7/modules/i18n/i18n.mjs
- **License** : MIT

---

## tbsync

- **Files** : `/vendor/tbsync/protocol.mjs`, `/vendor/tbsync/status.mjs`, `/vendor/tbsync/provider.mjs`
- **Source** : `tbsync-new/tbsync/` in the TbSync host repo (canonical). Mirror byte-for-byte; do not edit in this provider.
