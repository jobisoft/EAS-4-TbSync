# Privacy policy

Last updated: 2026-11-27

EAS-4-TbSync synchronizes contacts, calendars and tasks between Thunderbird and an Exchange ActiveSync server chosen by the user. The developer runs no server for it and receives no data. It has no telemetry, analytics or advertising.

## Data and transmission

After the user adds an account, the add-on sends synchronized data (contacts, events, tasks), GAL and free/busy queries, credentials, a random device ID and the user agent only to:

- the user's EAS server, entered by the user or found through Autodiscover at the email domain. Plain `http://` is upgraded to `https://`, unless the user enters an explicit port or a loopback address.
- `login.microsoftonline.com`, for Microsoft 365 sign-in through OAuth 2.0. The add-on never sees the Microsoft password.

No other host is contacted, and no remote code is loaded.

## Local storage

Account settings, the password or OAuth refresh token and synchronization state are stored in the add-on's `storage.local` in the Thunderbird profile. Credentials are stored unencrypted. Technical messages go to the local TbSync event log, which is never sent anywhere automatically.

## User control

The user can disable or delete an account at any time in TbSync, and uninstalling the add-on removes its stored data. Data held by the server operator or Microsoft is subject to their policies.

## Contact

<https://github.com/jobisoft/EAS-4-TbSync/issues>
