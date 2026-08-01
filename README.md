# EAS-4-TbSync

This provider add-on adds Exchange ActiveSync (EAS v2.5, v14.0, 14.1 & v16.1) synchronization capabilities to [TbSync](https://github.com/jobisoft/TbSync/).

More information can be found in the [wiki](https://github.com/jobisoft/EAS-4-TbSync/wiki/About:-Provider-for-Exchange-ActiveSync) of this repository

## Known limitations

### Recurring events from a foreign time zone (±1 hour DST drift)

A recurring series can drift by one hour around daylight-saving transitions
**only when all three of the following hold**:

1. the series was authored in a different time zone than your Thunderbird
   default zone, **and**
2. the server sends no usable `TimeZone` blob with the event — Z-Push, Kopano
   and Grommunio send an all-zero one, **and**
3. the occurrence falls in a different DST state than the series master.

In that situation the authoring zone is simply not transmitted, so the series
is anchored to your default zone and individual occurrences across a DST
boundary can be off by an hour. This is inherent to the protocol payload, not
something the client can recover.

Not affected:

* Exchange, on every protocol version including 16.0/16.1 — it sends a full
  `TimeZone` blob, which is matched back to the correct IANA zone (including
  disambiguation between same-offset zones such as Berlin/Paris via their DST
  transition dates). [MS-ASCAL] §2.2.2.44 lists the element as supported in
  every version, and Exchange Online populates it on 16.1; we send it too.
* Single, non-recurring events — their exact UTC instant is always preserved.
* All-day events on the inbound/display path — the calendar date is read back
  correctly for both server families (real-blob and no/empty-blob), in every
  host time zone. (Pushing an all-day event *back* to a non-Exchange server
  that ignores the `TimeZone` blob we send can still shift the date by a day
  for users east of UTC; real Exchange round-trips cleanly.)

## Want to add or fix a localization?
To help translating this project, please visit [crowdin.com](https://crowdin.com/profile/jobisoft), where the localizations are managed. If you want to add a new language, just contact me and I will set it up.


## External data sources

* TbSync uses a [timezone mapping](https://github.com/mattjohnsonpint/TimeZoneConverter/blob/master/src/TimeZoneConverter/Data/Mapping.csv.gz) provided by [Matt Johnson-Pint](https://github.com/mattjohnsonpint)


## Icon sources and attributions

#### CC0 Public Domain
* [365_*.png] by [Microsoft / Wikimedia](https://commons.wikimedia.org/w/index.php?curid=21546299), converted from [SVG to PNG](https://ezgif.com/svg-to-png)

#### CC-BY 3.0
* [eas*.png] by [FatCow Web Hosting](https://www.iconfinder.com/icons/64484/exchange_ms_icon)
* [exchange_300.png] derived from [Microsoft Exchange Icon #270871](https://icon-library.net/icon/microsoft-exchange-icon-10.html), [resized](www.simpleimageresizer.com/)
