# EAS-4-TbSync

## Exchange Active Sync (EAS) protocol v 16.1 Initial Support

Initial support introduced in this version:
- Basic Calendar/Contacts/Tasks editing and synchronization is working
Known problems:
- Calendar: Editing single event (with Attendees defined) in recurring series fails, changed item cannot be synchronized to Exchange
- ....

([EAS protocol v 14 support ends on 01.03.2026](https://techcommunity.microsoft.com/blog/exchange/exchange-online-activesync-device-support-update/4477997))

This provider add-on adds Exchange ActiveSync (EAS v2.5 & v14) synchronization capabilities to [TbSync](https://github.com/jobisoft/TbSync/).

More information can be found in the [wiki](https://github.com/jobisoft/EAS-4-TbSync/wiki/About:-Provider-for-Exchange-ActiveSync) of this repository

## Want to add or fix a localization?
To help translating this project, please visit [crowdin.com](https://crowdin.com/profile/jobisoft), where the localizations are managed. If you want to add a new language, just contact me and I will set it up.


## External data sources

* TbSync uses a [timezone mapping](https://github.com/mj1856/TimeZoneConverter/blob/master/src/TimeZoneConverter/Data/Mapping.csv.gz) provided by [Matt Johnson](https://github.com/mj1856)


## Icon sources and attributions

#### CC0 Public Domain
* [365_*.png] by [Microsoft / Wikimedia](https://commons.wikimedia.org/w/index.php?curid=21546299), converted from [SVG to PNG](https://ezgif.com/svg-to-png)

#### CC-BY 3.0
* [eas*.png] by [FatCow Web Hosting](https://www.iconfinder.com/icons/64484/exchange_ms_icon)
* [exchange_300.png] derived from [Microsoft Exchange Icon #270871](https://icon-library.net/icon/microsoft-exchange-icon-10.html), [resized](www.simpleimageresizer.com/)
