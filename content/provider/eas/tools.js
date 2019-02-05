/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

eas.tools = {

    //gets the synclimit a.k.a. filtertype from pref settings, and checks if valid
    getFilterType: function () {
        let filterType = tbSync.prefSettings.getIntPref("eas.synclimit").toString();
        return (["0","4","5","6","7"].includes(filterType)) ? filterType : Services.prefs.getDefaultBranch("extensions.tbsync.").getIntPref("eas.synclimit").toString();
    },
    
    //read file from within the XPI package
    fetchFile: function (aURL, returnType = "Array") {
        return new Promise((resolve, reject) => {
            let uri = Services.io.newURI(aURL);
            let channel = Services.io.newChannelFromURI2(uri,
                                 null,
                                 Services.scriptSecurityManager.getSystemPrincipal(),
                                 null,
                                 Components.interfaces.nsILoadInfo.SEC_REQUIRE_SAME_ORIGIN_DATA_INHERITS,
                                 Components.interfaces.nsIContentPolicy.TYPE_OTHER);

            NetUtil.asyncFetch(channel, (inputStream, status) => {
                if (!Components.isSuccessCode(status)) {
                    reject(status);
                    return;
                }

                try {
                    let data = NetUtil.readInputStreamToString(inputStream, inputStream.available());
                    if (returnType == "Array") {
                        resolve(data.replace("\r","").split("\n"))
                    } else {
                        resolve(data);
                    }
                } catch (ex) {
                    reject(ex);
                }
            });
        });
    },
    
    //Date has a toISOString method, which returns the Date obj as extended ISO 8601,
    //however EAS MS-ASCAL uses compact/basic ISO 8601,
    dateToBasicISOString : function (date) {
        function pad(number) {
          if (number < 10) {
            return '0' + number;
          }
          return number.toString();
        }

        return pad(date.getUTCFullYear()) +
            pad(date.getUTCMonth() + 1) +
            pad(date.getUTCDate()) +
            'T' + 
            pad(date.getUTCHours()) +
            pad(date.getUTCMinutes()) +
            pad(date.getUTCSeconds()) +
            'Z';
    },

    //Save replacement for cal.createDateTime, which accepts compact/basic and also extended ISO 8601, 
    //cal.createDateTime only supports compact/basic
    createDateTime: function(str) {
        let datestring = str;
        if (str.indexOf("-") == 4) {
            //this looks like extended ISO 8601
            let tempDate = new Date(str);
            datestring = eas.tools.dateToBasicISOString(tempDate);
        }
        return cal.createDateTime(datestring);
    },    

    // Convert TB date to UTC and return it as  basic or extended ISO 8601  String
    getIsoUtcString: function(origdate, requireExtendedISO = false, fakeUTC = false) {
        let date = origdate.clone();
        //floating timezone cannot be converted to UTC (cause they float) - we have to overwrite it with the local timezone
        if (date.timezone.tzid == "floating") date.timezone = eas.defaultTimezoneInfo.timezone;
        //to get the UTC string we could use icalString (which does not work on allDayEvents, or calculate it from nativeTime)
        date.isDate = 0;
        let UTC = date.getInTimezone(eas.utcTimezone);        
        if (fakeUTC) UTC = date.clone();
        
        function pad(number) {
            if (number < 10) {
                return '0' + number;
            }
            return number;
        }
        
        if (requireExtendedISO) {
            return UTC.year + 
                    "-" + pad(UTC.month + 1 ) + 
                    "-" + pad(UTC.day) +
                    "T" + pad(UTC.hour) +
                    ":" + pad(UTC.minute) + 
                    ":" + pad(UTC.second) + 
                    "." + "000" +
                    "Z";            
        } else {            
            return UTC.icalString;
        }
    },
    





    //guess the IANA timezone (used by TB) based on the current offset (standard or daylight)
    guessTimezoneByCurrentOffset: function(curOffset, utcDateTime) {
        //if we only now the current offset and the current date, we need to actually try each TZ.
        let tzService = cal.getTimezoneService();

        //first try default tz
        let test = utcDateTime.getInTimezone(eas.defaultTimezoneInfo.timezone);
        tbSync.dump("Matching TZ via current offset: " + test.timezone.tzid + " @ " + curOffset, test.timezoneOffset/-60);
        if (test.timezoneOffset/-60 == curOffset) return test.timezone;
        
        //second try UTC
        test = utcDateTime.getInTimezone(eas.utcTimezone);
        tbSync.dump("Matching TZ via current offset: " + test.timezone.tzid + " @ " + curOffset, test.timezoneOffset/-60);
        if (test.timezoneOffset/-60 == curOffset) return test.timezone;
        
        //third try all others
        let enumerator = tzService.timezoneIds;
        while (enumerator.hasMore()) {
            let id = enumerator.getNext();
            let test = utcDateTime.getInTimezone(tzService.getTimezone(id));
            tbSync.dump("Matching TZ via current offset: " + test.timezone.tzid + " @ " + curOffset, test.timezoneOffset/-60);
            if (test.timezoneOffset/-60 == curOffset) return test.timezone;
        }
        
        //return default TZ as fallback
        return eas.defaultTimezoneInfo.timezone;
    },

  //guess the IANA timezone (used by TB) based on stdandard offset, daylight offset and standard name
    guessTimezoneByStdDstOffset: function(stdOffset, dstOffset, stdName = "") {
                    
            //get a list of all zones
            //alternativly use cal.fromRFC3339 - but this is only doing this:
            //https://dxr.mozilla.org/comm-central/source/calendar/base/modules/calProviderUtils.jsm

            //cache timezone data on first attempt
            if (eas.cachedTimezoneData === null) {
                eas.cachedTimezoneData = {};
                eas.cachedTimezoneData.iana = {};
                eas.cachedTimezoneData.abbreviations = {};
                eas.cachedTimezoneData.stdOffset = {};
                eas.cachedTimezoneData.bothOffsets = {};                    
                    
                let tzService = cal.getTimezoneService();

                //cache timezones data from internal IANA data
                let enumerator = tzService.timezoneIds;
                while (enumerator.hasMore()) {
                    let id = enumerator.getNext();
                    let timezone = tzService.getTimezone(id);
                    let tzInfo = eas.tools.getTimezoneInfo(timezone);

                    eas.cachedTimezoneData.bothOffsets[tzInfo.std.offset+":"+tzInfo.dst.offset] = timezone;
                    eas.cachedTimezoneData.stdOffset[tzInfo.std.offset] = timezone;

                    eas.cachedTimezoneData.abbreviations[tzInfo.std.abbreviation] = id;
                    eas.cachedTimezoneData.iana[id] = tzInfo;
                    
                    //tbSync.dump("TZ ("+ tzInfo.std.id + " :: " + tzInfo.dst.id +  " :: " + tzInfo.std.displayname + " :: " + tzInfo.dst.displayname + " :: " + tzInfo.std.offset + " :: " + tzInfo.dst.offset + ")", tzService.getTimezone(id));
                }

                //make sure, that UTC timezone is there
                eas.cachedTimezoneData.bothOffsets["0:0"] = eas.utcTimezone;

                //multiple TZ share the same offset and abbreviation, make sure the default timezone is present
                eas.cachedTimezoneData.abbreviations[eas.defaultTimezoneInfo.std.abbreviation] = eas.defaultTimezoneInfo.std.id;
                eas.cachedTimezoneData.bothOffsets[eas.defaultTimezoneInfo.std.offset+":"+eas.defaultTimezoneInfo.dst.offset] = eas.defaultTimezoneInfo.timezone;
                eas.cachedTimezoneData.stdOffset[eas.defaultTimezoneInfo.std.offset] = eas.defaultTimezoneInfo.timezone;
                
            }

            /*
                1. Try to find name in Windows names and map to IANA -> if found, does the stdOffset match? -> if so, done
                2. Try to parse our own format, split name and test each chunk for IANA -> if found, does the stdOffset match? -> if so, done
                3. Try if one of the chunks matches international code -> if found, does the stdOffset match? -> if so, done
                4. Fallback: Use just the offsets  */


            //check for windows timezone name
            if (eas.windowsTimezoneMap[stdName] && eas.cachedTimezoneData.iana[eas.windowsTimezoneMap[stdName]] && eas.cachedTimezoneData.iana[eas.windowsTimezoneMap[stdName]].std.offset == stdOffset ) {
                //the windows timezone maps multiple IANA zones to one (Berlin*, Rome, Bruessel)
                //check the windowsZoneName of the default TZ and of the winning, if they match, use default TZ
                //so Rome could win, even Berlin is the default IANA zone
                if (eas.defaultTimezoneInfo.std.windowsZoneName && eas.windowsTimezoneMap[stdName] != eas.defaultTimezoneInfo.std.id && eas.cachedTimezoneData.iana[eas.windowsTimezoneMap[stdName]].std.offset == eas.defaultTimezoneInfo.std.offset && stdName == eas.defaultTimezoneInfo.std.windowsZoneName) {
                    tbSync.dump("Timezone matched via windows timezone name ("+stdName+") with default TZ overtake", eas.windowsTimezoneMap[stdName] + " -> " + eas.defaultTimezoneInfo.std.id);
                    return eas.defaultTimezoneInfo.timezone;
                }
                
                tbSync.dump("Timezone matched via windows timezone name ("+stdName+")", eas.windowsTimezoneMap[stdName]);
                return eas.cachedTimezoneData.iana[eas.windowsTimezoneMap[stdName]].timezone;
            }

            let parts = stdName.replace(/[;,()\[\]]/g," ").split(" ");
            for (let i = 0; i < parts.length; i++) {
                //check for IANA
                if (eas.cachedTimezoneData.iana[parts[i]] && eas.cachedTimezoneData.iana[parts[i]].std.offset == stdOffset) {
                    tbSync.dump("Timezone matched via IANA", parts[i]);
                    return eas.cachedTimezoneData.iana[parts[i]].timezone;
                }

                //check for international abbreviation for standard period (CET, CAT, ...)
                if (eas.cachedTimezoneData.abbreviations[parts[i]] && eas.cachedTimezoneData.iana[eas.cachedTimezoneData.abbreviations[parts[i]]] && eas.cachedTimezoneData.iana[eas.cachedTimezoneData.abbreviations[parts[i]]].std.offset == stdOffset) {
                    tbSync.dump("Timezone matched via international abbreviation (" + parts[i] +")", eas.cachedTimezoneData.abbreviations[parts[i]]);
                    return eas.cachedTimezoneData.iana[eas.cachedTimezoneData.abbreviations[parts[i]]].timezone;
                }
            }

            //fallback to zone based on stdOffset and dstOffset, if we have that cached
            if (eas.cachedTimezoneData.bothOffsets[stdOffset+":"+dstOffset]) {
                tbSync.dump("Timezone matched via both offsets (std:" + stdOffset +", dst:" + dstOffset + ")", eas.cachedTimezoneData.bothOffsets[stdOffset+":"+dstOffset].tzid);
                return eas.cachedTimezoneData.bothOffsets[stdOffset+":"+dstOffset];
            }

            //fallback to zone based on stdOffset only, if we have that cached
            if (eas.cachedTimezoneData.stdOffset[stdOffset]) {
                tbSync.dump("Timezone matched via std offset (" + stdOffset +")", eas.cachedTimezoneData.stdOffset[stdOffset].tzid);
                return eas.cachedTimezoneData.stdOffset[stdOffset];
            }
            
            //return default timezone, if everything else fails
            tbSync.dump("Timezone could not be matched via offsets (std:" + stdOffset +", dst:" + dstOffset + "), using default timezone", eas.defaultTimezoneInfo.std.id);
            return eas.defaultTimezoneInfo.timezone;
    },
    
    //extract standard and daylight timezone data
    getTimezoneInfo: function (timezone) {        
        let tzInfo = {};

        tzInfo.std = eas.tools.getTimezoneInfoObject(timezone, "standard");
        tzInfo.dst = eas.tools.getTimezoneInfoObject(timezone, "daylight");
        
        if (tzInfo.dst === null) tzInfo.dst  = tzInfo.std;        

        tzInfo.timezone = timezone;
        return tzInfo;
    },

     //get timezone info for standard/daylight
    getTimezoneInfoObject: function (timezone, standardOrDaylight) {       
        
        //handle UTC
        if (timezone.isUTC) {
            let obj = {}
            obj.id = "UTC";
            obj.offset = 0;
            obj.abbreviation = "UTC";
            obj.displayname = "Coordinated Universal Time (UTC)";
            return obj;
        }
                
        //we could parse the icalstring by ourself, but I wanted to use ICAL.parse - TODO try catch
        let info = ICAL.parse("BEGIN:VCALENDAR\r\n" + timezone.icalComponent.toString() + "\r\nEND:VCALENDAR");
        let comp = new ICAL.Component(info);
        let vtimezone =comp.getFirstSubcomponent("vtimezone");
        let id = vtimezone.getFirstPropertyValue("tzid").toString();
        let zone = vtimezone.getFirstSubcomponent(standardOrDaylight);

        if (zone) { 
            let obj = {};
            obj.id = id;
            
            //get offset
            let utcOffset = zone.getFirstPropertyValue("tzoffsetto").toString();
            let o = parseInt(utcOffset.replace(":","")); //-330 =  - 3h 30min
            let h = Math.floor(o / 100); //-3 -> -180min
            let m = o - (h*100) //-330 - -300 = -30
            obj.offset = -1*((h*60) + m);

            //get international abbreviation (CEST, CET, CAT ... )
            obj.abbreviation = "";
            try {
                obj.abbreviation = zone.getFirstPropertyValue("tzname").toString();
            } catch(e) {
                tbSync.dump("Failed TZ", timezone.icalComponent.toString());
            }
            
            //get displayname
            obj.displayname = /*"("+utcOffset+") " +*/ obj.id;// + ", " + obj.abbreviation;
                
            //get DST switch date
            let rrule = zone.getFirstPropertyValue("rrule");
            let dtstart = zone.getFirstPropertyValue("dtstart");
            if (rrule && dtstart) {
                /*

                    THE switchdate PART OF THE OBJECT IS MICROSOFT SPECIFIC, EVERYTHING ELSE IS THUNDERBIRD GENERIC, I LET IT SIT HERE ANYHOW
                    
                    https://msdn.microsoft.com/en-us/library/windows/desktop/ms725481(v=vs.85).aspx

                    To select the correct day in the month, set the wYear member to zero, the wHour and wMinute members to
                    the transition time, the wDayOfWeek member to the appropriate weekday, and the wDay member to indicate
                    the occurrence of the day of the week within the month (1 to 5, where 5 indicates the final occurrence during the
                    month if that day of the week does not occur 5 times).

                    Using this notation, specify 02:00 on the first Sunday in April as follows: 
                        wHour = 2, wMonth = 4, wDayOfWeek = 0, wDay = 1. 
                    Specify 02:00 on the last Thursday in October as follows: 
                        wHour = 2, wMonth = 10, wDayOfWeek = 4, wDay = 5.
                        
                    So we have to parse the RRULE to exract wDay
                    RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=10 */         

                let parts =rrule.toString().split(";");
                let rules = {};
                for (let i = 0; i< parts.length; i++) {
                    let sub = parts[i].split("=");
                    if (sub.length == 2) rules[sub[0]] = sub[1];
                }
                
                if (rules.FREQ == "YEARLY" && rules.BYDAY && rules.BYMONTH && rules.BYDAY.length > 2) {
                    obj.switchdate = {};
                    obj.switchdate.month = parseInt(rules.BYMONTH);

                    let days = ["SU","MO","TU","WE","TH","FR","SA"];
                    obj.switchdate.dayOfWeek = days.indexOf(rules.BYDAY.substring(rules.BYDAY.length-2));                
                    obj.switchdate.weekOfMonth = parseInt(rules.BYDAY.substring(0, rules.BYDAY.length-2));
                    if (obj.switchdate.weekOfMonth<0 || obj.switchdate.weekOfMonth>5) obj.switchdate.weekOfMonth = 5;

                    //get switch time from dtstart
                    let dttime = eas.tools.createDateTime(dtstart.toString());
                    obj.switchdate.hour = dttime.hour;
                    obj.switchdate.minute = dttime.minute;
                    obj.switchdate.second = dttime.second;                                    
                }            
            }

            return obj;
        }
        return null;
    },
    




    
    
    promisifyAddressbook: function (addressbook) {
    /* 
        Return obj with identical interface to promisifyCalendar. But we currently do not need a promise. 
            adoptItem(card)
            modifyItem(newcard, existingcard)
            deleteItem(card)
            getItem(id)

        Avail API:
            addressBook.modifyCard(card);
            addressBook.getCardFromProperty("localId", ClientId, false);
            addressBook.deleteCards(cardsToDelete);
            card.setProperty('ServerId', ServerId);
    */
        let apiWrapper = {
            adoptItem: function (item) { 
                /* add card to addressbook */
                addressbook.addCard(item.card);
            },

            modifyItem: function (newitem, existingitem) {
                /* modify card */
                addressbook.modifyCard(newitem.card);
            },

            deleteItem: function (item) {
                /* remove card from addressBook */
                let cardsToDelete = Components.classes["@mozilla.org/array;1"].createInstance(Components.interfaces.nsIMutableArray);
                cardsToDelete.appendElement(item.card, false);
                addressbook.deleteCards(cardsToDelete);
            },

            getItem: function (searchId) {
                /* return array of items matching */
                let items = [];
                let card = addressbook.getCardFromProperty("TBSYNCID", searchId, true); //3rd param enables case sensitivity
                
                if (card) {
                    items.push(tbSync.eas.sync.Contacts.createItem(card));
                }
                
                return items;
            }
        };
    
        return apiWrapper;
    },
    
}

//TODO: Invites
/*
if (tbSync.lightningIsAvailable()) {
    cal.itip.checkAndSendOrigial = cal.itip.checkAndSend;
    cal.itip.checkAndSend = function(aOpType, aItem, aOriginalItem) {
        //if this item is added_by_user, do not call checkAndSend yet, because the UID is wrong, we need to sync first to get the correct ID - TODO
        tbSync.dump("cal.checkAndSend", aOpType);
        cal.itip.checkAndSendOrigial(aOpType, aItem, aOriginalItem);
    }
}
*/
