/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

const cal = TbSync.lightning.cal;
const ICAL = TbSync.lightning.ICAL;

var Calendar = {

    // --------------------------------------------------------------------------- //
    // Read WBXML and set Thunderbird item
    // --------------------------------------------------------------------------- //
    setThunderbirdItemFromWbxml: function (tbItem, data, id, syncdata) {
        
        let item = tbItem instanceof TbSync.lightning.TbItem ? tbItem.nativeItem : tbItem;
        
        let asversion = syncdata.accountData.getAccountProperty("asversion");
        item.id = id;
        let easTZ = new eas.tools.TimeZoneDataStructure();

        eas.sync.setItemSubject(item, syncdata, data);
        eas.sync.setItemLocation(item, syncdata, data);
        eas.sync.setItemCategories(item, syncdata, data);
        eas.sync.setItemBody(item, syncdata, data);

        //timezone
        let stdOffset = eas.defaultTimezoneInfo.std.offset;
        let dstOffset = eas.defaultTimezoneInfo.dst.offset;

        if (data.TimeZone) {
            if (data.TimeZone == "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==") {
                TbSync.dump("Recieve TZ", "No timezone data received, using local default timezone.");
            } else {
                //load timezone struct into EAS TimeZone object
                easTZ.easTimeZone64 = data.TimeZone;
                if (TbSync.prefs.getIntPref("log.userdatalevel") > 2) TbSync.dump("Recieve TZ", item.title + easTZ.toString());
                stdOffset = easTZ.utcOffset;
                dstOffset = easTZ.daylightBias + easTZ.utcOffset;
            }
        }

        if (data.StartTime) {
            let utc = cal.createDateTime(data.StartTime); //format "19800101T000000Z" - UTC
            item.startDate = utc.getInTimezone(eas.tools.guessTimezoneByStdDstOffset(stdOffset, dstOffset, easTZ.standardName));
            if (data.AllDayEvent && data.AllDayEvent == "1") {
                item.startDate.timezone = (cal.dtz && cal.dtz.floating) ? cal.dtz.floating : cal.floating();
                item.startDate.isDate = true;
            }
        }

        if (data.EndTime) {
            let utc = cal.createDateTime(data.EndTime);
            item.endDate = utc.getInTimezone(eas.tools.guessTimezoneByStdDstOffset(stdOffset, dstOffset, easTZ.standardName));
            if (data.AllDayEvent && data.AllDayEvent == "1") {
                item.endDate.timezone = (cal.dtz && cal.dtz.floating) ? cal.dtz.floating : cal.floating();
                item.endDate.isDate = true;
            }
        }

        //stamp time cannot be set and it is not needed, an updated version is only send to the server, if there was a change, so stamp will be updated


        //EAS Reminder
        item.clearAlarms();
        if (data.Reminder && data.StartTime) {
            let alarm = cal.createAlarm();
            alarm.related = Components.interfaces.calIAlarm.ALARM_RELATED_START;
            alarm.offset = cal.createDuration();
            alarm.offset.inSeconds = (0-parseInt(data.Reminder)*60);
            alarm.action ="DISPLAY";
            item.addAlarm(alarm);
            
            let alarmData = cal.alarms.calculateAlarmDate(item, alarm);
            let startDate = cal.createDateTime(data.StartTime);
            let nowDate = eas.tools.getNowUTC();
            if (startDate.compare(nowDate) < 0) {
                // Mark alarm as ACK if in the past.
                item.alarmLastAck = nowDate;
            }
        }

        eas.sync.mapEasPropertyToThunderbird ("BusyStatus", "TRANSP", data, item);
        eas.sync.mapEasPropertyToThunderbird ("Sensitivity", "CLASS", data, item);

        if (data.ResponseType) {
            //store original EAS value 
            item.setProperty("X-EAS-ResponseType", eas.xmltools.checkString(data.ResponseType, "0")); //some server send empty ResponseType ???
        }

        //Attendees - remove all Attendees and re-add the ones from XML
        item.removeAllAttendees();
        if (data.Attendees && data.Attendees.Attendee) {
            let att = [];
            if (Array.isArray(data.Attendees.Attendee)) att = data.Attendees.Attendee;
            else att.push(data.Attendees.Attendee);
            for (let i = 0; i < att.length; i++) {
                if (att[i].Email && eas.tools.isString(att[i].Email) && att[i].Name) { //req.

                    let attendee = cal.createAttendee();

                    //is this attendee the local EAS user?
                    let isSelf = (att[i].Email == syncdata.accountData.getAccountProperty("user"));
                    
                    attendee["id"] = cal.email.prependMailTo(att[i].Email);
                    attendee["commonName"] = att[i].Name;
                    //default is "FALSE", only if THIS attendee isSelf, use ResponseRequested (we cannot respond for other attendee) - ResponseType is not send back to the server, it is just a local information
                    attendee["rsvp"] = (isSelf && data.ResponseRequested) ? "TRUE" : "FALSE";		

                    //not supported in 2.5
                    switch (att[i].AttendeeType) {
                        case "1": //required
                            attendee["role"] = "REQ-PARTICIPANT";
                            attendee["userType"] = "INDIVIDUAL";
                            break;
                        case "2": //optional
                            attendee["role"] = "OPT-PARTICIPANT";
                            attendee["userType"] = "INDIVIDUAL";
                            break;
                        default : //resource or unknown
                            attendee["role"] = "NON-PARTICIPANT";
                            attendee["userType"] = "RESOURCE";
                            break;
                    }

                    //not supported in 2.5 - if attendeeStatus is missing, check if this isSelf and there is a ResponseType
                    if (att[i].AttendeeStatus)
                        attendee["participationStatus"] = eas.sync.MAP_EAS2TB.ATTENDEESTATUS[att[i].AttendeeStatus];
                    else if (isSelf && data.ResponseType) 
                        attendee["participationStatus"] = eas.sync.MAP_EAS2TB.ATTENDEESTATUS[data.ResponseType];
                    else 
                        attendee["participationStatus"] = "NEEDS-ACTION";

                    // status  : [NEEDS-ACTION, ACCEPTED, DECLINED, TENTATIVE, DELEGATED, COMPLETED, IN-PROCESS]
                    // rolemap : [REQ-PARTICIPANT, OPT-PARTICIPANT, NON-PARTICIPANT, CHAIR]
                    // typemap : [INDIVIDUAL, GROUP, RESOURCE, ROOM]

                    // Add attendee to event
                    item.addAttendee(attendee);
                } else {
                    TbSync.eventlog.add("info", syncdata, "Attendee without required name and/or email found. Skipped.");
                }
            }
        }
        
        if (data.OrganizerName && data.OrganizerEmail && eas.tools.isString(data.OrganizerEmail)) {
            //Organizer
            let organizer = cal.createAttendee();
            organizer.id = cal.email.prependMailTo(data.OrganizerEmail);
            organizer.commonName = data.OrganizerName;
            organizer.rsvp = "FALSE";
            organizer.role = "CHAIR";
            organizer.userType = null;
            organizer.participationStatus = "ACCEPTED";
            organizer.isOrganizer = true;
            item.organizer = organizer;
        }

        eas.sync.setItemRecurrence(item, syncdata, data, eas.tools.guessTimezoneByStdDstOffset(stdOffset, dstOffset, easTZ.standardName));
        
        // BusyStatus is always representing the status of the current user in terms of availability.
        // It has nothing to do with the status of a meeting. The user could be just the organizer, but does not need to attend, so he would be free.
        // The correct map is between BusyStatus and TRANSP (show time as avail, busy, unset)
        // A new event always sets TRANSP to busy, so unset is indeed a good way to store Tentiative
        // However:
        //  - EAS Meetingstatus only knows ACTIVE or CANCELLED, but not CONFIRMED or TENTATIVE
        //  - TB STATUS has UNSET, CONFIRMED, TENTATIVE, CANCELLED
        //  -> Special case: User sets BusyStatus to TENTIATIVE -> TRANSP is unset and also set STATUS to TENTATIVE
        // The TB STATUS is the correct map for EAS Meetingstatus and should be unset, if it is not a meeting EXCEPT if set to TENTATIVE
        let tbStatus = (data.BusyStatus && data.BusyStatus == "1" ?  "TENTATIVE" : null);
        
        if (data.MeetingStatus) {
            //store original EAS value 
            item.setProperty("X-EAS-MeetingStatus", data.MeetingStatus);
            //bitwise representation for Meeting, Received, Cancelled:
            let M = data.MeetingStatus & 0x1;
            let R = data.MeetingStatus & 0x2;
            let C = data.MeetingStatus & 0x4;

            // We can map M+C to TB STATUS (TENTATIVE, CONFIRMED, CANCELLED, unset).
            if (M) {
                if (C) tbStatus = "CANCELLED";
                else if (!tbStatus) tbStatus = "CONFIRMED"; // do not override "TENTIATIVE"
            }
            
            //we can also use the R information, to update our fallbackOrganizerName
            if (!R && data.OrganizerName) syncdata.target.calendar.setProperty("fallbackOrganizerName", data.OrganizerName);            
        }

        if (tbStatus) item.setProperty("STATUS", tbStatus)
        else item.deleteProperty("STATUS");

        //TODO: attachements (needs EAS 16.0!)
    },









    // --------------------------------------------------------------------------- //
    //read TB event and return its data as WBXML
    // --------------------------------------------------------------------------- //
    getWbxmlFromThunderbirdItem: function (tbItem, syncdata, isException = false) {
        let item = tbItem instanceof TbSync.lightning.TbItem ? tbItem.nativeItem : tbItem;

        let asversion = syncdata.accountData.getAccountProperty("asversion");
        let wbxml = eas.wbxmltools.createWBXML("", syncdata.type); //init wbxml with "" and not with precodes, and set initial codepage
        let nowDate = new Date();

        /*
         *  We do not use ghosting, that means, if we do not include a value in CHANGE, it is removed from the server. 
         *  However, this does not seem to work on all fields. Furthermore, we need to include any (empty) container to blank its childs.
         */

        //Order of tags taken from https://msdn.microsoft.com/en-us/library/dn338917(v=exchg.80).aspx
        
        //timezone
        if (!isException) {
            let easTZ = new eas.tools.TimeZoneDataStructure();

            //if there is no end and no start (or both are floating) use default timezone info
            let tzInfo = null;
            if (item.startDate && item.startDate.timezone.tzid != "floating") tzInfo = eas.tools.getTimezoneInfo(item.startDate.timezone);
            else if (item.endDate && item.endDate.timezone.tzid != "floating") tzInfo = eas.tools.getTimezoneInfo(item.endDate.timezone);
            if (!tzInfo) tzInfo = eas.defaultTimezoneInfo;
            
            easTZ.utcOffset =   tzInfo.std.offset;
            easTZ.standardBias = 0;
            easTZ.daylightBias =  tzInfo.dst.offset -  tzInfo.std.offset;

            easTZ.standardName = eas.ianaToWindowsTimezoneMap.hasOwnProperty(tzInfo.std.displayname) ? eas.ianaToWindowsTimezoneMap[tzInfo.std.displayname] : tzInfo.std.displayname;
            easTZ.daylightName = eas.ianaToWindowsTimezoneMap.hasOwnProperty(tzInfo.dst.displayname) ? eas.ianaToWindowsTimezoneMap[tzInfo.dst.displayname] : tzInfo.dst.displayname;

            if (tzInfo.std.switchdate && tzInfo.dst.switchdate) {
                easTZ.standardDate.wMonth = tzInfo.std.switchdate.month;
                easTZ.standardDate.wDay = tzInfo.std.switchdate.weekOfMonth;
                easTZ.standardDate.wDayOfWeek = tzInfo.std.switchdate.dayOfWeek;
                easTZ.standardDate.wHour = tzInfo.std.switchdate.hour;
                easTZ.standardDate.wMinute = tzInfo.std.switchdate.minute;
                easTZ.standardDate.wSecond = tzInfo.std.switchdate.second;
                
                easTZ.daylightDate.wMonth = tzInfo.dst.switchdate.month;
                easTZ.daylightDate.wDay = tzInfo.dst.switchdate.weekOfMonth;
                easTZ.daylightDate.wDayOfWeek = tzInfo.dst.switchdate.dayOfWeek;
                easTZ.daylightDate.wHour = tzInfo.dst.switchdate.hour;
                easTZ.daylightDate.wMinute = tzInfo.dst.switchdate.minute;
                easTZ.daylightDate.wSecond = tzInfo.dst.switchdate.second;
            }
            
            wbxml.atag("TimeZone", easTZ.easTimeZone64);
             if (TbSync.prefs.getIntPref("log.userdatalevel") > 2) TbSync.dump("Send TZ", item.title + easTZ.toString());
        }
        
        //AllDayEvent (for simplicity, we always send a value)
        wbxml.atag("AllDayEvent", (item.startDate && item.startDate.isDate && item.endDate && item.endDate.isDate) ? "1" : "0");

        //Body
        wbxml.append(eas.sync.getItemBody(item, syncdata));

        //BusyStatus (Free, Tentative, Busy) is taken from TRANSP (busy, free, unset=tentative)
        //However if STATUS is set to TENTATIVE, overide TRANSP and set BusyStatus to TENTATIVE
        if (item.hasProperty("STATUS") && item.getProperty("STATUS") == "TENTATIVE") {
            wbxml.atag("BusyStatus","1");
        } else {
            wbxml.atag("BusyStatus", eas.sync.mapThunderbirdPropertyToEas("TRANSP", "BusyStatus", item));
        }
        
        //Organizer
        if (!isException) {
            if (item.organizer && item.organizer.commonName) wbxml.atag("OrganizerName", item.organizer.commonName);
            if (item.organizer && item.organizer.id) wbxml.atag("OrganizerEmail",  cal.email.removeMailTo(item.organizer.id));
        }

        //DtStamp in UTC
        wbxml.atag("DtStamp", item.stampTime ? eas.tools.getIsoUtcString(item.stampTime) : eas.tools.dateToBasicISOString(nowDate));

        //EndTime in UTC
        wbxml.atag("EndTime", item.endDate ? eas.tools.getIsoUtcString(item.endDate) : eas.tools.dateToBasicISOString(nowDate));
        
        //Location
        wbxml.atag("Location", (item.hasProperty("location")) ? item.getProperty("location") : "");

        //EAS Reminder (TB getAlarms) - at least with zpush blanking by omitting works, horde does not work
        let alarms = item.getAlarms({});
        if (alarms.length>0) {

            let reminder = -1;
            if (alarms[0].offset !== null) {
                reminder = 0 - alarms[0].offset.inSeconds/60;
            } else if (item.startDate) {
                let timeDiff =item.startDate.getInTimezone(eas.utcTimezone).subtractDate(alarms[0].alarmDate.getInTimezone(eas.utcTimezone));     
                reminder = timeDiff.inSeconds/60;
                TbSync.eventlog.add("info", syncdata, "Converting absolute alarm to relative alarm (not supported).", item.icalString);
            }
            if (reminder >= 0) wbxml.atag("Reminder", reminder.toString());
            else TbSync.eventlog.add("info", syncdata, "Droping alarm after start date (not supported).", item.icalString);

        }

        //Sensitivity (CLASS)
        wbxml.atag("Sensitivity", eas.sync.mapThunderbirdPropertyToEas("CLASS", "Sensitivity", item));

        //Subject (obmitting these, should remove them from the server - that does not work reliably, so we send blanks)
        wbxml.atag("Subject", (item.title) ? item.title : "");

        //StartTime in UTC
        wbxml.atag("StartTime", item.startDate ? eas.tools.getIsoUtcString(item.startDate) : eas.tools.dateToBasicISOString(nowDate));

        //UID (limit to 300)
        //each TB event has an ID, which is used as EAS serverId - however there is a second UID in the ApplicationData
        //since we do not have two different IDs to use, we use the same ID
        if (!isException) { //docs say it would be allowed in exception in 2.5, but it does not work, if present
            wbxml.atag("UID", item.id);
        }
        //IMPORTANT in EAS v16 it is no longer allowed to send a UID
        //Only allowed in exceptions in v2.5


        //EAS MeetingStatus
        // 0 (000) The event is an appointment, which has no attendees.
        // 1 (001) The event is a meeting and the user is the meeting organizer.
        // 3 (011) This event is a meeting, and the user is not the meeting organizer; the meeting was received from someone else.
        // 5 (101) The meeting has been canceled and the user was the meeting organizer.
        // 7 (111) The meeting has been canceled. The user was not the meeting organizer; the meeting was received from someone else

        //there are 3 fields; Meeting, Owner, Cancelled
        //M can be reconstructed from #of attendees (looking at the old value is not wise, since it could have been changed)
        //C can be reconstucted from TB STATUS
        //O can be reconstructed by looking at the original value, or (if not present) by comparing EAS ownerID with TB ownerID

        let countAttendees = {};
        let attendees = item.getAttendees(countAttendees);
        //if (!(isException && asversion == "2.5")) { //MeetingStatus is not supported in exceptions in EAS 2.5        
        if (!isException) { //Exchange 2010 does not seem to support MeetingStatus at all in exceptions
            if (countAttendees == 0) wbxml.atag("MeetingStatus", "0");
            else {
                //get owner information
                let isReceived = false;
                if (item.hasProperty("X-EAS-MEETINGSTATUS")) isReceived = item.getProperty("X-EAS-MEETINGSTATUS") & 0x2;
                else isReceived = (item.organizer && item.organizer.id && cal.email.removeMailTo(item.organizer.id) != syncdata.accountData.getAccountProperty("user"));

                //either 1,3,5 or 7
                if (item.hasProperty("STATUS") && item.getProperty("STATUS") == "CANCELLED") {
                    //either 5 or 7
                    wbxml.atag("MeetingStatus", (isReceived ? "7" : "5"));
                } else {
                    //either 1 or 3
                    wbxml.atag("MeetingStatus", (isReceived ? "3" : "1"));
                }
            }
        }
        
        //Attendees
        let TB_responseType = null;        
        if (!(isException && asversion == "2.5")) { //attendees are not supported in exceptions in EAS 2.5
            if (countAttendees.value > 0) {
                wbxml.otag("Attendees");
                    for (let attendee of attendees) {
                        wbxml.otag("Attendee");
                            wbxml.atag("Email", cal.email.removeMailTo(attendee.id));
                            wbxml.atag("Name", (attendee.commonName ? attendee.commonName : cal.email.removeMailTo(attendee.id).split("@")[0]));
                            if (asversion != "2.5") {
                                //it's pointless to send AttendeeStatus, 
                                // - if we are the owner of a meeting, TB does not have an option to actually set the attendee status (on behalf of an attendee) in the UI
                                // - if we are an attendee (of an invite) we cannot and should not set status of other attendees and or own status must be send through a MeetingResponse
                                // -> all changes of attendee status are send from the server to us, either via ResponseType or via AttendeeStatus
                                //wbxml.atag("AttendeeStatus", eas.sync.MAP_TB2EAS.ATTENDEESTATUS[attendee.participationStatus]);

                                if (attendee.userType == "RESOURCE" || attendee.userType == "ROOM" || attendee.role == "NON-PARTICIPANT") wbxml.atag("AttendeeType","3");
                                else if (attendee.role == "REQ-PARTICIPANT" || attendee.role == "CHAIR") wbxml.atag("AttendeeType","1");
                                else wbxml.atag("AttendeeType","2"); //leftovers are optional
                            }
                        wbxml.ctag();
                    }
                wbxml.ctag();
            } else {
                wbxml.atag("Attendees");
            }
        }

        //Categories (see https://github.com/jobisoft/TbSync/pull/35#issuecomment-359286374)
        if (!isException) {
            wbxml.append(eas.sync.getItemCategories(item, syncdata));
        }

        //recurrent events (implemented by Chris Allan)
        if (!isException) {
            wbxml.append(eas.sync.getItemRecurrence(item, syncdata));
        }


        //---------------------------
        
        //TP PRIORITY (9=LOW, 5=NORMAL, 1=HIGH) not mapable to EAS Event
        //TODO: attachements (needs EAS 16.0!)
        
        //https://dxr.mozilla.org/comm-central/source/calendar/base/public/calIAlarm.idl
        //TbSync.dump("ALARM ("+i+")", [, alarms[i].related, alarms[i].repeat, alarms[i].repeatOffset, alarms[i].repeatDate, alarms[i].action].join("|"));

        return wbxml.getBytes();
    }
    
    
        /*
        //loop over all properties
        let propEnum = item.propertyEnumerator;
        while (propEnum.hasMoreElements()) {
            let prop = propEnum.getNext().QueryInterface(Components.interfaces.nsIProperty);
            let pname = prop.name;
            TbSync.dump("PROP", pname + " = " + prop.value);
        }
        */
    
}
