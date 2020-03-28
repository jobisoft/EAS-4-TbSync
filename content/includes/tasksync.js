/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

var Tasks = {

    // --------------------------------------------------------------------------- //
    // Read WBXML and set Thunderbird item
    // --------------------------------------------------------------------------- //
    setThunderbirdItemFromWbxml: function (tbItem, data, id, syncdata, mode = "standard") {

        let item = tbItem instanceof TbSync.lightning.TbItem ? tbItem.nativeItem : tbItem;
        
        let asversion = syncdata.accountData.getAccountProperty("asversion");
        item.id = id;
        eas.sync.setItemSubject(item, syncdata, data);
        if (TbSync.prefs.getIntPref("log.userdatalevel") > 2) TbSync.dump("Processing " + mode + " task item", item.title + " (" + id + ")");

        eas.sync.setItemBody(item, syncdata, data);
        eas.sync.setItemCategories(item, syncdata, data);
        eas.sync.setItemRecurrence(item, syncdata, data);

        let dueDate = null;
        if (data.DueDate && data.UtcDueDate) {
            //extract offset from EAS data
            let DueDate = new Date(data.DueDate);
            let UtcDueDate = new Date(data.UtcDueDate);
            let offset = (UtcDueDate.getTime() - DueDate.getTime())/60000;

            //timezone is identified by its offset
            let utc = cal.createDateTime(eas.tools.dateToBasicISOString(UtcDueDate)); //format "19800101T000000Z" - UTC
            dueDate = utc.getInTimezone(eas.tools.guessTimezoneByCurrentOffset(offset, utc));
            item.dueDate = dueDate;
        }

        if (data.StartDate && data.UtcStartDate) {
            //extract offset from EAS data
            let StartDate = new Date(data.StartDate);
            let UtcStartDate = new Date(data.UtcStartDate);
            let offset = (UtcStartDate.getTime() - StartDate.getTime())/60000;

            //timezone is identified by its offset
            let utc = cal.createDateTime(eas.tools.dateToBasicISOString(UtcStartDate)); //format "19800101T000000Z" - UTC
            item.entryDate = utc.getInTimezone(eas.tools.guessTimezoneByCurrentOffset(offset, utc));
        } else {
            //there is no start date? if this is a recurring item, we MUST add an entryDate, otherwise Thunderbird will not display the recurring items
            if (data.Recurrence) {
                if (dueDate) {
                    item.entryDate = dueDate; 
                    TbSync.eventlog.add("info", syncdata, "Copy task dueData to task startDate, because Thunderbird needs a startDate for recurring items.", item.icalString);
                } else {
                    TbSync.eventlog.add("info", syncdata, "Task without startDate and without dueDate but with recurrence info is not supported by Thunderbird. Recurrence will be lost.", item.icalString);
                }
            }
        }

        eas.sync.mapEasPropertyToThunderbird ("Sensitivity", "CLASS", data, item);
        eas.sync.mapEasPropertyToThunderbird ("Importance", "PRIORITY", data, item);

        item.clearAlarms();
        if (data.ReminderSet && data.ReminderTime && data.UtcStartDate) {        
            let UtcDate = eas.tools.createDateTime(data.UtcStartDate);
            let UtcAlarmDate = eas.tools.createDateTime(data.ReminderTime);
            let alarm = cal.createAlarm();
            alarm.related = Components.interfaces.calIAlarm.ALARM_RELATED_START; //TB saves new alarms as offsets, so we add them as such as well
            alarm.offset = UtcAlarmDate.subtractDate(UtcDate);
            alarm.action = "DISPLAY";
            item.addAlarm(alarm);
        }
        
        //status/percentage cannot be mapped
        if (data.Complete) {
          if (data.Complete == "0") {
            item.isCompleted = false;
          } else {
            item.isCompleted = true;
            if (data.DateCompleted) item.completedDate = eas.tools.createDateTime(data.DateCompleted);
          }
        }            
    },

/*
    Regenerate: After complete, the completed task is removed from the series and stored as an new entry. The series starts an week (as set) after complete date with one less occurence

    */







    // --------------------------------------------------------------------------- //
    //read TB event and return its data as WBXML
    // --------------------------------------------------------------------------- //
    getWbxmlFromThunderbirdItem: function (tbItem, syncdata) {
        let item = tbItem instanceof TbSync.lightning.TbItem ? tbItem.nativeItem : tbItem;

        let asversion = syncdata.accountData.getAccountProperty("asversion");
        let wbxml = eas.wbxmltools.createWBXML("", syncdata.type); //init wbxml with "" and not with precodes, and set initial codepage

        //Order of tags taken from: https://msdn.microsoft.com/en-us/library/dn338924(v=exchg.80).aspx
        
        //Subject
        wbxml.atag("Subject", (item.title) ? item.title : "");
        
        //Body
        wbxml.append(eas.sync.getItemBody(item, syncdata));

        //Importance
        wbxml.atag("Importance", eas.sync.mapThunderbirdPropertyToEas("PRIORITY", "Importance", item));

        //tasks is using extended ISO 8601 (2019-01-18T00:00:00.000Z)  instead of basic (20190118T000000Z), 
        //eas.tools.getIsoUtcString returns extended if true as second parameter is present
        
        // TB will enforce a StartDate if it has a recurrence
        let localStartDate = null;
        if (item.entryDate) {
            wbxml.atag("UtcStartDate", eas.tools.getIsoUtcString(item.entryDate, true));
            //to fake the local time as UTC, getIsoUtcString needs the third parameter to be true
            localStartDate = eas.tools.getIsoUtcString(item.entryDate, true, true);
            wbxml.atag("StartDate", localStartDate);
        }

        // Tasks without DueDate are breaking O365 - use StartDate as DueDate
        if (item.entryDate || item.dueDate) {
            wbxml.atag("UtcDueDate", eas.tools.getIsoUtcString(item.dueDate ? item.dueDate : item.entryDate, true));
            //to fake the local time as UTC, getIsoUtcString needs the third parameter to be true
            wbxml.atag("DueDate", eas.tools.getIsoUtcString(item.dueDate ? item.dueDate : item.entryDate, true, true));
        }
        
        //Categories
        wbxml.append(eas.sync.getItemCategories(item, syncdata));

        //Recurrence (only if localStartDate has been set)
        if (localStartDate) wbxml.append(eas.sync.getItemRecurrence(item, syncdata, localStartDate));
        
        //Complete
        if (item.isCompleted) {
                wbxml.atag("Complete", "1");
                wbxml.atag("DateCompleted", eas.tools.getIsoUtcString(item.completedDate, true));		
        } else {
                wbxml.atag("Complete", "0");
        }

        //Sensitivity
        wbxml.atag("Sensitivity", eas.sync.mapThunderbirdPropertyToEas("CLASS", "Sensitivity", item));

        //ReminderTime and ReminderSet
        let alarms = item.getAlarms({});
        if (alarms.length>0 && (item.entryDate || item.dueDate)) {
            let reminderTime;
            if (alarms[0].offset) {
                //create Date obj from entryDate by converting item.entryDate to an extended UTC ISO string, which can be parsed by Date
                //if entryDate is missing, the startDate of this object is set to its dueDate
                let UtcDate = new Date(eas.tools.getIsoUtcString(item.entryDate ? item.entryDate : item.dueDate, true));
                //add offset
                UtcDate.setSeconds(UtcDate.getSeconds() + alarms[0].offset.inSeconds);
                reminderTime = UtcDate.toISOString();
            } else {
                reminderTime = eas.tools.getIsoUtcString(alarms[0].alarmDate, true);
            }                
            wbxml.atag("ReminderTime", reminderTime);
            wbxml.atag("ReminderSet", "1");
        } else {
            wbxml.atag("ReminderSet", "0");
        }
        
        return wbxml.getBytes();
    },
}
