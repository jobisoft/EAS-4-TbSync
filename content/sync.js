/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

// - https://dxr.mozilla.org/comm-central/source/calendar/base/public/calIEvent.idl
// - https://dxr.mozilla.org/comm-central/source/calendar/base/public/calIItemBase.idl
// - https://dxr.mozilla.org/comm-central/source/calendar/base/public/calICalendar.idl
// - https://dxr.mozilla.org/comm-central/source/calendar/base/modules/calAsyncUtils.jsm

// https://msdn.microsoft.com/en-us/library/dd299454(v=exchg.80).aspx

var sync = {

    
        
    finish: function (aStatus = "", msg = "", details = "") {
        let status = tbSync.StatusData.SUCCESS
        switch (aStatus) {
            // custom status types
            case "resyncFolder":
                status = aStatus;
                break;
            
            case "":
            case "ok":
                status = tbSync.StatusData.SUCCESS;
                break;
            
            case "info":
                status = tbSync.StatusData.INFO;
                break;
            
            case "rerun":
                status = tbSync.StatusData.RERUN;
                break;
            
            case "warning":
                status = tbSync.StatusData.WARNING;
                break;
            
            case "error":
                status = tbSync.StatusData.ERROR;
                break;

            default:
                console.log("TbSync/EAS: Unknown status <"+aStatus+">");
                status = tbSync.StatusData.ERROR;
                break;
        }
        
        let e = new Error(); 
        e.name = "eas4tbsync";
        e.message = status.toUpperCase() + ": " + msg.toString() + " (" + details.toString() + ")";
        e.failed = (status != tbSync.StatusData.SUCCESS);
        e.statusData = new tbSync.StatusData(status, msg.toString(), details.toString());        
        return e; 
    }, 

    // update folders avail on server and handle added, removed and renamed
    // folders
    folderList: async function(syncData) {
        //should we recheck options/commands? Always check, if we have no info about asversion!
        if (syncData.accountData.getAccountProperty("asversion", "") == "" || (Date.now() - syncData.accountData.getAccountProperty("lastEasOptionsUpdate")) > 86400000 ) {
            await eas.network.getServerOptions(syncData);
        }
                        
        //only update the actual used asversion, if we are currently not connected or it has not yet been set
        if (syncData.accountData.getAccountProperty("asversion", "") == "" || !syncData.accountData.isConnected()) {
            //eval the currently in the UI selected EAS version
            let asversionselected = syncData.accountData.getAccountProperty("asversionselected");
            let allowedVersionsString = syncData.accountData.getAccountProperty("allowedEasVersions").trim();
            let allowedVersionsArray = allowedVersionsString.split(",");

            if (asversionselected == "auto") {
                if (allowedVersionsArray.includes("14.0")) syncData.accountData.setAccountProperty("asversion", "14.0");
                else if (allowedVersionsArray.includes("2.5")) syncData.accountData.setAccountProperty("asversion", "2.5");
                else if (allowedVersionsString == "") {
                    throw eas.sync.finish("error", "InvalidServerOptions");
                } else {
                    throw eas.sync.finish("error", "nosupportedeasversion::"+allowedVersionsArray.join(", "));
                }
            } else if (allowedVersionsString != "" && !allowedVersionsArray.includes(asversionselected)) {
                throw eas.sync.finish("error", "notsupportedeasversion::"+asversionselected+"::"+allowedVersionsArray.join(", "));
            } else {
                //just use the value set by the user
                syncData.accountData.setAccountProperty("asversion", asversionselected);
            }
        }
        
        //do we need to get a new policy key?
        if (syncData.accountData.getAccountProperty("provision") == "1" && syncData.accountData.getAccountProperty("policykey") == "0") {
            await eas.network.getPolicykey(syncData);
        }
        
        //set device info
        await eas.network.setDeviceInformation (syncData);

        syncData.setSyncState("prepare.request.folders"); 
        let foldersynckey = syncData.accountData.getAccountProperty("foldersynckey");

        //build WBXML to request foldersync
        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.switchpage("FolderHierarchy");
        wbxml.otag("FolderSync");
            wbxml.atag("SyncKey", foldersynckey);
        wbxml.ctag();

        syncData.setSyncState("send.request.folders"); 
        let response = await eas.network.sendRequest(wbxml.getBytes(), "FolderSync", syncData);

        syncData.setSyncState("eval.response.folders"); 
        let wbxmlData = eas.network.getDataFromResponse(response);
        eas.network.checkStatus(syncData, wbxmlData,"FolderSync.Status");

        let synckey = eas.xmltools.getWbxmlDataField(wbxmlData,"FolderSync.SyncKey");
        if (synckey) {
            syncData.accountData.setAccountProperty("foldersynckey", synckey);
        } else {
            throw eas.sync.finish("error", "wbxmlmissingfield::FolderSync.SyncKey");
        }
        
        //if we reach this point, wbxmlData contains FolderSync node, so the next "if" will not fail with an javascript error, 
        //no need to use save getWbxmlDataField function
        
        //are there any changes in folder hierarchy
        if (wbxmlData.FolderSync.Changes) {
            //looking for additions
            let add = eas.xmltools.nodeAsArray(wbxmlData.FolderSync.Changes.Add);
            for (let count = 0; count < add.length; count++) {
                //only add allowed folder types to DB (include trash(4), so we can find trashed folders
                if (!["9","14","8","13","7","15", "4"].includes(add[count].Type))
                    continue;

                let existingFolder = syncData.accountData.getFolder("serverID", add[count].ServerId);
                if (existingFolder) {
                    //server has send us an ADD for a folder we alreay have, treat as update
                    existingFolder.setFolderProperty("foldername", add[count].DisplayName);
                    existingFolder.setFolderProperty("type", add[count].Type);
                    existingFolder.setFolderProperty("parentID", add[count].ParentId);
                } else {
                    //create folder obj for new  folder settings
                    let newFolder = syncData.accountData.createNewFolder();
                    switch (add[count].Type) {
                        case "9": //contact
                        case "14": 
                            newFolder.setFolderProperty("targetType", "addressbook");
                            break;
                        case "8": //event
                        case "13":
                            newFolder.setFolderProperty("targetType", "calendar");
                            break;
                        case "7": //todo
                        case "15":
                            newFolder.setFolderProperty("targetType", "calendar");
                            break;
                        default:
                            newFolder.setFolderProperty("targetType", "none");
                            break;
                        
                    }
                    
                    newFolder.setFolderProperty("serverID", add[count].ServerId);
                    newFolder.setFolderProperty("foldername", add[count].DisplayName);
                    newFolder.setFolderProperty("type", add[count].Type);
                    newFolder.setFolderProperty("parentID", add[count].ParentId);

                    //do we have a cached folder?
                    let cachedFolderData = syncData.accountData.getFolderFromCache("serverID",  add[count].ServerId);
                    if (cachedFolderData) {
                        // copy fields from cache which we want to re-use
                        newFolder.setFolderProperty("targetColor", cachedFolderData.getFolderProperty("targetColor"));
                        newFolder.setFolderProperty("targetName", cachedFolderData.getFolderProperty("targetName"));
                        newFolder.setFolderProperty("downloadonly", cachedFolderData.getFolderProperty("downloadonly"));
                    }
                }
            }
            
            //looking for updates
            let update = eas.xmltools.nodeAsArray(wbxmlData.FolderSync.Changes.Update);
            for (let count = 0; count < update.length; count++) {
                let existingFolder = syncData.accountData.getFolder("serverID", update[count].ServerId);
                if (existingFolder) {
                    //update folder
                    existingFolder.setFolderProperty("foldername", update[count].DisplayName);
                    existingFolder.setFolderProperty("type", update[count].Type);
                    existingFolder.setFolderProperty("parentID", update[count].ParentId);
                }
            }

            //looking for deletes
            let del = eas.xmltools.nodeAsArray(wbxmlData.FolderSync.Changes.Delete);
            for (let count = 0; count < del.length; count++) {
                let existingFolder = syncData.accountData.getFolder("serverID", del[count].ServerId);
                if (existingFolder) {
                    existingFolder.targetData.decoupleTarget("[deleted on server]", /* move folder into cache */ true);
                }
            }
        }
    },
    




    deleteFolder: async function (syncData)  {
        if (!syncData.currentFolderData) {
            return;
        }
        
        if (!syncData.accountData.getAccountProperty("allowedEasCommands").split(",").includes("FolderDelete")) {
            throw eas.sync.finish("error", "notsupported::FolderDelete");
        }

        syncData.setSyncState("prepare.request.deletefolder");
        let foldersynckey = syncData.accountData.getAccountProperty("foldersynckey");

        //request foldersync
        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.switchpage("FolderHierarchy");
        wbxml.otag("FolderDelete");
            wbxml.atag("SyncKey", foldersynckey);
            wbxml.atag("ServerId", syncData.currentFolderData.getFolderProperty("serverID"));
        wbxml.ctag();

        syncData.setSyncState("send.request.deletefolder");
        let response = await eas.network.sendRequest(wbxml.getBytes(), "FolderDelete", syncData);

        syncData.setSyncState("eval.response.deletefolder");
        let wbxmlData = eas.network.getDataFromResponse(response);

        eas.network.checkStatus(syncData, wbxmlData,"FolderDelete.Status");

        let synckey = eas.xmltools.getWbxmlDataField(wbxmlData,"FolderDelete.SyncKey");
        if (synckey) {
            syncData.accountData.setAccountProperty("foldersynckey", synckey);
            syncData.currentFolderData.remove();
        } else {
            throw eas.sync.finish("error", "wbxmlmissingfield::FolderDelete.SyncKey");
        }
    },





    singleFolder: async function (syncData)  {
        let folderReSyncs = 0;
        
        do {                
            let rerun = false;
        
            folderReSyncs++;
            if (folderReSyncs > 2) {
                throw eas.sync.finish("warning", "resync-loop");
            }

            // add target to syncData (getTarget() will throw "nolightning" if lightning missing)
            try {
                // accessing the target for the first time will check if it is avail and if not will create it (if possible)
                syncData.target = syncData.currentFolderData.targetData.getTarget();
            } catch (e) {
                throw eas.sync.finish("warning", e.message);
            }

            //get syncData type, which is also used in WBXML for the CLASS element
            syncData.type = null;
            switch (syncData.currentFolderData.getFolderProperty("type")) {
                case "9": //contact
                case "14": 
                    syncData.type = "Contacts";
                    break;
                case "8": //event
                case "13":
                    syncData.type = "Calendar";
                    break;
                case "7": //todo
                case "15":
                    syncData.type = "Tasks";
                    break;
                default:
                     throw eas.sync.finish("info", "skipped");
                    break;
            }

            syncData.setSyncState("preparing");
            
            //get synckey if needed
            syncData.synckey = syncData.currentFolderData.getFolderProperty("synckey");                
            if (syncData.synckey == "") {
                await eas.network.getSynckey(syncData);
            }
            
            //sync folder
            syncData.timeOfLastSync = syncData.currentFolderData.getFolderProperty( "lastsynctime") / 1000;
            syncData.timeOfThisSync = (Date.now() / 1000) - 1;
            
            let lightningBatch = false;
            let lightningReadOnly = "";
            let error = null;
            
            try {
                switch (syncData.type) {
                    case "Contacts": 
                        await eas.sync.easFolder(syncData);
                        break;

                    case "Calendar":
                    case "Tasks":                            
                        //save current value of readOnly (or take it from the setting)
                        lightningReadOnly = syncData.target.calendar.getProperty("readOnly") || (syncData.currentFolderData.getFolderProperty( "downloadonly") == "1");                       
                        syncData.target.calendar.setProperty("readOnly", false);
                        
                        lightningBatch = true;
                        syncData.target.calendar.startBatch();

                        await eas.sync.easFolder(syncData);
                        break;
                }
            } catch (report) {
                //Filter out custom status types, which must be handled here and may not be passed on (because unknown to the outside)
                if (report.name == "eas4tbsync" && report.statusData.type == "resyncFolder") {
                    rerun = true;
                }  else {
                    error = report;
                }
            }
            
            if (lightningBatch) {
                syncData.target.calendar.endBatch();
                syncData.target.calendar.setProperty("readOnly", lightningReadOnly);
            }
            
            if (error) throw error;
            if (!rerun) break;
        }
        while (true);
    },










    // ---------------------------------------------------------------------------
    // MAIN FUNCTIONS TO SYNC AN EAS FOLDER
    // ---------------------------------------------------------------------------

    easFolder: async function (syncData)  {
        syncData.progressData.reset();

        if (syncData.currentFolderData.getFolderProperty("downloadonly") == "1") {		
            await eas.sync.revertLocalChanges(syncData);
        }

        await eas.network.getItemEstimate (syncData);
        await eas.sync.requestRemoteChanges (syncData); 

        if (syncData.currentFolderData.getFolderProperty("downloadonly") != "1") {		
            await eas.sync.sendLocalChanges(syncData);
        }
    },
    

    requestRemoteChanges: async function (syncData)  {
        do {
            syncData.setSyncState("prepare.request.remotechanges");
            syncData.request = "";
            syncData.response = "";
        
            // BUILD WBXML
            let wbxml = eas.wbxmltools.createWBXML();
            wbxml.otag("Sync");
                wbxml.otag("Collections");
                    wbxml.otag("Collection");
                        if (syncData.accountData.getAccountProperty("asversion") == "2.5") wbxml.atag("Class", syncData.type);
                        wbxml.atag("SyncKey", syncData.synckey);
                        wbxml.atag("CollectionId", syncData.currentFolderData.getFolderProperty("serverID"));
                        wbxml.atag("DeletesAsMoves");
                        wbxml.atag("GetChanges");
                        wbxml.atag("WindowSize",  eas.prefs.getIntPref("maxitems").toString());

                        if (syncData.accountData.getAccountProperty("asversion") != "2.5") {
                            wbxml.otag("Options");
                                if (syncData.type == "Calendar") wbxml.atag("FilterType", eas.tools.getFilterType());
                                wbxml.atag("Class", syncData.type);
                                wbxml.switchpage("AirSyncBase");
                                wbxml.otag("BodyPreference");
                                    wbxml.atag("Type", "1");
                                wbxml.ctag();
                                wbxml.switchpage("AirSync");
                            wbxml.ctag();
                        } else if (syncData.type == "Calendar") { //in 2.5 we only send it to filter Calendar
                            wbxml.otag("Options");
                                 wbxml.atag("FilterType", eas.tools.getFilterType());
                            wbxml.ctag();
                        }

                    wbxml.ctag();
                wbxml.ctag();
            wbxml.ctag();

            //SEND REQUEST
            syncData.setSyncState("send.request.remotechanges");
            let response = await eas.network.sendRequest(wbxml.getBytes(), "Sync", syncData);

            //VALIDATE RESPONSE
            // get data from wbxml response, some servers send empty response if there are no changes, which is not an error
            let wbxmlData = eas.network.getDataFromResponse(response, eas.flags.allowEmptyResponse);
            if (wbxmlData === null) return;
        
            //check status, throw on error
            eas.network.checkStatus(syncData, wbxmlData,"Sync.Collections.Collection.Status");

            //PROCESS COMMANDS        
            await eas.sync.processCommands(wbxmlData, syncData);

            //Update count in UI
            syncData.setSyncState("eval.response.remotechanges");

            //update synckey
            eas.network.updateSynckey(syncData, wbxmlData);
            
            if (!eas.xmltools.hasWbxmlDataField(wbxmlData,"Sync.Collections.Collection.MoreAvailable")) {
                //Feedback from users: They want to see the final count
                await tbSync.tools.sleep(100, false);
                return;
            }
        } while (true);
                
    },


    sendLocalChanges: async function (syncData)  {        
        let maxnumbertosend = eas.prefs.getIntPref("maxitems");
        syncData.progressData.reset(0, syncData.target.getItemsFromChangeLog().length);

        //keep track of failed items
        syncData.failedItems = [];
        
        let done = false;
        let numberOfItemsToSend = maxnumbertosend;
        do {
            syncData.setSyncState("prepare.request.localchanges");
            syncData.request = "";
            syncData.response = "";

            //get changed items from ChangeLog
            let changes = syncData.target.getItemsFromChangeLog(numberOfItemsToSend);
            let c=0;
            let e=0;

            //keep track of send items during this request
            let changedItems = [];
            let addedItems = {};
            let sendItems = [];
                
            // BUILD WBXML
            let wbxml = eas.wbxmltools.createWBXML();
            wbxml.otag("Sync");
                wbxml.otag("Collections");
                    wbxml.otag("Collection");
                        if (syncData.accountData.getAccountProperty("asversion") == "2.5") wbxml.atag("Class", syncData.type);
                        wbxml.atag("SyncKey", syncData.synckey);
                        wbxml.atag("CollectionId", syncData.currentFolderData.getFolderProperty("serverID"));
                        wbxml.otag("Commands");

                            for (let i=0; i<changes.length; i++) if (!syncData.failedItems.includes(changes[i].id)) {
                                //tbSync.dump("CHANGES",(i+1) + "/" + changes.length + " ("+changes[i].status+"," + changes[i].id + ")");
                                let item = null;
                                switch (changes[i].status) {

                                    case "added_by_user":
                                        item = await syncData.target.getItem(changes[i].id);
                                        if (item) {
                                            //filter out bad object types for this folder
                                            if (syncData.type == eas.sync.getEasItemType(item)) {
                                                //create a temp clientId, to cope with too long or invalid clientIds (for EAS)
                                                let clientId = Date.now() + "-" + c;
                                                addedItems[clientId] = changes[i].id;
                                                sendItems.push({type: changes[i].status, id: changes[i].id});
                                                
                                                wbxml.otag("Add");
                                                wbxml.atag("ClientId", clientId); //Our temp clientId will get replaced by an id generated by the server
                                                    wbxml.otag("ApplicationData");
                                                        wbxml.switchpage(syncData.type);

/*wbxml.atag("TimeZone", "xP///0UAdQByAG8AcABlAC8AQgBlAHIAbABpAG4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAoAAAAFAAIAAAAAAAAAAAAAAEUAdQByAG8AcABlAC8AQgBlAHIAbABpAG4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAAFAAEAAAAAAAAAxP///w==");
wbxml.atag("AllDayEvent", "0");
wbxml.switchpage("AirSyncBase");
wbxml.otag("Body");
    wbxml.atag("Type", "1");
    wbxml.atag("EstimatedDataSize", "0");
    wbxml.atag("Data");
wbxml.ctag();

wbxml.switchpage(syncData.type);						
wbxml.atag("BusyStatus", "2");
wbxml.atag("OrganizerName", "REDACTED.REDACTED");
wbxml.atag("OrganizerEmail", "REDACTED.REDACTED@REDACTED");
wbxml.atag("DtStamp", "20190131T091024Z");
wbxml.atag("EndTime", "20180906T083000Z");
wbxml.atag("Location");
wbxml.atag("Reminder", "5");
wbxml.atag("Sensitivity", "0");
wbxml.atag("Subject", "SE-CN weekly sync");
wbxml.atag("StartTime", "20180906T080000Z");
wbxml.atag("UID", "1D51E503-9DFE-4A46-A6C2-9129E5E00C1D");
wbxml.atag("MeetingStatus", "3");
wbxml.otag("Attendees");
    wbxml.otag("Attendee");
        wbxml.atag("Email", "REDACTED.REDACTED@REDACTED");
        wbxml.atag("Name", "REDACTED.REDACTED");
        wbxml.atag("AttendeeType", "1");
    wbxml.ctag();
wbxml.ctag();
wbxml.atag("Categories");
wbxml.otag("Recurrence");
    wbxml.atag("Type", "1");
    wbxml.atag("DayOfWeek", "16");
    wbxml.atag("Interval", "1");
wbxml.ctag();
wbxml.otag("Exceptions");
    wbxml.otag("Exception");
        wbxml.atag("ExceptionStartTime", "20181227T090000Z");
        wbxml.atag("Deleted", "1");
    wbxml.ctag();
wbxml.ctag();*/

                                                        wbxml.append(eas.sync.getWbxmlFromThunderbirdItem(item, syncData));
                                                        wbxml.switchpage("AirSync");
                                                    wbxml.ctag();
                                                wbxml.ctag();
                                                c++;
                                            } else {
                                                eas.sync.updateFailedItems(syncData, "forbidden" + eas.sync.getEasItemType(item) +"ItemIn" + syncData.type + "Folder", item.primaryKey, item.toString());
                                                e++;
                                            }
                                        } else {
                                            syncData.target.removeItemFromChangeLog(changes[i].id);
                                        }
                                        break;
                                    
                                    case "modified_by_user":
                                        item = await syncData.target.getItem(changes[i].id);
                                        if (item) {
                                            //filter out bad object types for this folder
                                            if (syncData.type == eas.sync.getEasItemType(item)) {
                                                wbxml.otag("Change");
                                                wbxml.atag("ServerId", changes[i].id);
                                                    wbxml.otag("ApplicationData");
                                                        wbxml.switchpage(syncData.type);
                                                        wbxml.append(eas.sync.getWbxmlFromThunderbirdItem(item, syncData));
                                                        wbxml.switchpage("AirSync");
                                                    wbxml.ctag();
                                                wbxml.ctag();
                                                changedItems.push(changes[i].id);
                                                sendItems.push({type: changes[i].status, id: changes[i].id});
                                                c++;
                                            } else {
                                                eas.sync.updateFailedItems(syncData, "forbidden" + eas.sync.getEasItemType(item) +"ItemIn" + syncData.type + "Folder", item.primaryKey, item.toString());
                                                e++;
                                            }
                                        } else {
                                            syncData.target.removeItemFromChangeLog(changes[i].id);
                                        }
                                        break;
                                    
                                    case "deleted_by_user":
                                        wbxml.otag("Delete");
                                        wbxml.atag("ServerId", changes[i].id);
                                        wbxml.ctag();
                                        changedItems.push(changes[i].id);
                                        sendItems.push({type: changes[i].status, id: changes[i].id});
                                        c++;
                                        break;
                                }
                            }

                        wbxml.ctag(); //Commands
                    wbxml.ctag(); //Collection
                wbxml.ctag(); //Collections
            wbxml.ctag(); //Sync


            if (c > 0) { //if there was at least one actual local change, send request

                //SEND REQUEST & VALIDATE RESPONSE
                syncData.setSyncState("send.request.localchanges");
                let response = await eas.network.sendRequest(wbxml.getBytes(), "Sync", syncData);
                
                syncData.setSyncState("eval.response.localchanges");

                //get data from wbxml response
                let wbxmlData = eas.network.getDataFromResponse(response);
            
                //check status and manually handle error states which support softfails
                let errorcause = eas.network.checkStatus(syncData, wbxmlData, "Sync.Collections.Collection.Status", "", true);
                switch (errorcause) {
                    case "":
                        break;
                    
                    case "Sync.4": //Malformed request
                    case "Sync.6": //Invalid item
                        //some servers send a global error - to catch this, we reduce the number of items we send to the server
                        if (sendItems.length == 1) {
                            //the request contained only one item, so we know which one failed
                            if (sendItems[0].type == "deleted_by_user") {
                                //we failed to delete an item, discard and place message in log
                                syncData.target.removeItemFromChangeLog(sendItems[0].id);
                                tbSync.errorlog.add("warning", syncData.errorInfo, "ErrorOnDelete::"+sendItems[0].id);
                            } else {
                                let foundItem = await syncData.target.getItem(sendItems[0].id);                    
                                if (foundItem) {
                                    eas.sync.updateFailedItems(syncData, errorcause, foundItem.primaryKey, foundItem.toString());
                                } else {
                                    //should not happen
                                    syncData.target.removeItemFromChangeLog(sendItems[0].id);                                    
                                }
                            }
                            syncData.progressData.inc();
                            //restore numberOfItemsToSend
                            numberOfItemsToSend = maxnumbertosend;                            
                        } else if (sendItems.length > 1) {
                            //reduce further
                            numberOfItemsToSend = Math.min(1, Math.round(sendItems.length / 5));
                        } else {
                            //sendItems.length == 0 ??? recheck but this time let it handle all cases
                            eas.network.checkStatus(syncData, wbxmlData, "Sync.Collections.Collection.Status");
                        }
                        break;

                    default:
                        //recheck but this time let it handle all cases
                        eas.network.checkStatus(syncData, wbxmlData, "Sync.Collections.Collection.Status");
                }        
                
                await tbSync.tools.sleep(10);

                if (errorcause == "") {
                    //PROCESS RESPONSE        
                    await eas.sync.processResponses(wbxmlData, syncData, addedItems, changedItems);
                
                    //PROCESS COMMANDS        
                    await eas.sync.processCommands(wbxmlData, syncData);

                    //remove all items from changelog that did not fail
                    for (let a=0; a < changedItems.length; a++) {
                        syncData.target.removeItemFromChangeLog(changedItems[a]);
                        syncData.progressData.inc();
                    }

                    //update synckey
                    eas.network.updateSynckey(syncData, wbxmlData);
                }
            
            } else if (e==0) { //if there was no local change and also no error (which will not happen twice) finish

                done = true;

            }
        
        } while (!done);
        
        //was there an error?
        if (syncData.failedItems.length > 0) {
            throw eas.finish("warning", "ServerRejectedSomeItems::" + syncData.failedItems.length);                            
        }
        
    },




    revertLocalChanges: async function (syncData)  {       
        let maxnumbertosend = eas.prefs.getIntPref("maxitems");
        syncData.progressData.reset(0, syncData.target.getItemsFromChangeLog().length);
        if (syncData.progressData.todo == 0) {
            return;
        }

        let viaItemOperations = (syncData.accountData.getAccountProperty("allowedEasCommands").split(",").includes("ItemOperations"));

        //get changed items from ChangeLog
        do {
            syncData.setSyncState("prepare.request.revertlocalchanges");
            let changes = syncData.target.getItemsFromChangeLog(maxnumbertosend);
            let c=0;
            syncData.request = "";
            syncData.response = "";
            
            // BUILD WBXML
            let wbxml = eas.wbxmltools.createWBXML();
            if (viaItemOperations) {
                wbxml.switchpage("ItemOperations");
                wbxml.otag("ItemOperations");
            } else {
                wbxml.otag("Sync");
                    wbxml.otag("Collections");
                        wbxml.otag("Collection");
                            if (syncData.accountData.getAccountProperty("asversion") == "2.5") wbxml.atag("Class", syncData.type);
                            wbxml.atag("SyncKey", syncData.synckey);
                            wbxml.atag("CollectionId", syncData.currentFolderData.getFolderProperty("serverID"));
                            wbxml.otag("Commands");
            }

            for (let i=0; i<changes.length; i++) {
                let item = null;
                let ServerId = changes[i].id;
                let foundItem = await syncData.target.getItem(ServerId);

                switch (changes[i].status) {
                    case "added_by_user": //remove
                        if (foundItem) {
                            await syncData.target.deleteItem(foundItem);
                        }
                    break;
                    
                    case "modified_by_user":
                        if (foundItem) { //delete item so it can be replaced with a fresh copy, the changelog entry will be changed from modified to deleted
                            await syncData.target.deleteItem(foundItem);
                        }
                    case "deleted_by_user":
                        if (viaItemOperations) {
                            wbxml.otag("Fetch");
                                wbxml.atag("Store", "Mailbox");
                                wbxml.switchpage("AirSync");
                                wbxml.atag("CollectionId", syncData.currentFolderData.getFolderProperty("serverID"));
                                wbxml.atag("ServerId", ServerId);
                                wbxml.switchpage("ItemOperations");
                                wbxml.otag("Options");
                                    wbxml.switchpage("AirSyncBase");
                                    wbxml.otag("BodyPreference");
                                        wbxml.atag("Type","1");
                                    wbxml.ctag();
                                    wbxml.switchpage("ItemOperations");
                                wbxml.ctag();
                            wbxml.ctag();
                        } else {
                            wbxml.otag("Fetch");
                                wbxml.atag("ServerId", ServerId);
                            wbxml.ctag();
                        }
                        c++;
                        break;
                }
            }

            if (viaItemOperations) {
                wbxml.ctag(); //ItemOperations
            } else {
                            wbxml.ctag(); //Commands
                        wbxml.ctag(); //Collection
                    wbxml.ctag(); //Collections
                wbxml.ctag(); //Sync
            }

            if (c > 0) { //if there was at least one actual local change, send request
                let error = false;
                let wbxmlData = "";
                
                //SEND REQUEST & VALIDATE RESPONSE
                try {
                    syncData.setSyncState("send.request.revertlocalchanges");
                    let response = await eas.network.sendRequest(wbxml.getBytes(), (viaItemOperations) ? "ItemOperations" : "Sync", syncData);
                        
                    syncData.setSyncState("eval.response.revertlocalchanges");

                    //get data from wbxml response
                    wbxmlData = eas.network.getDataFromResponse(response);
                } catch (e) {
                    //we do not handle errors, IF there was an error, wbxmlData is empty and will trigger the fallback
                }
                
                let fetchPath = (viaItemOperations) ? "ItemOperations.Response.Fetch" : "Sync.Collections.Collection.Responses.Fetch";
                if (eas.xmltools.hasWbxmlDataField(wbxmlData, fetchPath)) {
                
                    //looking for additions
                    let add = eas.xmltools.nodeAsArray(eas.xmltools.getWbxmlDataField(wbxmlData, fetchPath));
                    for (let count = 0; count < add.length; count++) {
                        await tbSync.tools.sleep(2);

                        let ServerId = add[count].ServerId;
                        let data = (viaItemOperations) ? add[count].Properties : add[count].ApplicationData;
                        
                        if (data && ServerId) {
                            let foundItem = await syncData.target.getItem(ServerId);
                            if (!foundItem) { //do NOT add, if an item with that ServerId was found
                                    let newItem = eas.sync.createItem(syncData);
                                    try {
                                        eas.sync[syncData.type].setThunderbirdItemFromWbxml(newItem, data, ServerId, syncData);
                                        await syncData.target.addItem(newItem);
                                    } catch (e) {
                                        eas.xmltools.printXmlData(add[count], true); //include application data in log                  
                                        tbSync.errorlog.add("warning", syncData.errorInfo, "BadItemSkipped::JavaScriptError", newItem.toString());
                                        throw e; // unable to add item to Thunderbird - fatal error
                                    }
                            } else {
                                //should not happen, since we deleted that item beforehand
                                syncData.target.removeItemFromChangeLog(ServerId);
                            }
                            syncData.progressData.inc();
                        } else {
                            error = true;
                            break;
                        }
                    }
                } else {
                    error = true;
                }
                
                if (error) {
                    //if ItemOperations.Fetch fails, fall back to Sync.Fetch, if that fails, fall back to resync
                    if (viaItemOperations) {
                        viaItemOperations = false;
                        tbSync.errorlog.add("info", syncData.errorInfo, "Server returned error during ItemOperations.Fetch, falling back to Sync.Fetch.");
                    } else {
                        await eas.sync.revertLocalChangesViaResync(syncData);
                        return;
                    }
                }
                            
            } else { //if there was no more local change we need to revert, return

                return;

            }
        
        } while (true);
        
    },

    revertLocalChangesViaResync: async function (syncData) {
        tbSync.errorlog.add("info", syncData.errorInfo, "Server does not support ItemOperations.Fetch and/or Sync.Fetch, must revert via resync.");
        let changes = syncData.target.getItemsFromChangeLog();

        syncData.progressData.reset(0, changes.length);
        syncData.setSyncState("prepare.request.revertlocalchanges");
        
        //remove all changes, so we can get them fresh from the server
        for (let i=0; i<changes.length; i++) {
            let item = null;
            let ServerId = changes[i].id;
            syncData.target.removeItemFromChangeLog(ServerId);
            let foundItem = await syncData.target.getItem(ServerId);
            if (foundItem) { //delete item with that ServerId
                await syncData.target.deleteItem(foundItem);
            }
            syncData.progressData.inc();
        }
        
        //This will resync all missing items fresh from the server
        tbSync.errorlog.add("info", syncData.errorInfo, "RevertViaFolderResync");
        eas.onResetTarget(syncData.currentFolderData);
        throw tbSync.eas.finish("resyncFolder", "RevertViaFolderResync"); 
    },




    // ---------------------------------------------------------------------------
    // SUB FUNCTIONS CALLED BY  MAIN FUNCTION
    // ---------------------------------------------------------------------------
    
    processCommands:  async function (wbxmlData, syncData)  {
        //any commands for us to work on? If we reach this point, Sync.Collections.Collection is valid, 
        //no need to use the save getWbxmlDataField function
        if (wbxmlData.Sync.Collections.Collection.Commands) {
        
            //looking for additions
            let add = eas.xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Commands.Add);
            for (let count = 0; count < add.length; count++) {
                await tbSync.tools.sleep(2);

                let ServerId = add[count].ServerId;
                let data = add[count].ApplicationData;

                let foundItem = await syncData.target.getItem(ServerId);
                if (!foundItem) {
                    //do NOT add, if an item with that ServerId was found
                    let newItem = eas.sync.createItem(syncData);
                    try {
                        eas.sync[syncData.type].setThunderbirdItemFromWbxml(newItem, data, ServerId, syncData);
                        await syncData.target.addItem(newItem);
                    } catch (e) {
                        eas.xmltools.printXmlData(add[count], true); //include application data in log                  
                        tbSync.errorlog.add("warning", syncData.errorInfo, "BadItemSkipped::JavaScriptError", newItem.toString());
                        throw e; // unable to add item to Thunderbird - fatal error
                    }
                } else {
                    tbSync.errorlog.add("info", syncData.errorInfo, "Add request, but element exists already, skipped.", ServerId);
                }
                syncData.progressData.inc();
            }

            //looking for changes
            let upd = eas.xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Commands.Change);
            //inject custom change object for debug
            //upd = JSON.parse('[{"ServerId":"2tjoanTeS0CJ3QTsq5vdNQAAAAABDdrY6Gp03ktAid0E7Kub3TUAAAoZy4A1","ApplicationData":{"DtStamp":"20171109T142149Z"}}]');
            for (let count = 0; count < upd.length; count++) {
                await tbSync.tools.sleep(2);

                let ServerId = upd[count].ServerId;
                let data = upd[count].ApplicationData;

                syncData.progressData.inc();
                let foundItem = await syncData.target.getItem(ServerId);
                if (foundItem) { //only update, if an item with that ServerId was found
                                        
                    let keys = Object.keys(data);
                    //replace by smart merge
                    if (keys.length == 1 && keys[0] == "DtStamp") tbSync.dump("DtStampOnly", keys); //ignore DtStamp updates (fix with smart merge)
                    else {
                        
                        if (foundItem.changelogStatus !== null) {
                            tbSync.errorlog.add("info", syncData.errorInfo, "Change request from server, but also local modifications, server wins!", ServerId);
                            foundItem.changelogStatus = null;
                        }
                        
                        let newItem = foundItem.clone();
                        try {
                            eas.sync[syncData.type].setThunderbirdItemFromWbxml(newItem, data, ServerId, syncData);
                            await syncData.target.modifyItem(newItem, foundItem);
                        } catch (e) {
                            tbSync.errorlog.add("warning", syncData.errorInfo, "BadItemSkipped::JavaScriptError", newItem.toString());
                            eas.xmltools.printXmlData(upd[count], true);  //include application data in log                   
                            throw e; // unable to mod item to Thunderbird - fatal error
                        }
                    }
                    
                }
            }
            
            //looking for deletes
            let del = eas.xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Commands.Delete).concat(eas.xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Commands.SoftDelete));
            for (let count = 0; count < del.length; count++) {
                await tbSync.tools.sleep(2);

                let ServerId = del[count].ServerId;

                let foundItem = await syncData.target.getItem(ServerId);
                if (foundItem) { //delete item with that ServerId
                    await syncData.target.deleteItem(foundItem);
                }
                syncData.progressData.inc();
            }
        
        }
    },


    updateFailedItems: function (syncData, cause, id, data) {                
        //something is wrong with this item, move it to the end of changelog and go on
        if (!syncData.failedItems.includes(id)) {
            //the extra parameter true will re-add the item to the end of the changelog
            syncData.target.removeItemFromChangeLog(id, true);                        
            syncData.failedItems.push(id);            
            tbSync.errorlog.add("info", syncData.errorInfo, "BadItemSkipped::" + tbSync.getString("status." + cause ,"eas"), "\n\nRequest:\n" + syncData.request + "\n\nResponse:\n" + syncData.response + "\n\nElement:\n" + data);
        }
    },


    processResponses: async function (wbxmlData, syncData, addedItems, changedItems)  {
            //any responses for us to work on?  If we reach this point, Sync.Collections.Collection is valid, 
            //no need to use the save getWbxmlDataField function
            if (wbxmlData.Sync.Collections.Collection.Responses) {

                //looking for additions (Add node contains, status, old ClientId and new ServerId)
                let add = eas.xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Responses.Add);
                for (let count = 0; count < add.length; count++) {
                    await tbSync.tools.sleep(2);

                    //get the true Thunderbird UID of this added item (we created a temp clientId during add)
                    add[count].ClientId = addedItems[add[count].ClientId];

                    //look for an item identfied by ClientId and update its id to the new id received from the server
                    let foundItem = await syncData.target.getItem(add[count].ClientId);                    
                    if (foundItem) {

                        //Check status, stop sync if bad, allow soft fail
                        let errorcause = eas.network.checkStatus(syncData, add[count],"Status","Sync.Collections.Collection.Responses.Add["+count+"].Status", true);
                        if (errorcause !== "") {
                            //something is wrong with this item, move it to the end of changelog and go on
                            eas.sync.updateFailedItems(syncData, errorcause, foundItem.primaryKey, foundItem.toString());
                        } else {
                            let newItem = foundItem.clone();
                            newItem.id = add[count].ServerId;
                            syncData.target.removeItemFromChangeLog(add[count].ClientId);
                            await syncData.target.modifyItem(newItem, foundItem);
                            syncData.progressData.inc();
                        }

                    }
                }

                //looking for modifications 
                let upd = eas.xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Responses.Change);
                for (let count = 0; count < upd.length; count++) {
                    let foundItem = await syncData.target.getItem(upd[count].ServerId);                    
                    if (foundItem) {

                        //Check status, stop sync if bad, allow soft fail
                        let errorcause = eas.network.checkStatus(syncData, upd[count],"Status","Sync.Collections.Collection.Responses.Change["+count+"].Status", true);
                        if (errorcause !== "") {
                            //something is wrong with this item, move it to the end of changelog and go on
                            eas.sync.updateFailedItems(syncData, errorcause, foundItem.primaryKey, foundItem.toString());
                            //also remove from changedItems
                            let p = changedItems.indexOf(upd[count].ServerId);
                            if (p>-1) changedItems.splice(p,1);
                        }

                    }
                }

                //looking for deletions 
                let del = eas.xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Responses.Delete);
                for (let count = 0; count < del.length; count++) {
                    //What can we do about failed deletes? SyncLog
                    eas.network.checkStatus(syncData, del[count],"Status","Sync.Collections.Collection.Responses.Delete["+count+"].Status", true);
                }
                
            }
    },










    // ---------------------------------------------------------------------------
    // HELPER FUNCTIONS AND DEFINITIONS
    // ---------------------------------------------------------------------------
        
    MAP_EAS2TB : {
        //EAS Importance: 0 = LOW | 1 = NORMAL | 2 = HIGH
        Importance : { "0":"9", "1":"5", "2":"1"}, //to PRIORITY
        //EAS Sensitivity :  0 = Normal  |  1 = Personal  |  2 = Private  |  3 = Confidential
        Sensitivity : { "0":"PUBLIC", "1":"unset", "2":"PRIVATE", "3":"CONFIDENTIAL"}, //to CLASS
        //EAS BusyStatus:  0 = Free  |  1 = Tentative  |  2 = Busy  |  3 = Work  |  4 = Elsewhere
        BusyStatus : {"0":"TRANSPARENT", "1":"unset", "2":"OPAQUE", "3":"OPAQUE", "4":"OPAQUE"}, //to TRANSP
        //EAS AttendeeStatus: 0 =Response unknown (but needed) |  2 = Tentative  |  3 = Accept  |  4 = Decline  |  5 = Not responded (and not needed) || 1 = Organizer in ResponseType
        ATTENDEESTATUS : {"0": "NEEDS-ACTION", "1":"Orga", "2":"TENTATIVE", "3":"ACCEPTED", "4":"DECLINED", "5":"ACCEPTED"},
        },

    MAP_TB2EAS : {
        //TB PRIORITY: 9 = LOW | 5 = NORMAL | 1 = HIGH
        PRIORITY : { "9":"0", "5":"1", "1":"2","unset":"1"}, //to Importance
        //TB CLASS: PUBLIC, PRIVATE, CONFIDENTIAL)
        CLASS : { "PUBLIC":"0", "PRIVATE":"2", "CONFIDENTIAL":"3", "unset":"1"}, //to Sensitivity
        //TB TRANSP : free = TRANSPARENT, busy = OPAQUE)
        TRANSP : {"TRANSPARENT":"0", "unset":"1", "OPAQUE":"2"}, // to BusyStatus
        //TB STATUS: NEEDS-ACTION, ACCEPTED, DECLINED, TENTATIVE, (DELEGATED, COMPLETED, IN-PROCESS - for todo)
        ATTENDEESTATUS : {"NEEDS-ACTION":"0", "ACCEPTED":"3", "DECLINED":"4", "TENTATIVE":"2", "DELEGATED":"5","COMPLETED":"5", "IN-PROCESS":"5"},
        },
    
    mapEasPropertyToThunderbird : function (easProp, tbProp, data, item) {
        if (data[easProp]) {
            //store original EAS value
            let easPropValue = eas.xmltools.checkString(data[easProp]);
            item.setProperty("X-EAS-" + easProp, easPropValue);
            //map EAS value to TB value  (use setCalItemProperty if there is one option which can unset/delete the property)
            eas.tools.setCalItemProperty(item, tbProp, eas.sync.MAP_EAS2TB[easProp][easPropValue]);
        }
    },

    mapThunderbirdPropertyToEas: function (tbProp, easProp, item) {
        if (item.hasProperty("X-EAS-" + easProp) && eas.tools.getCalItemProperty(item, tbProp) == eas.sync.MAP_EAS2TB[easProp][item.getProperty("X-EAS-" + easProp)]) {
            //we can use our stored EAS value, because it still maps to the current TB value
            return item.getProperty("X-EAS-" + easProp);
        } else {
            return eas.sync.MAP_TB2EAS[tbProp][eas.tools.getCalItemProperty(item, tbProp)]; 
        }
    },

    getEasItemType(aItem) {
        if (aItem instanceof tbSync.addressbook.AbItem) {
            return "Contacts";
        } else if (aItem instanceof tbSync.lightning.TbItem) {
            return aItem.isTodo ? "Tasks" : "Calendar";
        } else  {
            throw "Unknown aItem.";
        }
    },

    createItem(syncData) {
        switch (syncData.type) {
            case "Contacts":
                return syncData.target.createNewCard();
                break;
            
            case "Tasks":
                return syncData.target.createNewTodo();
                break;
            
            case "Calendar":
                return syncData.target.createNewEvent();
                break;
            
            default:
                throw "Unknown item type <" + syncData.type + ">";
        }
    },
    
    getWbxmlFromThunderbirdItem(item, syncData, isException = false) {
        try {
            let wbxml = eas.sync[syncData.type].getWbxmlFromThunderbirdItem(item, syncData, isException);
            return wbxml;
        } catch (e) {
            tbSync.errorlog.add("warning", syncData.errorInfo, "BadItemSkipped::JavaScriptError", item.toString());
            throw e; // unable to read item from Thunderbird - fatal error
        }        
    },







    // ---------------------------------------------------------------------------
    // LIGHTNING HELPER FUNCTIONS AND DEFINITIONS
    // These functions are needed only by tasks and events, so they
    // are placed here, even though they are not type independent,
    // but I did not want to add another "lightning" sub layer.
    //
    // The item in these functions is a native lightning item.
    // ---------------------------------------------------------------------------
        
    setItemSubject: function (item, syncData, data) {
        if (data.Subject) item.title = eas.xmltools.checkString(data.Subject);
    },
    
    setItemLocation: function (item, syncData, data) {
        if (data.Location) item.setProperty("location", eas.xmltools.checkString(data.Location));
    },


    setItemCategories: function (item, syncData, data) {
        if (data.Categories && data.Categories.Category) {
            let cats = [];
            if (Array.isArray(data.Categories.Category)) cats = data.Categories.Category;
            else cats.push(data.Categories.Category);
            item.setCategories(cats.length, cats);
        }
    },
    
    getItemCategories: function (item, syncData) {
        let asversion = syncData.accountData.getAccountProperty("asversion");
        let wbxml = eas.wbxmltools.createWBXML("", syncData.type); //init wbxml with "" and not with precodes, also activate type codePage (Calendar, Tasks, Contacts etc)

        //to properly "blank" categories, we need to always include the container
        let categories = item.getCategories({});
        if (categories.length > 0) {
            wbxml.otag("Categories");
                for (let i=0; i<categories.length; i++) wbxml.atag("Category", categories[i]);
            wbxml.ctag();
        } else {
            wbxml.atag("Categories");
        }
        return wbxml.getBytes();
    },


    setItemBody: function (item, syncData, data) {
        let asversion = syncData.accountData.getAccountProperty("asversion");
        if (asversion == "2.5") {
            if (data.Body) item.setProperty("description", eas.xmltools.checkString(data.Body));
        } else {
            if (data.Body && /* data.Body.EstimatedDataSize > 0  && */ data.Body.Data) item.setProperty("description", eas.xmltools.checkString(data.Body.Data)); //EstimatedDataSize is optional
        }
    },

    getItemBody: function (item, syncData) {
        let asversion = syncData.accountData.getAccountProperty("asversion");
        let wbxml = eas.wbxmltools.createWBXML("", syncData.type); //init wbxml with "" and not with precodes, also activate type codePage (Calendar, Tasks, Contacts etc)

        let description = (item.hasProperty("description")) ? item.getProperty("description") : "";
        if (asversion == "2.5") {
            wbxml.atag("Body", description);
        } else {
            wbxml.switchpage("AirSyncBase");
            wbxml.otag("Body");
                wbxml.atag("Type", "1");
                wbxml.atag("EstimatedDataSize", "" + description.length);
                wbxml.atag("Data", description);
            wbxml.ctag();
            //does not work with horde at the moment, does not work with task, does not work with exceptions
            //if (syncData.accountData.getAccountProperty("horde") == "0") wbxml.atag("NativeBodyType", "1");

            //return to code page of this type
            wbxml.switchpage(syncData.type);
        }
        return wbxml.getBytes();
    },

    setItemRecurrence: function (item, syncData, data) {
        if (data.Recurrence) {
            item.recurrenceInfo = tbSync.lightning.cal.createRecurrenceInfo();
            item.recurrenceInfo.item = item;
            let recRule = tbSync.lightning.cal.createRecurrenceRule();
            switch (data.Recurrence.Type) {
            case "0":
                recRule.type = "DAILY";
                break;
            case "1":
                recRule.type = "WEEKLY";
                break;
            case "2":
            case "3":
                recRule.type = "MONTHLY";
                break;
            case "5":
            case "6":
                recRule.type = "YEARLY";
                break;
            }

            if (data.Recurrence.CalendarType) {
                // TODO
            }
            if (data.Recurrence.DayOfMonth) {
                recRule.setComponent("BYMONTHDAY", 1, [data.Recurrence.DayOfMonth]);
            }
            if (data.Recurrence.DayOfWeek) {
                let DOW = data.Recurrence.DayOfWeek;
                if (DOW == 127 && (recRule.type == "MONTHLY" || recRule.type == "YEARLY")) {
                    recRule.setComponent("BYMONTHDAY", 1, [-1]);
                }
                else {
                    let days = [];
                    for (let i = 0; i < 7; ++i) {
                        if (DOW & 1 << i) days.push(i + 1);
                    }
                    if (data.Recurrence.WeekOfMonth) {
                        for (let i = 0; i < days.length; ++i) {
                            if (data.Recurrence.WeekOfMonth == 5) {
                                days[i] = -1 * (days[i] + 8);
                            }
                            else {
                                days[i] += 8 * (data.Recurrence.WeekOfMonth - 0);
                            }
                        }
                    }
                    recRule.setComponent("BYDAY", days.length, days);
                }
            }
            if (data.Recurrence.FirstDayOfWeek) {
                //recRule.setComponent("WKST", 1, [data.Recurrence.FirstDayOfWeek]); // WKST is not a valid component
                //recRule.weekStart = data.Recurrence.FirstDayOfWeek; // - (NS_ERROR_NOT_IMPLEMENTED) [calIRecurrenceRule.weekStart]
                tbSync.errorlog.add("info", syncData.errorInfo, "FirstDayOfWeek tag ignored (not supported).", item.icalString);                
            }

            if (data.Recurrence.Interval) {
                recRule.interval = data.Recurrence.Interval;
            }
            if (data.Recurrence.IsLeapMonth) {
                // TODO
            }
            if (data.Recurrence.MonthOfYear) {
                recRule.setComponent("BYMONTH", 1, [data.Recurrence.MonthOfYear]);
            }
            if (data.Recurrence.Occurrences) {
                recRule.count = data.Recurrence.Occurrences;
            }
            if (data.Recurrence.Until) {
                //time string could be in compact/basic or extended form of ISO 8601, 
                //cal.createDateTime only supports  compact/basic, our own method takes both styles
                recRule.untilDate = eas.tools.createDateTime(data.Recurrence.Until);
            }
            if (data.Recurrence.Start) {
                tbSync.errorlog.add("info", syncData.errorInfo, "Start tag in recurring task is ignored, recurrence will start with first entry.", item.icalString);
            }
        
            item.recurrenceInfo.insertRecurrenceItemAt(recRule, 0);

            if (data.Exceptions && syncData.type == "Calendar") { // only events, tasks cannot have exceptions
                // Exception could be an object or an array of objects
                let exceptions = [].concat(data.Exceptions.Exception);
                for (let exception of exceptions) {
                    let dateTime = tbSync.lightning.cal.createDateTime(exception.ExceptionStartTime);
                    if (data.AllDayEvent == "1") {
                        dateTime.isDate = true;
                        // Pass to replacement event unless overriden
                        if (!exception.AllDayEvent) {
                            exception.AllDayEvent = "1";
                        }
                    }
                    if (exception.Deleted == "1") {
                        item.recurrenceInfo.removeOccurrenceAt(dateTime);
                    }
                    else {
                        let replacement = item.recurrenceInfo.getOccurrenceFor(dateTime);
                        eas.sync.Calendar.setThunderbirdItemFromWbxml(replacement, exception, replacement.id, syncData);
                        // Reminders should carry over from parent, but setThunderbirdItemFromWbxml clears all alarms
                        if (!exception.Reminder && item.getAlarms({}).length) {
                            replacement.addAlarm(item.getAlarms({})[0]);
                        }
                        // Removing a reminder requires EAS 16.0
                        item.recurrenceInfo.modifyException(replacement, true);
                    }
                }
            }
        }
    },

    getItemRecurrence: function (item, syncData, localStartDate = null) {
        let asversion = syncData.accountData.getAccountProperty("asversion");
        let wbxml = eas.wbxmltools.createWBXML("", syncData.type); //init wbxml with "" and not with precodes, also activate type codePage (Calendar, Tasks etc)

        if (item.recurrenceInfo && (syncData.type == "Calendar" || syncData.type == "Tasks")) {
            let deleted = [];
            let hasRecurrence = false;
            let startDate = (syncData.type == "Calendar") ? item.startDate : item.entryDate;

            for (let recRule of item.recurrenceInfo.getRecurrenceItems({})) {
                if (recRule.date) {
                    if (recRule.isNegative) {
                        // EXDATE
                        deleted.push(recRule);
                    }
                    else {
                        // RDATE
                        tbSync.errorlog.add("info", syncData.errorInfo, "Ignoring RDATE rule (not supported)", recRule.icalString);
                    }
                    continue;
                }
                if (recRule.isNegative) {
                    // EXRULE
                    tbSync.errorlog.add("info", syncData.errorInfo, "Ignoring EXRULE rule (not supported)", recRule.icalString);
                    continue;
                }

                // RRULE
                wbxml.otag("Recurrence");
                hasRecurrence = true;

                let type = 0;
                let monthDays = recRule.getComponent("BYMONTHDAY", {});
                let weekDays  = recRule.getComponent("BYDAY", {});
                let months    = recRule.getComponent("BYMONTH", {});
                let weeks     = [];

                // Unpack 1MO style days
                for (let i = 0; i < weekDays.length; ++i) {
                    if (weekDays[i] > 8) {
                        weeks[i] = Math.floor(weekDays[i] / 8);
                        weekDays[i] = weekDays[i] % 8;
                    }
                    else if (weekDays[i] < -8) {
                        // EAS only supports last week as a special value, treat
                        // all as last week or assume every month has 5 weeks?
                        // Change to last week
                        //weeks[i] = 5;
                        // Assumes 5 weeks per month for week <= -2
                        weeks[i] = 6 - Math.floor(-weekDays[i] / 8);
                        weekDays[i] = -weekDays[i] % 8;
                    }
                }
                if (monthDays[0] && monthDays[0] == -1) {
                    weeks = [5];
                    weekDays = [1, 2, 3, 4, 5, 6, 7]; // 127
                    monthDays[0] = null;
                }
                // Type
                if (recRule.type == "WEEKLY") {
                    type = 1;
                    if (!weekDays.length) {
                        weekDays = [startDate.weekday + 1];
                    }
                }
                else if (recRule.type == "MONTHLY" && weeks.length) {
                    type = 3;
                }
                else if (recRule.type == "MONTHLY") {
                    type = 2;
                    if (!monthDays.length) {
                        monthDays = [startDate.day];
                    }
                }
                else if (recRule.type == "YEARLY" && weeks.length) {
                    type = 6;
                }
                else if (recRule.type == "YEARLY") {
                    type = 5;
                    if (!monthDays.length) {
                        monthDays = [startDate.day];
                    }
                    if (!months.length) {
                        months = [startDate.month + 1];
                    }
                }
                wbxml.atag("Type", type.toString());
                
                //Tasks need a Start tag, but we cannot allow a start date different from the start of the main item (thunderbird does not support that)
                if (localStartDate) wbxml.atag("Start", localStartDate);
                
                // TODO: CalendarType: 14.0 and up
                // DayOfMonth
                if (monthDays[0]) {
                    // TODO: Multiple days of month - multiple Recurrence tags?
                    wbxml.atag("DayOfMonth", monthDays[0].toString());
                }
                // DayOfWeek
                if (weekDays.length) {
                    let bitfield = 0;
                    for (let day of weekDays) {
                        bitfield |= 1 << (day - 1);
                    }
                    wbxml.atag("DayOfWeek", bitfield.toString());
                }
                // FirstDayOfWeek: 14.1 and up
                //wbxml.atag("FirstDayOfWeek", recRule.weekStart); - (NS_ERROR_NOT_IMPLEMENTED) [calIRecurrenceRule.weekStart]
                // Interval
                wbxml.atag("Interval", recRule.interval.toString());
                // TODO: IsLeapMonth: 14.0 and up
                // MonthOfYear
                if (months.length) {
                    wbxml.atag("MonthOfYear", months[0].toString());
                }
                // Occurrences
                if (recRule.isByCount) {
                    wbxml.atag("Occurrences", recRule.count.toString());
                }
                // Until
                else if (recRule.untilDate != null) {
                    //Events need the Until data in compact form, Tasks in the basic form
                    wbxml.atag("Until", eas.tools.getIsoUtcString(recRule.untilDate, (syncData.type == "Tasks")));
                }
                // WeekOfMonth
                if (weeks.length) {
                    wbxml.atag("WeekOfMonth", weeks[0].toString());
                }
                wbxml.ctag();
            }
            
            if (syncData.type == "Calendar" && hasRecurrence) { //Exceptions only allowed in Calendar and only if a valid Recurrence was added
                let modifiedIds = item.recurrenceInfo.getExceptionIds({});
                if (deleted.length || modifiedIds.length) {
                    wbxml.otag("Exceptions");
                    for (let exception of deleted) {
                        wbxml.otag("Exception");
                            wbxml.atag("ExceptionStartTime", eas.tools.getIsoUtcString(exception.date));
                            wbxml.atag("Deleted", "1");
                            //Docs say it is allowed, but if present, it does not work
                            //if (asversion == "2.5") {
                            //    wbxml.atag("UID", item.id);
                            //}
                        wbxml.ctag();
                    }
                    for (let exceptionId of modifiedIds) {
                        let replacement = item.recurrenceInfo.getExceptionFor(exceptionId);
                        wbxml.otag("Exception");
                            wbxml.atag("ExceptionStartTime", eas.tools.getIsoUtcString(exceptionId));
                            wbxml.append(eas.sync.getWbxmlFromThunderbirdItem(replacement, syncData, true));
                        wbxml.ctag();
                    }
                    wbxml.ctag();
                }
            }
        }

        return wbxml.getBytes();
    }

}
