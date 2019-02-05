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

eas.sync = {

    // ---------------------------------------------------------------------------
    // MAIN FUNCTIONS TO SYNC AN EAS FOLDER
    // ---------------------------------------------------------------------------

    start: Task.async (function* (syncdata)  {
        //sync
        syncdata.done = 0;
        syncdata.todo = 0;

        if (tbSync.db.getFolderSetting(syncdata.account, syncdata.folderID, "downloadonly") == "1") {		
            yield eas.sync.revertLocalChanges (syncdata);
        }

        yield eas.getItemEstimate (syncdata);
        yield eas.sync.requestRemoteChanges (syncdata); 

        if (tbSync.db.getFolderSetting(syncdata.account, syncdata.folderID, "downloadonly") != "1") {		
            yield eas.sync.sendLocalChanges (syncdata);
        }

        //if everything was OK, we still throw, to get into catch
        throw eas.finishSync();
    }),
    

    requestRemoteChanges: Task.async (function* (syncdata)  {
        syncdata.done = 0;
        do {
            tbSync.setSyncState("prepare.request.remotechanges", syncdata.account, syncdata.folderID);
            syncdata.request = "";
            syncdata.response = "";
        
            // BUILD WBXML
            let wbxml = tbSync.wbxmltools.createWBXML();
            wbxml.otag("Sync");
                wbxml.otag("Collections");
                    wbxml.otag("Collection");
                        if (tbSync.db.getAccountSetting(syncdata.account, "asversion") == "2.5") wbxml.atag("Class", syncdata.type);
                        wbxml.atag("SyncKey", syncdata.synckey);
                        wbxml.atag("CollectionId", syncdata.folderID);
                        wbxml.atag("DeletesAsMoves");
                        wbxml.atag("GetChanges");
                        wbxml.atag("WindowSize",  tbSync.prefSettings.getIntPref("eas.maxitems").toString());

                        if (tbSync.db.getAccountSetting(syncdata.account, "asversion") != "2.5") {
                            wbxml.otag("Options");
                                if (syncdata.type == "Calendar") wbxml.atag("FilterType", eas.tools.getFilterType());
                                wbxml.atag("Class", syncdata.type);
                                wbxml.switchpage("AirSyncBase");
                                wbxml.otag("BodyPreference");
                                    wbxml.atag("Type", "1");
                                wbxml.ctag();
                                wbxml.switchpage("AirSync");
                            wbxml.ctag();
                        } else if (syncdata.type == "Calendar") { //in 2.5 we only send it to filter Calendar
                            wbxml.otag("Options");
                                 wbxml.atag("FilterType", eas.tools.getFilterType());
                            wbxml.ctag();
                        }

                    wbxml.ctag();
                wbxml.ctag();
            wbxml.ctag();

            //SEND REQUEST
            tbSync.setSyncState("send.request.remotechanges", syncdata.account, syncdata.folderID);
            let response = yield eas.sendRequest(wbxml.getBytes(), "Sync", syncdata);

            //VALIDATE RESPONSE
            // get data from wbxml response, some servers send empty response if there are no changes, which is not an error
            let wbxmlData = eas.getDataFromResponse(response, eas.flags.allowEmptyResponse);
            if (wbxmlData === null) return;
        
            //check status, throw on error
            eas.checkStatus(syncdata, wbxmlData,"Sync.Collections.Collection.Status");

            //PROCESS COMMANDS        
            yield eas.sync.processCommands(wbxmlData, syncdata);

            //Update count in UI
            tbSync.setSyncState("eval.response.remotechanges", syncdata.account, syncdata.folderID);

            //update synckey
            eas.updateSynckey(syncdata, wbxmlData);
            
            if (!xmltools.hasWbxmlDataField(wbxmlData,"Sync.Collections.Collection.MoreAvailable")) {
                //Feedback from users: They want to see the final count
                yield tbSync.sleep(100, false);
                return;
            }
        } while (true);
                
    }),


    sendLocalChanges: Task.async (function* (syncdata)  {        
        let maxnumbertosend = tbSync.prefSettings.getIntPref("eas.maxitems");

        syncdata.done = 0;
        syncdata.todo = db.getItemsFromChangeLog(syncdata.targetId, 0, "_by_user").length;
        
        //keep track of failed items
        syncdata.failedItems = [];
        
        let done = false;
        let numberOfItemsToSend = maxnumbertosend;
        do {
            tbSync.setSyncState("prepare.request.localchanges", syncdata.account, syncdata.folderID);
            syncdata.request = "";
            syncdata.response = "";

            //get changed items from ChangeLog
            let changes = db.getItemsFromChangeLog(syncdata.targetId, numberOfItemsToSend, "_by_user");
            let c=0;
            let e=0;

            //keep track of send items during this request
            let changedItems = [];
            let addedItems = {};
            let sendItems = [];
                
            // BUILD WBXML
            let wbxml = tbSync.wbxmltools.createWBXML();
            wbxml.otag("Sync");
                wbxml.otag("Collections");
                    wbxml.otag("Collection");
                        if (tbSync.db.getAccountSetting(syncdata.account, "asversion") == "2.5") wbxml.atag("Class", syncdata.type);
                        wbxml.atag("SyncKey", syncdata.synckey);
                        wbxml.atag("CollectionId", syncdata.folderID);
                        wbxml.otag("Commands");

                            for (let i=0; i<changes.length; i++) if (!syncdata.failedItems.includes(changes[i].id)) {
                                //tbSync.dump("CHANGES",(i+1) + "/" + changes.length + " ("+changes[i].status+"," + changes[i].id + ")");
                                let items = null;
                                switch (changes[i].status) {

                                    case "added_by_user":
                                        items = yield syncdata.targetObj.getItem(changes[i].id);
                                        if (items.length > 0) {
                                            //filter out bad object types for this folder
                                            if (syncdata.type == eas.sync.getEasItemType(items[0])) {
                                                //create a temp clientId, to cope with too long or invalid clientIds (for EAS)
                                                let clientId = Date.now() + "-" + c;
                                                addedItems[clientId] = changes[i].id;
                                                sendItems.push({type: changes[i].status, id: changes[i].id});
                                                
                                                wbxml.otag("Add");
                                                wbxml.atag("ClientId", clientId); //Our temp clientId will get replaced by an id generated by the server
                                                    wbxml.otag("ApplicationData");
                                                        wbxml.switchpage(syncdata.type);

/*wbxml.atag("TimeZone", "xP///0UAdQByAG8AcABlAC8AQgBlAHIAbABpAG4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAoAAAAFAAIAAAAAAAAAAAAAAEUAdQByAG8AcABlAC8AQgBlAHIAbABpAG4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAAFAAEAAAAAAAAAxP///w==");
wbxml.atag("AllDayEvent", "0");
wbxml.switchpage("AirSyncBase");
wbxml.otag("Body");
    wbxml.atag("Type", "1");
    wbxml.atag("EstimatedDataSize", "0");
    wbxml.atag("Data");
wbxml.ctag();

wbxml.switchpage(syncdata.type);						
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

                                                        wbxml.append(eas.sync.getWbxmlFromThunderbirdItem(items[0], syncdata));
                                                        wbxml.switchpage("AirSync");
                                                    wbxml.ctag();
                                                wbxml.ctag();
                                                c++;
                                            } else {
                                                eas.sync.updateFailedItems(syncdata, "forbidden" + eas.sync.getEasItemType(items[0]) +"ItemIn" + syncdata.type + "Folder", items[0].id, items[0].icalString);
                                                e++;
                                            }
                                        } else {
                                            db.removeItemFromChangeLog(syncdata.targetId, changes[i].id);
                                        }
                                        break;
                                    
                                    case "modified_by_user":
                                        items = yield syncdata.targetObj.getItem(changes[i].id);
                                        if (items.length > 0) {
                                            //filter out bad object types for this folder
                                            if (syncdata.type == eas.sync.getEasItemType(items[0])) {
                                                wbxml.otag("Change");
                                                wbxml.atag("ServerId", changes[i].id);
                                                    wbxml.otag("ApplicationData");
                                                        wbxml.switchpage(syncdata.type);
                                                        wbxml.append(eas.sync.getWbxmlFromThunderbirdItem(items[0], syncdata));
                                                        wbxml.switchpage("AirSync");
                                                    wbxml.ctag();
                                                wbxml.ctag();
                                                changedItems.push(changes[i].id);
                                                sendItems.push({type: changes[i].status, id: changes[i].id});
                                                c++;
                                            } else {
                                                eas.sync.updateFailedItems(syncdata, "forbidden" + eas.sync.getEasItemType(items[0]) +"ItemIn" + syncdata.type + "Folder", items[0].id, items[0].icalString);
                                                e++;
                                            }
                                        } else {
                                            db.removeItemFromChangeLog(syncdata.targetId, changes[i].id);
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
                tbSync.setSyncState("send.request.localchanges", syncdata.account, syncdata.folderID);
                let response = yield eas.sendRequest(wbxml.getBytes(), "Sync", syncdata);
                
                tbSync.setSyncState("eval.response.localchanges", syncdata.account, syncdata.folderID);

                //get data from wbxml response
                let wbxmlData = eas.getDataFromResponse(response);
            
                //check status and manually handle error states which support softfails
                let errorcause = eas.checkStatus(syncdata, wbxmlData, "Sync.Collections.Collection.Status", "", true);
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
                                db.removeItemFromChangeLog(syncdata.targetId, sendItems[0].id);
                                tbSync.errorlog("warning", syncdata, "ErrorOnDelete::"+sendItems[0].id);
                            } else {
                                let foundItems = yield syncdata.targetObj.getItem(sendItems[0].id);                    
                                if (foundItems.length > 0) {
                                    eas.sync.updateFailedItems(syncdata, errorcause, foundItems[0].id, foundItems[0].icalString);
                                } else {
                                    //should not happen
                                    db.removeItemFromChangeLog(syncdata.targetId, sendItems[0].id);                                    
                                }
                            }
                            syncdata.done++;                            
                            //restore numberOfItemsToSend
                            numberOfItemsToSend = maxnumbertosend;                            
                        } else if (sendItems.length > 1) {
                            //reduce further
                            numberOfItemsToSend = Math.min(1, Math.round(sendItems.length / 5));
                        } else {
                            //sendItems.length == 0 ??? recheck but this time let it handle all cases
                            eas.checkStatus(syncdata, wbxmlData, "Sync.Collections.Collection.Status");
                        }
                        break;

                    default:
                        //recheck but this time let it handle all cases
                        eas.checkStatus(syncdata, wbxmlData, "Sync.Collections.Collection.Status");
                }        
                
                yield tbSync.sleep(10);

                if (errorcause == "") {
                    //PROCESS RESPONSE        
                    yield eas.sync.processResponses(wbxmlData, syncdata, addedItems, changedItems);
                
                    //PROCESS COMMANDS        
                    yield eas.sync.processCommands(wbxmlData, syncdata);

                    //remove all items from changelog that did not fail
                    for (let a=0; a < changedItems.length; a++) {
                        db.removeItemFromChangeLog(syncdata.targetId, changedItems[a]);
                        syncdata.done++;
                    }

                    //update synckey
                    eas.updateSynckey(syncdata, wbxmlData);
                }
            
            } else if (e==0) { //if there was no local change and also no error (which will not happen twice) finish

                done = true;

            }
        
        } while (!done);
        
        //was there an error?
        if (syncdata.failedItems.length > 0) {
            throw eas.finishSync("ServerRejectedSomeItems::" + syncdata.failedItems.length);                            
        }
        
    }),




    revertLocalChanges: Task.async (function* (syncdata)  {       
        let maxnumbertosend = tbSync.prefSettings.getIntPref("eas.maxitems");
        syncdata.done = 0;
        syncdata.todo = db.getItemsFromChangeLog(syncdata.targetId, 0, "_by_user").length;
        if (syncdata.todo == 0) {
            return;
        }

        let viaItemOperations = (tbSync.db.getAccountSetting(syncdata.account, "allowedEasCommands").split(",").includes("ItemOperations"));

        //get changed items from ChangeLog
        do {
            tbSync.setSyncState("prepare.request.revertlocalchanges", syncdata.account, syncdata.folderID);
            let changes = db.getItemsFromChangeLog(syncdata.targetId, maxnumbertosend, "_by_user");
            let c=0;
            syncdata.request = "";
            syncdata.response = "";
            
            // BUILD WBXML
            let wbxml = tbSync.wbxmltools.createWBXML();
            if (viaItemOperations) {
                wbxml.switchpage("ItemOperations");
                wbxml.otag("ItemOperations");
            } else {
                wbxml.otag("Sync");
                    wbxml.otag("Collections");
                        wbxml.otag("Collection");
                            if (tbSync.db.getAccountSetting(syncdata.account, "asversion") == "2.5") wbxml.atag("Class", syncdata.type);
                            wbxml.atag("SyncKey", syncdata.synckey);
                            wbxml.atag("CollectionId", syncdata.folderID);
                            wbxml.otag("Commands");
            }

            for (let i=0; i<changes.length; i++) {
                let items = null;
                let ServerId = changes[i].id;
                let foundItems = yield syncdata.targetObj.getItem(ServerId);

                switch (changes[i].status) {
                    case "added_by_user": //remove
                        if (foundItems.length > 0) {
                            db.addItemToChangeLog(syncdata.targetId, ServerId, "deleted_by_server"); //this changelog entry will be removed by the listener
                            yield syncdata.targetObj.deleteItem(foundItems[0]);
                        }
                    break;
                    
                    case "modified_by_user":
                        if (foundItems.length > 0) { //delete item so it can be replaced with a fresh copy, the changelog entry will be changed from modified to deleted
                            yield syncdata.targetObj.deleteItem(foundItems[0]);
                        }
                    case "deleted_by_user":
                        if (viaItemOperations) {
                            wbxml.otag("Fetch");
                                wbxml.atag("Store", "Mailbox");
                                wbxml.switchpage("AirSync");
                                wbxml.atag("CollectionId", syncdata.folderID);
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
                    tbSync.setSyncState("send.request.revertlocalchanges", syncdata.account, syncdata.folderID);
                    let response = yield eas.sendRequest(wbxml.getBytes(), (viaItemOperations) ? "ItemOperations" : "Sync", syncdata);
                        
                    tbSync.setSyncState("eval.response.revertlocalchanges", syncdata.account, syncdata.folderID);

                    //get data from wbxml response
                    wbxmlData = eas.getDataFromResponse(response);
                } catch (e) {
                    //we do not handle errors, IF there was an error, wbxmlData is empty and will trigger the fallback
                }
                
                let fetchPath = (viaItemOperations) ? "ItemOperations.Response.Fetch" : "Sync.Collections.Collection.Responses.Fetch";
                if (xmltools.hasWbxmlDataField(wbxmlData, fetchPath)) {
                
                    //looking for additions
                    let add = xmltools.nodeAsArray(xmltools.getWbxmlDataField(wbxmlData, fetchPath));
                    for (let count = 0; count < add.length; count++) {
                        yield tbSync.sleep(2);

                        let ServerId = add[count].ServerId;
                        let data = (viaItemOperations) ? add[count].Properties : add[count].ApplicationData;
                        
                        if (data && ServerId) {
                            let foundItems = yield syncdata.targetObj.getItem(ServerId);
                            if (foundItems.length == 0) { //do NOT add, if an item with that ServerId was found
                                    let newItem = eas.sync[syncdata.type].createItem();
                                    try {
                                        eas.sync[syncdata.type].setThunderbirdItemFromWbxml(newItem, data, ServerId, syncdata);
                                        db.addItemToChangeLog(syncdata.targetId, ServerId, "added_by_server");
                                        yield syncdata.targetObj.adoptItem(newItem);
                                    } catch (e) {
                                        xmltools.printXmlData(add[count], true); //include application data in log                  
                                        tbSync.errorlog("warning", syncdata, "BadItemSkipped::JavaScriptError", newItem.icalString);
                                        throw e; // unable to add item to Thunderbird - fatal error
                                    }
                            } else {
                                //should not happen, since we deleted that item beforehand
                                db.removeItemFromChangeLog(syncdata.targetId, ServerId);
                            }
                            syncdata.done++;
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
                        tbSync.errorlog("info", syncdata, "Server returned error during ItemOperations.Fetch, falling back to Sync.Fetch.");
                    } else {
                        yield eas.sync.revertLocalChangesViaResync(syncdata);
                        return;
                    }
                }
                            
            } else { //if there was no more local change we need to revert, return

                return;

            }
        
        } while (true);
        
    }),

    revertLocalChangesViaResync: Task.async (function* (syncdata) {
        tbSync.errorlog("info", syncdata, "Server does not support ItemOperations.Fetch and/or Sync.Fetch, must revert via resync.");
        let changes = db.getItemsFromChangeLog(syncdata.targetId, 0, "_by_user");
        syncdata.done = 0;
        syncdata.todo = changes.length;
        tbSync.setSyncState("prepare.request.revertlocalchanges", syncdata.account, syncdata.folderID);
        
        //remove all changes, so we can get them fresh from the server
        for (let i=0; i<changes.length; i++) {
            let items = null;
            let ServerId = changes[i].id;
            let foundItems = yield syncdata.targetObj.getItem(ServerId);
            if (foundItems.length > 0) { //delete item with that ServerId
                db.addItemToChangeLog(syncdata.targetId, ServerId, "deleted_by_server");
                yield syncdata.targetObj.deleteItem(foundItems[0]);
            } else {
                db.removeItemFromChangeLog(syncdata.targetId, ServerId);
            }
            syncdata.done++;
        }
        throw tbSync.eas.finishSync("RevertViaFolderResync", eas.flags.resyncFolder); //This will resync a fresh copy from the server
    }),




    // ---------------------------------------------------------------------------
    // SUB FUNCTIONS CALLED BY  MAIN FUNCTION
    // ---------------------------------------------------------------------------
    
    processCommands:  Task.async (function* (wbxmlData, syncdata)  {
        //any commands for us to work on? If we reach this point, Sync.Collections.Collection is valid, 
        //no need to use the save getWbxmlDataField function
        if (wbxmlData.Sync.Collections.Collection.Commands) {
        
            //looking for additions
            let add = xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Commands.Add);
            for (let count = 0; count < add.length; count++) {
                yield tbSync.sleep(2);

                let ServerId = add[count].ServerId;
                let data = add[count].ApplicationData;

                let foundItems = yield syncdata.targetObj.getItem(ServerId);
                if (foundItems.length == 0) {
                    //do NOT add, if an item with that ServerId was found
                    let newItem = eas.sync[syncdata.type].createItem();
                    try {
                        eas.sync[syncdata.type].setThunderbirdItemFromWbxml(newItem, data, ServerId, syncdata);
                        db.addItemToChangeLog(syncdata.targetId, ServerId, "added_by_server");
                        yield syncdata.targetObj.adoptItem(newItem); //yield pcal.addItem(newItem); //We are not using the added item after is has been added, so we might be faster using adoptItem
                    } catch (e) {
                        xmltools.printXmlData(add[count], true); //include application data in log                  
                        tbSync.errorlog("warning", syncdata, "BadItemSkipped::JavaScriptError", newItem.icalString);
                        throw e; // unable to add item to Thunderbird - fatal error
                    }
                } else {
                    tbSync.errorlog("info", syncdata, "Add request, but element exists already, skipped.", ServerId);
                }
                syncdata.done++;
            }

            //looking for changes
            let upd = xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Commands.Change);
            //inject custom change object for debug
            //upd = JSON.parse('[{"ServerId":"2tjoanTeS0CJ3QTsq5vdNQAAAAABDdrY6Gp03ktAid0E7Kub3TUAAAoZy4A1","ApplicationData":{"DtStamp":"20171109T142149Z"}}]');
            for (let count = 0; count < upd.length; count++) {
                yield tbSync.sleep(2);

                let ServerId = upd[count].ServerId;
                let data = upd[count].ApplicationData;

                syncdata.done++;
                let foundItems = yield syncdata.targetObj.getItem(ServerId);
                if (foundItems.length > 0) { //only update, if an item with that ServerId was found
                                        
                    let keys = Object.keys(data);
                    //replace by smart merge
                    if (keys.length == 1 && keys[0] == "DtStamp") tbSync.dump("DtStampOnly", keys); //ignore DtStamp updates (fix with smart merge)
                    else {
                        
                        if (tbSync.db.getItemStatusFromChangeLog(syncdata.targetId, ServerId) !== null) {
                            tbSync.errorlog("info", syncdata, "Change request from server, but also local modifications, server wins!", ServerId);
                        }
                        
                        let newItem = foundItems[0].clone();
                        try {
                            eas.sync[syncdata.type].setThunderbirdItemFromWbxml(newItem, data, ServerId, syncdata);
                            db.addItemToChangeLog(syncdata.targetId, ServerId, "modified_by_server");
                            yield syncdata.targetObj.modifyItem(newItem, foundItems[0]);
                        } catch (e) {
                            tbSync.errorlog("warning", syncdata, "BadItemSkipped::JavaScriptError", newItem.icalString);
                            xmltools.printXmlData(upd[count], true);  //include application data in log                   
                            throw e; // unable to mod item to Thunderbird - fatal error
                        }
                    }
                    
                }
            }
            
            //looking for deletes
            let del = xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Commands.Delete).concat(xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Commands.SoftDelete));
            for (let count = 0; count < del.length; count++) {
                yield tbSync.sleep(2);

                let ServerId = del[count].ServerId;

                let foundItems = yield syncdata.targetObj.getItem(ServerId);
                if (foundItems.length > 0) { //delete item with that ServerId
                    db.addItemToChangeLog(syncdata.targetId, ServerId, "deleted_by_server");
                    yield syncdata.targetObj.deleteItem(foundItems[0]);
                }
                syncdata.done++;
            }
        
        }
    }),


    updateFailedItems: function (syncdata, cause, id, data) {                
        //something is wrong with this item, move it to the end of changelog and go on
        if (!syncdata.failedItems.includes(id)) {
            //the extra parameter true will re-add the item to the end of the changelog
            db.removeItemFromChangeLog(syncdata.targetId, id, true);                        
            syncdata.failedItems.push(id);            
            tbSync.errorlog("info", syncdata, "BadItemSkipped::" + tbSync.getLocalizedMessage("status." + cause ,"eas"), "\n\nRequest:\n" + syncdata.request + "\n\nResponse:\n" + syncdata.response + "\n\nElement:\n" + data);
        }
    },


    processResponses:  Task.async (function* (wbxmlData, syncdata, addedItems, changedItems)  {
            //any responses for us to work on?  If we reach this point, Sync.Collections.Collection is valid, 
            //no need to use the save getWbxmlDataField function
            if (wbxmlData.Sync.Collections.Collection.Responses) {

                //looking for additions (Add node contains, status, old ClientId and new ServerId)
                let add = xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Responses.Add);
                for (let count = 0; count < add.length; count++) {
                    yield tbSync.sleep(2);

                    //get the true Thunderbird UID of this added item (we created a temp clientId during add)
                    add[count].ClientId = addedItems[add[count].ClientId];

                    //look for an item identfied by ClientId and update its id to the new id received from the server
                    let foundItems = yield syncdata.targetObj.getItem(add[count].ClientId);                    
                    if (foundItems.length > 0) {

                        //Check status, stop sync if bad, allow soft fail
                        let errorcause = eas.checkStatus(syncdata, add[count],"Status","Sync.Collections.Collection.Responses.Add["+count+"].Status", true);
                        if (errorcause !== "") {
                            //something is wrong with this item, move it to the end of changelog and go on
                            eas.sync.updateFailedItems(syncdata, errorcause, foundItems[0].id, foundItems[0].icalString);
                        } else {
                            let newItem = foundItems[0].clone();
                            newItem.id = add[count].ServerId;
                            db.removeItemFromChangeLog(syncdata.targetId, add[count].ClientId);
                            db.addItemToChangeLog(syncdata.targetId, newItem.id, "modified_by_server");
                            yield syncdata.targetObj.modifyItem(newItem, foundItems[0]);
                            syncdata.done++;
                        }

                    }
                }

                //looking for modifications 
                let upd = xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Responses.Change);
                for (let count = 0; count < upd.length; count++) {
                    let foundItems = yield syncdata.targetObj.getItem(upd[count].ServerId);                    
                    if (foundItems.length > 0) {

                        //Check status, stop sync if bad, allow soft fail
                        let errorcause = eas.checkStatus(syncdata, upd[count],"Status","Sync.Collections.Collection.Responses.Change["+count+"].Status", true);
                        if (errorcause !== "") {
                            //something is wrong with this item, move it to the end of changelog and go on
                            eas.sync.updateFailedItems(syncdata, errorcause, foundItems[0].id, foundItems[0].icalString);
                            //also remove from changedItems
                            let p = changedItems.indexOf(upd[count].ServerId);
                            if (p>-1) changedItems.splice(p,1);
                        }

                    }
                }

                //looking for deletions 
                let del = xmltools.nodeAsArray(wbxmlData.Sync.Collections.Collection.Responses.Delete);
                for (let count = 0; count < del.length; count++) {
                    //What can we do about failed deletes? SyncLog
                    eas.checkStatus(syncdata, del[count],"Status","Sync.Collections.Collection.Responses.Delete["+count+"].Status", true);
                }
                
            }
    }),










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
            let easPropValue = xmltools.checkString(data[easProp]);
            item.setProperty("X-EAS-" + easProp, easPropValue);
            //map EAS value to TB value  (use setCalItemProperty if there is one option which can unset/delete the property)
            tbSync.setCalItemProperty(item, tbProp, eas.sync.MAP_EAS2TB[easProp][easPropValue]);
        }
    },

    mapThunderbirdPropertyToEas: function (tbProp, easProp, item) {
        if (item.hasProperty("X-EAS-" + easProp) && tbSync.getCalItemProperty(item, tbProp) == eas.sync.MAP_EAS2TB[easProp][item.getProperty("X-EAS-" + easProp)]) {
            //we can use our stored EAS value, because it still maps to the current TB value
            return item.getProperty("X-EAS-" + easProp);
        } else {
            return eas.sync.MAP_TB2EAS[tbProp][tbSync.getCalItemProperty(item, tbProp)]; 
        }
    },

    getEasItemType(aItem) {
        if (aItem.card) {
            return "Contacts";
        }
        switch (tbSync.getItemType(aItem)) {
            case "tb-event": 
                return "Calendar";
            case "tb-todo": 
                return "Tasks";
            default:
                return "Unknown";
        }        
    },

    getWbxmlFromThunderbirdItem(item, syncdata, isException = false) {
        try {
            let wbxml = eas.sync[syncdata.type].getWbxmlFromThunderbirdItem(item, syncdata, isException);
            return wbxml;
        } catch (e) {
            tbSync.errorlog("warning", syncdata, "BadItemSkipped::JavaScriptError", item.icalString);
            throw e; // unable to read item from Thunderbird - fatal error
        }        
    },







    // ---------------------------------------------------------------------------
    // LIGHTNING HELPER FUNCTIONS AND DEFINITIONS
    // These functions are needed only by tasks and events, so they
    // are placed here, even though they are not type independent,
    // but I did not want to add another "lightning" sub layer.
    // ---------------------------------------------------------------------------
        
    setItemSubject: function (item, syncdata, data) {
        if (data.Subject) item.title = xmltools.checkString(data.Subject);
    },
    
    setItemLocation: function (item, syncdata, data) {
        if (data.Location) item.setProperty("location", xmltools.checkString(data.Location));
    },


    setItemCategories: function (item, syncdata, data) {
        if (data.Categories && data.Categories.Category) {
            let cats = [];
            if (Array.isArray(data.Categories.Category)) cats = data.Categories.Category;
            else cats.push(data.Categories.Category);
            item.setCategories(cats.length, cats);
        }
    },
    
    getItemCategories: function (item, syncdata) {
        let asversion = tbSync.db.getAccountSetting(syncdata.account, "asversion");
        let wbxml = tbSync.wbxmltools.createWBXML("", syncdata.type); //init wbxml with "" and not with precodes, also activate type codePage (Calendar, Tasks, Contacts etc)

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


    setItemBody: function (item, syncdata, data) {
        let asversion = tbSync.db.getAccountSetting(syncdata.account, "asversion");
        if (asversion == "2.5") {
            if (data.Body) item.setProperty("description", xmltools.checkString(data.Body));
        } else {
            if (data.Body && data.Body.EstimatedDataSize > 0 && data.Body.Data) item.setProperty("description", xmltools.checkString(data.Body.Data)); //CLEAR??? DataSize>0 ?? TODO
        }
    },

    getItemBody: function (item, syncdata) {
        let asversion = tbSync.db.getAccountSetting(syncdata.account, "asversion");
        let wbxml = tbSync.wbxmltools.createWBXML("", syncdata.type); //init wbxml with "" and not with precodes, also activate type codePage (Calendar, Tasks, Contacts etc)

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
            //if (tbSync.db.getAccountSetting(syncdata.account, "horde") == "0") wbxml.atag("NativeBodyType", "1");

            //return to code page of this type
            wbxml.switchpage(syncdata.type);
        }
        return wbxml.getBytes();
    },

    setItemRecurrence: function (item, syncdata, data) {
        if (data.Recurrence) {
            item.recurrenceInfo = cal.createRecurrenceInfo();
            item.recurrenceInfo.item = item;
            let recRule = cal.createRecurrenceRule();
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
                tbSync.errorlog("info", syncdata, "FirstDayOfWeek tag ignored (not supported).", item.icalString);                
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
                tbSync.errorlog("info", syncdata, "Start tag in recurring task is ignored, recurrence will start with first entry.", item.icalString);
            }
        
            item.recurrenceInfo.insertRecurrenceItemAt(recRule, 0);

            if (data.Exceptions && syncdata.type == "Calendar") { // only events, tasks cannot have exceptions
                // Exception could be an object or an array of objects
                let exceptions = [].concat(data.Exceptions.Exception);
                for (let exception of exceptions) {
                    let dateTime = cal.createDateTime(exception.ExceptionStartTime);
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
                        eas.sync.Calendar.setThunderbirdItemFromWbxml(replacement, exception, replacement.id, syncdata);
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

    getItemRecurrence: function (item, syncdata, localStartDate = null) {
        let asversion = tbSync.db.getAccountSetting(syncdata.account, "asversion");
        let wbxml = tbSync.wbxmltools.createWBXML("", syncdata.type); //init wbxml with "" and not with precodes, also activate type codePage (Calendar, Tasks etc)

        if (item.recurrenceInfo && (syncdata.type == "Calendar" || syncdata.type == "Tasks")) {
            let deleted = [];
            let hasRecurrence = false;
            let startDate = (syncdata.type == "Calendar") ? item.startDate : item.entryDate;

            for (let recRule of item.recurrenceInfo.getRecurrenceItems({})) {
                if (recRule.date) {
                    if (recRule.isNegative) {
                        // EXDATE
                        deleted.push(recRule);
                    }
                    else {
                        // RDATE
                        tbSync.errorlog("info", syncdata, "Ignoring RDATE rule (not supported)", recRule.icalString);
                    }
                    continue;
                }
                if (recRule.isNegative) {
                    // EXRULE
                    tbSync.errorlog("info", syncdata, "Ignoring EXRULE rule (not supported)", recRule.icalString);
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
                    wbxml.atag("Until", eas.tools.getIsoUtcString(recRule.untilDate, (syncdata.type == "Tasks")));
                }
                // WeekOfMonth
                if (weeks.length) {
                    wbxml.atag("WeekOfMonth", weeks[0].toString());
                }
                wbxml.ctag();
            }
            
            if (syncdata.type == "Calendar" && hasRecurrence) { //Exceptions only allowed in Calendar and only if a valid Recurrence was added
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
                            wbxml.append(eas.sync.getWbxmlFromThunderbirdItem(replacement, syncdata, true));
                        wbxml.ctag();
                    }
                    wbxml.ctag();
                }
            }
        }

        return wbxml.getBytes();
    }

}
