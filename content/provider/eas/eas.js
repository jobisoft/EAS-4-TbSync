/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";


var eas = {




    /**
     * Called to get passwords of accounts of this provider
     *
     * @param accountdata       [in] account data structure
     */
    getPassword: function (accountdata) {
        let host4PasswordManager = tbSync.getHost4PasswordManager(accountdata.provider, accountdata.host);
        return tbSync.getLoginInfo(host4PasswordManager, "TbSync", accountdata.user);
    },



    /**
     * Called to set passwords of accounts of this provider
     *
     * @param accountdata       [in] account data structure
     * @param newPassword       [in] new password
     */
    setPassword: function (accountdata, newPassword) {
        let host4PasswordManager = tbSync.getHost4PasswordManager(accountdata.provider, accountdata.host);
        tbSync.setLoginInfo(host4PasswordManager, "TbSync", accountdata.user, newPassword);
    },

    


    /**
     * Returns an array of folder settings, that should survive disable and re-enable
     */
    getPersistentFolderSettings: function () {
        return ["targetName", "targetColor", "downloadonly"];
    },



    /**
     * Return the thunderbird type (tb-contact, tb-event, tb-todo) for a given folder type of this provider. A provider could have multiple 
     * type definitions for a single thunderbird type (default calendar, shared address book, etc), this maps all possible provider types to
     * one of the three thunderbird types.
     *
     * @param type       [in] provider folder type
     */
    getThunderbirdFolderType: function(type) {
        switch (type) {
            case "9": 
            case "14": 
                return "tb-contact";
            case "8":
            case "13":
                return "tb-event";
            case "7":
            case "15":
                return "tb-todo";
            default:
                return "unknown ("+type + ")";
        };
    },


    


    /**
     * Is called if TbSync needs to create a new thunderbird address book associated with an account of this provider.
     *
     * @param newname       [in] name of the new address book
     * @param account       [in] id of the account this address book belongs to
     * @param folderID      [in] id of the folder this address book belongs to (sync target)
     *
     * return the id of the newAddressBook 
     */
    createAddressBook: function (newname, account, folderID) {
        let dirPrefId = MailServices.ab.newAddressBook(newname, "", 2);  /* kPABDirectory - return abManager.newAddressBook(name, "moz-abmdbdirectory://", 2); */
        let directory = MailServices.ab.getDirectoryFromId(dirPrefId);

        if (directory && directory instanceof Components.interfaces.nsIAbDirectory && directory.dirPrefId == dirPrefId) {
            directory.setStringValue("tbSyncIcon", "eas");
            return directory;		
        }
        return null;
    },



    /**
     * Is called if TbSync needs to create a new UID for an address book card
     *
     * @param aItem       [in] card that needs new ID
     *
     * returns the new id 
     */
    getNewCardID: function (aItem, folder) {
        return aItem.localId;
    },



    /**
     * Is called if TbSync needs to create a new lightning calendar associated with an account of this provider.
     *
     * @param newname       [in] name of the new calendar
     * @param account       [in] id of the account this calendar belongs to
     * @param folderID      [in] id of the folder this calendar belongs to (sync target)
     */
    createCalendar: function(newname, account, folderID) {
        let calManager = cal.getCalendarManager();
        //Alternative calendar, which uses calTbSyncCalendar
        //let newCalendar = calManager.createCalendar("TbSync", Services.io.newURI('tbsync-calendar://'));

        //Create the new standard calendar with a unique name
        let newCalendar = calManager.createCalendar("storage", Services.io.newURI("moz-storage-calendar://"));
        newCalendar.id = cal.getUUID();
        newCalendar.name = newname;

        newCalendar.setProperty("color", tbSync.db.getFolderSetting(account, folderID, "targetColor"));
        newCalendar.setProperty("relaxedMode", true); //sometimes we get "generation too old for modifyItem", check can be disabled with relaxedMode
        newCalendar.setProperty("calendar-main-in-composite",true);
        newCalendar.setProperty("readOnly", tbSync.db.getFolderSetting(account, folderID, "downloadonly") == "1");
        calManager.registerCalendar(newCalendar);

        //is there an email identity we can associate this calendar to? 
        //getIdentityKey returns "" if none found, which removes any association
        let key = tbSync.getIdentityKey(tbSync.db.getAccountSetting(account, "user"));
        newCalendar.setProperty("fallbackOrganizerName", newCalendar.getProperty("organizerCN"));
        newCalendar.setProperty("imip.identity.key", key);
        if (key === "") {
            //there is no matching email identity - use current default value as best guess and remove association
            //use current best guess 
            newCalendar.setProperty("organizerCN", newCalendar.getProperty("fallbackOrganizerName"));
            newCalendar.setProperty("organizerId", cal.email.prependMailTo(tbSync.db.getAccountSetting(account, "user")));
        }
        
        return newCalendar;
    },



    /**
     * Is called if TbSync needs to find contacts in the global address list (GAL / directory) of an account associated with this provider.
     * It is used for autocompletion while typing something into the address field of the message composer and for the address book search,
     * if something is typed into the search field of the Thunderbird address book.
     *
     * DO NOT IMPLEMENT AT ALL, IF NOT SUPPORTED
     *
     * TbSync will execute this only for queries longer than 3 chars.
     *
     * @param account       [in] id of the account which should be searched
     * @param currentQuery  [in] search query
     * @param caller        [in] "autocomplete" or "search"
    
     */
    abServerSearch: async function (account, currentQuery, caller)  {
        if (!tbSync.db.getAccountSetting(account, "allowedEasCommands").split(",").includes("Search")) {
            return null;
        }

        if (caller == "autocomplete" && tbSync.db.getAccountSetting(account, "galautocomplete") != "1") {
            return null;
        }
        
        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.switchpage("Search");
        wbxml.otag("Search");
            wbxml.otag("Store");
                wbxml.atag("Name", "GAL");
                wbxml.atag("Query", currentQuery);
                wbxml.otag("Options");
                    wbxml.atag("Range", "0-99"); //Z-Push needs a Range
                    //Not valid for GAL: https://msdn.microsoft.com/en-us/library/gg675461(v=exchg.80).aspx
                    //wbxml.atag("DeepTraversal");
                    //wbxml.atag("RebuildResults");
                wbxml.ctag();
            wbxml.ctag();
        wbxml.ctag();

        let syncData = {};
        syncData.account = account;
        syncData.folderID = "";
        syncData.syncstate = "SearchingGAL";
        
            
        let response = await eas.network.sendRequest(wbxml.getBytes(), "Search", syncData);
        let wbxmlData = eas.getDataFromResponse(response);
        let galdata = [];

        if (wbxmlData.Search && wbxmlData.Search.Response && wbxmlData.Search.Response.Store && wbxmlData.Search.Response.Store.Result) {
            let results = eas.xmltools.nodeAsArray(wbxmlData.Search.Response.Store.Result);
            let accountname = tbSync.db.getAccountSetting(account, "accountname");
        
            for (let count = 0; count < results.length; count++) {
                if (results[count].Properties) {
                    //tbSync.window.console.log('Found contact:' + results[count].Properties.DisplayName);
                    let resultset = {};

                    switch (caller) {
                        case "search":
                            resultset.properties = {};                    
                            resultset.properties["FirstName"] = results[count].Properties.FirstName;
                            resultset.properties["LastName"] = results[count].Properties.LastName;
                            resultset.properties["DisplayName"] = results[count].Properties.DisplayName;
                            resultset.properties["PrimaryEmail"] = results[count].Properties.EmailAddress;
                            resultset.properties["CellularNumber"] = results[count].Properties.MobilePhone;
                            resultset.properties["HomePhone"] = results[count].Properties.HomePhone;
                            resultset.properties["WorkPhone"] = results[count].Properties.Phone;
                            resultset.properties["Company"] = accountname; //results[count].Properties.Company;
                            resultset.properties["Department"] = results[count].Properties.Title;
                            resultset.properties["JobTitle"] = results[count].Properties.Office;
                            break;
                       
                        case "autocomplete":
                            resultset.autocomplete = {};                    
                            resultset.autocomplete.value = results[count].Properties.DisplayName + " <" + results[count].Properties.EmailAddress + ">";
                            resultset.autocomplete.account = account;
                            break;
                    }
                    
                    galdata.push(resultset);
                }
            }
        }
        
        return galdata;
    },



    /**
     * Is called if TbSync needs to synchronize an account.
     *
     * @param syncData      [in] object that contains the account and maybe the folder which needs to worked on
     *                           you are free to add more fields to this object which you need (persistent) during sync
     * @param job           [in] identifier about what is to be done, the standard job is "sync", you are free to add
     *                           custom jobs like "deletefolder" via your own accountSettings.xul
     */
    start: async function (syncData, job)  {
        let accountReSyncs = 0;
        
        do {
            try {
                accountReSyncs++;
                syncData.todo = 0;
                syncData.done = 0;

                if (accountReSyncs > 3) {
                    throw eas.sync.finishSync("resync-loop", eas.flags.abortWithError);
                }

                // check if enabled
                if (!tbSync.isEnabled(syncData.account)) {
                    throw eas.sync.finishSync("disabled", eas.flags.abortWithError);
                }

                // check if connection has data
                let connection = eas.network.getAuthData(syncData);
                if (connection.host == "" || connection.user == "") {
                    throw eas.sync.finishSync("nouserhost", eas.flags.abortWithError);
                }
                

                switch (job) {
                    case "sync":

                        //get all folders, which need to be synced
                        await eas.getPendingFolders(syncData);
                        //update folder list in GUI
                        Services.obs.notifyObservers(null, "tbsync.updateFolderList", syncData.account);
                        //sync all pending folders
                        await eas.syncPendingFolders(syncData); //inside here we throw and catch FinischFolderSync
                        throw eas.sync.finishSync();
                        break;
                        
                    case "deletefolder":
                        //TODO: foldersync first ???
                        await eas.deleteFolder(syncData);
                        //update folder list in GUI
                        Services.obs.notifyObservers(null, "tbsync.updateFolderList", syncData.account);
                        throw eas.sync.finishSync();
                        break;
                        
                    default:
                        throw eas.sync.finishSync("unknown", eas.flags.abortWithError);

                }

            } catch (report) { 
                    
                switch (report.type) {
                    case eas.flags.resyncAccount:
                        tbSync.errorlog("info", syncData, "Forced Account Resync", report.message);                        
                        continue;

                    case eas.flags.abortWithServerError: 
                        //Could not connect to server. Can we rerun autodiscover? If not, fall through to abortWithError              
                        if (syncData.accountData.getAccountProperty("servertype") == "auto") {
                            let errorcode = await eas.updateServerConnectionViaAutodiscover(syncData);
                            switch (errorcode) {
                                case 401:
                                case 403: //failed to authenticate
                                    report.message = "401"
                                    tbSync.finishAccountSync(syncData, report);
                                    return;                            
                                case 200: //server and/or user was updated, retry
                                    Services.obs.notifyObservers(null, "tbsync.updateAccountSettingsGui", syncData.account);
                                    continue;
                                default: //autodiscover failed, fall through to abortWithError
                            }                        
                        }

                    case eas.flags.abortWithError: //fatal error, finish account sync
                    case eas.flags.syncNextFolder: //no more folders left, finish account sync
                    case eas.flags.resyncFolder: //should not happen here, just in case
                        tbSync.finishAccountSync(syncData, report);
                        return;

                    default:
                        //there was some other error
                        report.type = "JavaScriptError";
                        tbSync.finishAccountSync(syncData, report);
                        Components.utils.reportError(report);
                        return;
                }

            }

        } while (true);

    },





    // * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    // * HELPER FUNCTIONS BEYOND THE API
    // * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    
    getPendingFolders: async function (syncData)  {
        //this function sets all folders which ougth to be synced to pending, either a specific one (if folderID is set) or all avail
        if (syncData.folderID != "") {
            //just set the specified folder to pending
            tbSync.db.setFolderSetting(syncData.account, syncData.folderID, "status", "pending");
        } else {
            //scan all folders and set the enabled ones to pending
           syncData.setSyncState("prepare.request.folders", syncData.account); 
            let foldersynckey = syncData.accountData.getAccountProperty("foldersynckey");

            //build WBXML to request foldersync
            let wbxml = eas.wbxmltools.createWBXML();
            wbxml.switchpage("FolderHierarchy");
            wbxml.otag("FolderSync");
                wbxml.atag("SyncKey", foldersynckey);
            wbxml.ctag();

           syncData.setSyncState("send.request.folders", syncData.account); 
            let response = await eas.network.sendRequest(wbxml.getBytes(), "FolderSync", syncData);

           syncData.setSyncState("eval.response.folders", syncData.account); 
            let wbxmlData = eas.getDataFromResponse(response);

            eas.network.checkStatus(syncData, wbxmlData,"FolderSync.Status");

            let synckey = eas.xmltools.getWbxmlDataField(wbxmlData,"FolderSync.SyncKey");
            if (synckey) {
                syncData.accountData.setAccountProperty("foldersynckey", synckey);
            } else {
                throw eas.sync.finishSync("wbxmlmissingfield::FolderSync.SyncKey", eas.flags.abortWithError);
            }
            
            //if we reach this point, wbxmlData contains FolderSync node, so the next if will not fail with an javascript error, 
            //no need to use save getWbxmlDataField function
            
            //are there any changes in folder hierarchy
            if (wbxmlData.FolderSync.Changes) {
                //looking for additions
                let add = eas.xmltools.nodeAsArray(wbxmlData.FolderSync.Changes.Add);
                for (let count = 0; count < add.length; count++) {
                    //only add allowed folder types to DB
                    if (!["9","14","8","13","7","15","4"].includes(add[count].Type)) 
                        continue;

                    let existingFolder = tbSync.db.getFolder(syncData.account, add[count].ServerId);
                    if (existingFolder !== null && existingFolder.cached == "0") {
                        //there was an error at the server, he has send us an ADD for a folder we alreay have, treat as update
                        tbSync.db.setFolderSetting(existingFolder.account, existingFolder.folderID, "name", add[count].DisplayName);
                        tbSync.db.setFolderSetting(existingFolder.account, existingFolder.folderID, "type", add[count].Type);
                        tbSync.db.setFolderSetting(existingFolder.account, existingFolder.folderID, "parentID", add[count].ParentId);
                    } else {
                        //create folder obj for new  folder settings
                        let newFolder = {};

                        newFolder.folderID = add[count].ServerId;
                        newFolder.name = add[count].DisplayName;
                        newFolder.type = add[count].Type;
                        newFolder.parentID = add[count].ParentId;

                        //if there is a cached version of this folderID, addFolder will merge all persistent settings - all other settings not defined here will be set to their defaults
                        tbSync.db.addFolder(syncData.account, newFolder);
                    }
                }
                
                //looking for updates
                let update = eas.xmltools.nodeAsArray(wbxmlData.FolderSync.Changes.Update);
                for (let count = 0; count < update.length; count++) {
                    //get a reference
                    let folder = tbSync.db.getFolder(syncData.account, update[count].ServerId);
                    if (folder !== null) {
                        //update folder
                        tbSync.db.setFolderSetting(folder.account, folder.folderID, "name", update[count].DisplayName);
                        tbSync.db.setFolderSetting(folder.account, folder.folderID, "type", update[count].Type);
                        tbSync.db.setFolderSetting(folder.account, folder.folderID, "parentID", update[count].ParentId);
                    }
                }

                //looking for deletes
                let del = eas.xmltools.nodeAsArray(wbxmlData.FolderSync.Changes.Delete);
                for (let count = 0; count < del.length; count++) {

                    let folder = tbSync.db.getFolder(syncData.account, del[count].ServerId);
                    if (folder !== null) {
                        tbSync.takeTargetOffline("eas", folder, "[deleted from server]");
                    }
                }
            }

            tbSync.prepareFoldersForSync(syncData.account);            
        }
    },



    getNextPendingFolder: function (accountID) {
        //using getSortedData, to sync in the same order as shown in the list
        let sortedFolders = eas.folderList.getSortedData(accountID);       
        for (let i=0; i < sortedFolders.length; i++) {
            if (sortedFolders[i].statusCode != "pending") continue;
            return tbSync.db.getFolder(accountID, sortedFolders[i].folderID);
        }
        return null;
    },


    //Process all folders with PENDING status
    syncPendingFolders: async function (syncData)  {
        let folderReSyncs = 1;
        
        do {                
            //any pending folders left?
            let nextFolder = eas.getNextPendingFolder(syncData.account);
            if (nextFolder === null) {
                //all folders of this account have been synced
                return;
            };

            //The individual folder sync is placed inside a try ... catch block. If a folder sync has finished, a throwFinishSync error is thrown
            //and catched here. If that error has a message attached, it ist re-thrown to the main account sync loop, which will abort sync completely
            let calendarReadOnlyStatus = null;
            try {
                
                //resync loop control
                if (syncData.folderID == nextFolder.folderID) folderReSyncs++;
                else folderReSyncs = 1;
                syncData.folderID = nextFolder.folderID;;

                if (folderReSyncs > 3) {
                    throw eas.sync.finishSync("resync-loop");
                }

                //get syncData type, which is also used in WBXML for the CLASS element
                syncData.type = null;
                switch (eas.getThunderbirdFolderType(nextFolder.type)) {
                    case "tb-contact": 
                        syncData.type = "Contacts";
                        // check SyncTarget
                        if (!tbSync.checkAddressbook(syncData.account, syncData.folderID)) {
                            throw eas.sync.finishSync("notargets");
                        }
                        break;
                        
                    case "tb-event":
                        if (syncData.type === null) syncData.type = "Calendar";
                    case "tb-todo":
                        if (syncData.type === null) syncData.type = "Tasks";

                        // skip if lightning is not installed
                        if (tbSync.lightningIsAvailable() == false) {
                            throw eas.sync.finishSync("nolightning");
                        }
                        
                        // check SyncTarget
                        if (!tbSync.checkCalender(syncData.account, syncData.folderID)) {
                            throw eas.sync.finishSync("notargets");
                        }                        
                        break;
                        
                    default:
                        throw eas.sync.finishSync("skipped");
                };





               syncData.setSyncState("preparing", syncData.account, syncData.folderID);
                
                //get synckey if needed
                syncData.synckey = nextFolder.synckey;                
                if (syncData.synckey == "") {
                    await eas.getSynckey(syncData);
                }
                
                //sync folder
                syncData.timeOfLastSync = tbSync.db.getFolderSetting(syncData.account, syncData.folderID, "lastsynctime") / 1000;
                syncData.timeOfThisSync = (Date.now() / 1000) - 1;
                
                switch (syncData.type) {
                    case "Contacts": 
                        //get sync target of this addressbook
                        syncData.targetId = tbSync.db.getFolderSetting(syncData.account, syncData.folderID, "target");
                        syncData.addressbookObj = tbSync.getAddressBookObject(syncData.targetId);

                        //promisify addressbook, so it can be used together with await
                        syncData.targetObj = eas.tools.promisifyAddressbook(syncData.addressbookObj);
                        
                        await eas.sync.start(syncData);   //using new tbsync contacts sync code
                        break;

                    case "Calendar":
                    case "Tasks": 
                        syncData.targetId = tbSync.db.getFolderSetting(syncData.account, syncData.folderID, "target");
                        syncData.calendarObj = cal.getCalendarManager().getCalendarById(syncData.targetId);
                        
                        //promisify calender, so it can be used together with await
                        syncData.targetObj = cal.async.promisifyCalendar(syncData.calendarObj.wrappedJSObject);

                        syncData.calendarObj.startBatch();
                        //save current value of readOnly (or take it from the setting
                        calendarReadOnlyStatus = syncData.calendarObj.getProperty("readOnly") || (tbSync.db.getFolderSetting(syncData.account, syncData.folderID, "downloadonly") == "1");                       
                        syncData.calendarObj.setProperty("readOnly", false);
                        await eas.sync.start(syncData);
                        break;
                }

            } catch (report) { 
                
                if (calendarReadOnlyStatus !== null) { //null, true, false
                    syncData.calendarObj.setProperty("readOnly", calendarReadOnlyStatus);
                    syncData.calendarObj.endBatch();
                }
                
                switch (report.type) {
                    case eas.flags.abortWithError:  //if there was a fatal error during folder sync, re-throw error to finish account sync (with error)
                    case eas.flags.abortWithServerError:
                    case eas.flags.resyncAccount:   //if the entire account needs to be resynced, finish this folder and re-throw account (re)sync request                                                    
                        tbSync.finishFolderSync(syncData, report);
                        throw report;
                        break;

                    case eas.flags.syncNextFolder:
                        tbSync.finishFolderSync(syncData, report);
                        break;
                                            
                    case eas.flags.resyncFolder:
                        if (report.message == "RevertViaFolderResync") {
                            //the user requested to throw away local modifications, no need to backup, just invalidate the synckey
                            eas.onResetTarget(syncData.account, syncData.folderID);
                        } else {
                            //takeTargetOffline will backup the current folder and on next run, a fresh copy 
                            //of the folder will be synced down - the folder itself is NOT deleted (4th arg is false)
                            tbSync.errorlog("info", syncData, "Forced Folder Resync", report.message + "\n\n" + report.details);
                            tbSync.takeTargetOffline("eas", tbSync.db.getFolder(syncData.account, syncData.folderID), "[forced folder resync]", false);
                        }
                        continue;
                    
                    default:
                        report.type = "JavaScriptError";
                        tbSync.finishFolderSync(syncData, report);
                        //this is a fatal error, re-throw error to finish account sync
                        throw report;
                }
            }

        }
        while (true);
    },



    //WBXML FUNCTIONS
 

    getSynckey: async function (syncData) {
       syncData.setSyncState("prepare.request.synckey", syncData.account);
        //build WBXML to request a new syncKey
        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.otag("Sync");
            wbxml.otag("Collections");
                wbxml.otag("Collection");
                    if (syncData.accountData.getAccountProperty("asversion") == "2.5") wbxml.atag("Class", syncData.type);
                    wbxml.atag("SyncKey","0");
                    wbxml.atag("CollectionId",syncData.folderID);
                wbxml.ctag();
            wbxml.ctag();
        wbxml.ctag();
        
       syncData.setSyncState("send.request.synckey", syncData.account);
        let response = await eas.network.sendRequest(wbxml.getBytes(), "Sync", syncData);

       syncData.setSyncState("eval.response.synckey", syncData.account);
        // get data from wbxml response
        let wbxmlData = eas.getDataFromResponse(response);
        //check status
        eas.network.checkStatus(syncData, wbxmlData,"Sync.Collections.Collection.Status");
        //update synckey
        eas.updateSynckey(syncData, wbxmlData);
    },

    getItemEstimate: async function (syncData)  {
        syncData.todo = -1;
        
        if (!syncData.accountData.getAccountProperty("allowedEasCommands").split(",").includes("GetItemEstimate")) {
            return; //do not throw, this is optional
        }
        
       syncData.setSyncState("prepare.request.estimate", syncData.account, syncData.folderID);
        
        // BUILD WBXML
        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.switchpage("GetItemEstimate");
        wbxml.otag("GetItemEstimate");
            wbxml.otag("Collections");
                wbxml.otag("Collection");
                    if (syncData.accountData.getAccountProperty("asversion") == "2.5") { //got order for 2.5 directly from Microsoft support
                        wbxml.atag("Class", syncData.type); //only 2.5
                        wbxml.atag("CollectionId", syncData.folderID);
                        wbxml.switchpage("AirSync");
                        wbxml.atag("FilterType", eas.tools.getFilterType());
                        wbxml.atag("SyncKey", syncData.synckey);
                        wbxml.switchpage("GetItemEstimate");
                    } else { //14.0
                        wbxml.switchpage("AirSync");
                        wbxml.atag("SyncKey", syncData.synckey);
                        wbxml.switchpage("GetItemEstimate");
                        wbxml.atag("CollectionId", syncData.folderID);
                        wbxml.switchpage("AirSync");
                        wbxml.otag("Options");
                            if (syncData.type == "Calendar") wbxml.atag("FilterType", eas.tools.getFilterType());
                            wbxml.atag("Class", syncData.type);
                        wbxml.ctag();
                        wbxml.switchpage("GetItemEstimate");
                    }
                wbxml.ctag();
            wbxml.ctag();
        wbxml.ctag();

        //SEND REQUEST
       syncData.setSyncState("send.request.estimate", syncData.account, syncData.folderID);
        let response = await eas.network.sendRequest(wbxml.getBytes(), "GetItemEstimate", syncData, /* allowSoftFail */ true);

        //VALIDATE RESPONSE
       syncData.setSyncState("eval.response.estimate", syncData.account, syncData.folderID);

        // get data from wbxml response, some servers send empty response if there are no changes, which is not an error
        let wbxmlData = eas.getDataFromResponse(response, eas.flags.allowEmptyResponse);
        if (wbxmlData === null) return;

        let status = eas.xmltools.getWbxmlDataField(wbxmlData, "GetItemEstimate.Response.Status");
        let estimate = eas.xmltools.getWbxmlDataField(wbxmlData, "GetItemEstimate.Response.Collection.Estimate");

        if (status && status == "1") { //do not throw on error, with EAS v2.5 I get error 2 for tasks and calendars ???
            syncData.todo = estimate;
        }
    },

    getUserInfo: async function (syncData)  {
        if (!syncData.accountData.getAccountProperty("allowedEasCommands").split(",").includes("Settings")) {
            return;
        }

       syncData.setSyncState("prepare.request.getuserinfo", syncData.account);

        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.switchpage("Settings");
        wbxml.otag("Settings");
            wbxml.otag("UserInformation");
                wbxml.atag("Get");
            wbxml.ctag();
        wbxml.ctag();

       syncData.setSyncState("send.request.getuserinfo", syncData.account);
        let response = await eas.network.sendRequest(wbxml.getBytes(), "Settings", syncData);


       syncData.setSyncState("eval.response.getuserinfo", syncData.account);
        let wbxmlData = eas.getDataFromResponse(response);

        eas.network.checkStatus(syncData, wbxmlData,"Settings.Status");
    },



    deleteFolder: async function (syncData)  {
        if (syncData.folderID == "") {
            throw eas.sync.finishSync();
        } 
        
        if (!syncData.accountData.getAccountProperty("allowedEasCommands").split(",").includes("FolderDelete")) {
            throw eas.sync.finishSync("notsupported::FolderDelete", eas.flags.abortWithError);
        }

       syncData.setSyncState("prepare.request.deletefolder", syncData.account);
        let foldersynckey = syncData.accountData.getAccountProperty("foldersynckey");

        //request foldersync
        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.switchpage("FolderHierarchy");
        wbxml.otag("FolderDelete");
            wbxml.atag("SyncKey", foldersynckey);
            wbxml.atag("ServerId", syncData.folderID);
        wbxml.ctag();

       syncData.setSyncState("send.request.deletefolder", syncData.account);
        let response = await eas.network.sendRequest(wbxml.getBytes(), "FolderDelete", syncData);


       syncData.setSyncState("eval.response.deletefolder", syncData.account);
        let wbxmlData = eas.getDataFromResponse(response);

        eas.network.checkStatus(syncData, wbxmlData,"FolderDelete.Status");

        let synckey = eas.xmltools.getWbxmlDataField(wbxmlData,"FolderDelete.SyncKey");
        if (synckey) {
            syncData.accountData.setAccountProperty("foldersynckey", synckey);
            //this folder is not synced, no target to take care of, just remove the folder
            tbSync.db.deleteFolder(syncData.account, syncData.folderID);
            syncData.folderID = "";
            //update manager gui / folder list
            Services.obs.notifyObservers(null, "tbsync.updateFolderList", syncData.account);
            throw eas.sync.finishSync();
        } else {
            throw eas.sync.finishSync("wbxmlmissingfield::FolderDelete.SyncKey", eas.flags.abortWithError);
        }
    },


    
    updateSynckey: function (syncData, wbxmlData) {
        let synckey = eas.xmltools.getWbxmlDataField(wbxmlData,"Sync.Collections.Collection.SyncKey");

        if (synckey) {
            syncData.synckey = synckey;
            db.setFolderSetting(syncData.account, syncData.folderID, "synckey", synckey);
        } else {
            throw eas.sync.finishSync("wbxmlmissingfield::Sync.Collections.Collection.SyncKey", eas.flags.abortWithError);
        }
    },

        

 


    parentIsTrash: function (account, parentID) {
        if (parentID == "0") return false;
        if (tbSync.db.getFolder(account, parentID) && tbSync.db.getFolder(account, parentID).type == "4") return true;
        return false;
    },
    
    TimeZoneDataStructure : class {
        constructor() {
            this.buf = new DataView(new ArrayBuffer(172));
        }
        
/*		
        Buffer structure:
            @000    utcOffset (4x8bit as 1xLONG)

            @004     standardName (64x8bit as 32xWCHAR)
            @068     standardDate (16x8 as 1xSYSTEMTIME)
            @084     standardBias (4x8bit as 1xLONG)

            @088     daylightName (64x8bit as 32xWCHAR)
            @152    daylightDate (16x8 as 1xSTRUCT)
            @168    daylightBias (4x8bit as 1xLONG)
*/
        
        set easTimeZone64 (b64) {
            //clear buffer
            for (let i=0; i<172; i++) this.buf.setUint8(i, 0);
            //load content into buffer
            let content = (b64 == "") ? "" : atob(b64);
            for (let i=0; i<content.length; i++) this.buf.setUint8(i, content.charCodeAt(i));
        }
        
        get easTimeZone64 () {
            let content = "";
            for (let i=0; i<172; i++) content += String.fromCharCode(this.buf.getUint8(i));
            return (btoa(content));
        }
        
        getstr (byteoffset) {
            let str = "";
            //walk thru the buffer in 32 steps of 16bit (wchars)
            for (let i=0;i<32;i++) {
                let cc = this.buf.getUint16(byteoffset+i*2, true);
                if (cc == 0) break;
                str += String.fromCharCode(cc);
            }
            return str;
        }

        setstr (byteoffset, str) {
            //clear first
            for (let i=0;i<32;i++) this.buf.setUint16(byteoffset+i*2, 0);
            //walk thru the buffer in steps of 16bit (wchars)
            for (let i=0;i<str.length && i<32; i++) this.buf.setUint16(byteoffset+i*2, str.charCodeAt(i), true);
        }
        
        getsystemtime (buf, offset) {
            let systemtime = {
                get wYear () { return buf.getUint16(offset + 0, true); },
                get wMonth () { return buf.getUint16(offset + 2, true); },
                get wDayOfWeek () { return buf.getUint16(offset + 4, true); },
                get wDay () { return buf.getUint16(offset + 6, true); },
                get wHour () { return buf.getUint16(offset + 8, true); },
                get wMinute () { return buf.getUint16(offset + 10, true); },
                get wSecond () { return buf.getUint16(offset + 12, true); },
                get wMilliseconds () { return buf.getUint16(offset + 14, true); },
                toString() { return [this.wYear, this.wMonth, this.wDay].join("-") + ", " + this.wDayOfWeek + ", " + [this.wHour,this.wMinute,this.wSecond].join(":") + "." + this.wMilliseconds},

                set wYear (v) { buf.setUint16(offset + 0, v, true); },
                set wMonth (v) { buf.setUint16(offset + 2, v, true); },
                set wDayOfWeek (v) { buf.setUint16(offset + 4, v, true); },
                set wDay (v) { buf.setUint16(offset + 6, v, true); },
                set wHour (v) { buf.setUint16(offset + 8, v, true); },
                set wMinute (v) { buf.setUint16(offset + 10, v, true); },
                set wSecond (v) { buf.setUint16(offset + 12, v, true); },
                set wMilliseconds (v) { buf.setUint16(offset + 14, v, true); },
                };
            return systemtime;
        }
        
        get standardDate () {return this.getsystemtime (this.buf, 68); }
        get daylightDate () {return this.getsystemtime (this.buf, 152); }
            
        get utcOffset () { return this.buf.getInt32(0, true); }
        set utcOffset (v) { this.buf.setInt32(0, v, true); }

        get standardBias () { return this.buf.getInt32(84, true); }
        set standardBias (v) { this.buf.setInt32(84, v, true); }
        get daylightBias () { return this.buf.getInt32(168, true); }
        set daylightBias (v) { this.buf.setInt32(168, v, true); }
        
        get standardName () {return this.getstr(4); }
        set standardName (v) {return this.setstr(4, v); }
        get daylightName () {return this.getstr(88); }
        set daylightName (v) {return this.setstr(88, v); }
        
        toString () { return ["", 
            "utcOffset: "+ this.utcOffset,
            "standardName: "+ this.standardName,
            "standardDate: "+ this.standardDate.toString(),
            "standardBias: "+ this.standardBias,
            "daylightName: "+ this.daylightName,
            "daylightDate: "+ this.daylightDate.toString(),
            "daylightBias: "+ this.daylightBias].join("\n"); }
    },

    

    


    
    


    
    
    
    /**
     * Functions used by the folderlist in the main account settings tab
     */
    folderList: {

        /**
         * Is called before the context menu of the folderlist is shown, allows to 
         * show/hide custom menu options based on selected folder
         *
         * @param document       [in] document object of the account settings window
         * @param folder         [in] folder databasse object of the selected folder
         */
        onContextMenuShowing: function (document, folder) {
            let hideContextMenuDelete = true;

            if (folder !== null) {
                //if a folder in trash is selected, also show ContextMenuDelete (but only if FolderDelete is allowed)
                if (eas.parentIsTrash(folder.account, folder.parentID) && tbSync.db.getAccountSetting(folder.account, "allowedEasCommands").split(",").includes("FolderDelete")) {// folder in recycle bin
                    hideContextMenuDelete = false;
                    document.getElementById("TbSync.eas.FolderListContextMenuDelete").label = tbSync.getString("deletefolder.menuentry::" + folder.name, "eas");
                }                
            }

            document.getElementById("TbSync.eas.FolderListContextMenuDelete").hidden = hideContextMenuDelete;
        },



        /**
         * Returns an array of folderRowData objects, containing all information needed 
         * to fill the folderlist. The content of the folderRowData object is free to choose,
         * it will be passed back to getRow() and updateRow()
         *
         * @param account        [in] account id for which the folder data should be returned
         */
        getSortedData: function (account) {
            let folderData = [];
            let folders = tbSync.db.getFolders(account);
            let allowedTypesOrder = ["9","14","8","13","7","15"];
            let folderIDs = Object.keys(folders).filter(f => allowedTypesOrder.includes(folders[f].type)).sort((a, b) => (eas.folderList.getIdChain(allowedTypesOrder, account, a).localeCompare(eas.folderList.getIdChain(allowedTypesOrder, account, b))));
            
            for (let i=0; i < folderIDs.length; i++) {
                folderData.push(eas.folderList.getRowData(folders[folderIDs[i]]));
            }
            return folderData;
        },



        /**
         * Returns a folderRowData object, containing all information needed to fill one row
         * in the folderlist. The content of the folderRowData object is free to choose, it
         * will be passed back to getRow() and updateRow()
         *
         * Use tbSync.getSyncStatusMsg(folder, syncData, provider) to get a nice looking 
         * status message, including sync progress (if folder is synced)
         *
         * @param folder         [in] folder databasse object of requested folder
         * @param syncData       [in] optional syncData obj send by updateRow(),
         *                            needed to check if the folder is currently synced
         */
        getRowData: function (folder, syncData = null) {
            let rowData = {};
            rowData.account = folder.account;
            rowData.folderID = folder.folderID;
            rowData.selected = (folder.selected == "1");
            rowData.type = folder.type;
            rowData.name = folder.name;
            rowData.downloadonly = folder.downloadonly;
            rowData.statusCode = folder.status;
            rowData.statusMsg = tbSync.getSyncStatusMsg(folder, syncData, "eas");

            if (eas.parentIsTrash(folder.account, folder.parentID)) rowData.name = tbSync.getString("recyclebin", "eas") + " | " + rowData.name;

            return rowData;
        },
    


        /**
         * Returns an array of attribute objects, which define the number of columns 
         * and the look of the header
         */
        getHeader: function () {
            return [
                {style: "font-weight:bold;", label: "", width: "93"},
                {style: "font-weight:bold;", label: tbSync.getString("manager.resource"), width:"150"},
                {style: "font-weight:bold;", label: tbSync.getString("manager.status"), flex :"1"},
            ]
        },

        //not part of API
        updateReadOnly: function (event) {
            let p = event.target.parentNode.parentNode;
            let account = p.getAttribute('account');
            let folderID = p.getAttribute('folderID');
            let value = event.target.value;
            let type = tbSync.db.getFolderSetting(account, folderID, "type");

            //update value
            tbSync.db.setFolderSetting(account, folderID, "downloadonly", value);

            //update icon
            if (value == "0") {
                p.setAttribute('image','chrome://tbsync/skin/acl_rw.png');
            } else {
                p.setAttribute('image','chrome://tbsync/skin/acl_ro.png');
            }
                
            //update ro flag if calendar
            switch (type) {
                case "8":
                case "13":
                case "7":
                case "15":
                    {
                        let target = tbSync.db.getFolderSetting(account, folderID, "target");
                        if (target != "") {
                            let calManager = cal.getCalendarManager();
                            let targetCal = calManager.getCalendarById(target); 
                            targetCal.setProperty("readOnly", value == '1');
                        }
                    }
                break;
            }
        },

        /**
         * Is called to add a row to the folderlist. After this call, updateRow is called as well.
         *
         * @param document        [in] document object of the account settings window
         * @param rowData         [in] rowData object with all information needed to add the row
         * @param itemSelCheckbox [in] a checkbox object which can be used to allow the user to select/deselect this resource
         */        
        getRow: function (document, rowData, itemSelCheckbox) {
            //checkbox
            itemSelCheckbox.setAttribute("style", "margin: 0px 0px 0px 3px;");

            //icon
            let itemType = document.createElement("image");
            itemType.setAttribute("src", eas.folderList.getTypeImage(rowData));
            itemType.setAttribute("style", "margin: 0px 9px 0px 3px;");

            //read/write access             
            let itemACL = document.createElement("button");
            itemACL.setAttribute("image", "chrome://tbsync/skin/acl_" + (rowData.downloadonly == "1" ? "ro" : "rw") + ".png");
            itemACL.setAttribute("class", "plain");
            itemACL.setAttribute("style", "width: 35px; min-width: 35px; margin: 0; height:26px");
            itemACL.setAttribute("account", rowData.account);
            itemACL.setAttribute("folderID", rowData.folderID);
            itemACL.setAttribute("type", "menu");
            let menupopup = document.createElement("menupopup");
                let menuitem1 = document.createElement("menuitem");
                menuitem1.setAttribute("value", "1");
                menuitem1.setAttribute("class", "menuitem-iconic");
                menuitem1.setAttribute("label", tbSync.getString("acl.readonly", "eas"));
                menuitem1.setAttribute("image", "chrome://tbsync/skin/acl_ro2.png");
                menuitem1.addEventListener("command", eas.folderList.updateReadOnly);

                let menuitem2 = document.createElement("menuitem");
                menuitem2.setAttribute("value", "0");
                menuitem2.setAttribute("class", "menuitem-iconic");
                menuitem2.setAttribute("label", tbSync.getString("acl.readwrite", "eas"));
                menuitem2.setAttribute("image", "chrome://tbsync/skin/acl_rw2.png");
                menuitem2.addEventListener("command", eas.folderList.updateReadOnly);

                menupopup.appendChild(menuitem2);
                menupopup.appendChild(menuitem1);
            itemACL.appendChild(menupopup);
            
            //folder name
            let itemLabel = document.createElement("description");
            itemLabel.setAttribute("disabled", !rowData.selected);

            //status
            let itemStatus = document.createElement("description");
            itemStatus.setAttribute("disabled", !rowData.selected);

            //group1
            let itemHGroup1 = document.createElement("hbox");
            itemHGroup1.setAttribute("align", "center");
            itemHGroup1.appendChild(itemSelCheckbox);
            itemHGroup1.appendChild(itemType);
            itemHGroup1.appendChild(itemACL);

            let itemVGroup1 = document.createElement("vbox");
            itemVGroup1.setAttribute("width", "93");
            itemVGroup1.appendChild(itemHGroup1);

            //group2
            let itemHGroup2 = document.createElement("hbox");
            itemHGroup2.setAttribute("align", "center");
            itemHGroup2.setAttribute("width", "146");
            itemHGroup2.appendChild(itemLabel);

            let itemVGroup2 = document.createElement("vbox");
            itemVGroup2.setAttribute("style", "padding: 3px");
            itemVGroup2.appendChild(itemHGroup2);

            //group3
            let itemHGroup3 = document.createElement("hbox");
            itemHGroup3.setAttribute("align", "center");
            itemHGroup3.setAttribute("width", "200");
            itemHGroup3.appendChild(itemStatus);

            let itemVGroup3 = document.createElement("vbox");
            itemVGroup3.setAttribute("style", "padding: 3px");
            itemVGroup3.appendChild(itemHGroup3);

            //final row
            let row = document.createElement("hbox");
            row.setAttribute("style", "min-height: 24px;");
            row.appendChild(itemVGroup1);
            row.appendChild(itemVGroup2);            
            row.appendChild(itemVGroup3);            
            return row;             
        },		



        /**
         * Is called to update a row of the folderlist (the first cell is a select checkbox inserted by TbSync)
         *
         * @param document       [in] document object of the account settings window
         * @param listItem       [in] the listitem of the row, which needs to be updated
         * @param rowData        [in] rowData object with all information needed to add the row
         */        
        updateRow: function (document, item, rowData) {
            //acl image
            item.childNodes[0].childNodes[0].childNodes[0].childNodes[2].setAttribute("image", "chrome://tbsync/skin/acl_" + (rowData.downloadonly == "1" ? "ro" : "rw") + ".png");

            //select checkbox
            if (rowData.selected) {
                item.childNodes[0].childNodes[0].childNodes[0].childNodes[0].setAttribute("checked", true);
            } else {
                item.childNodes[0].childNodes[0].childNodes[0].childNodes[0].removeAttribute("checked");
            }

            if (item.childNodes[0].childNodes[1].childNodes[0].textContent != rowData.name) item.childNodes[0].childNodes[1].childNodes[0].textContent = rowData.name;
            if (item.childNodes[0].childNodes[2].childNodes[0].textContent != rowData.statusMsg) item.childNodes[0].childNodes[2].childNodes[0].textContent = rowData.statusMsg;
            item.childNodes[0].childNodes[1].childNodes[0].setAttribute("disabled", !rowData.selected);
            item.childNodes[0].childNodes[1].childNodes[0].setAttribute("style", rowData.selected ? "" : "font-style:italic");
            item.childNodes[0].childNodes[2].childNodes[0].setAttribute("style", rowData.selected ? "" : "font-style:italic");
        },


  







        //BEYOND API

        //Custom stuff, outside of interface, invoked by own functions in overlayed accountSettings.xul
        getIdChain: function (allowedTypesOrder, account, _folderID) {
            let folderID = _folderID;
            
            //create sort string so that child folders are directly below their parent folders, different folder types are grouped and trashed folders at the end
            let chain = folderID.toString().padStart(3,"0");
            let folder = tbSync.db.getFolder(account, folderID);
            
            while (folder && folder.parentID && folder.parentID != "0") {
                chain = folder.parentID.toString().padStart(3,"0") + "." + chain;
                folderID = folder.parentID;
                folder = tbSync.db.getFolder(account, folderID);
            }
            
            if (folder && folder.type) {
                let pos = allowedTypesOrder.indexOf(folder.type);
                chain = ((pos == -1) ? "ZZZ" : pos.toString().padStart(3,"0")) + "." + chain;
            }
            
            return chain;
        },
        
        deleteFolder: function(document, account) {
            let folderList = document.getElementById("tbsync.accountsettings.folderlist");
            if (folderList.selectedItem !== null && !folderList.disabled) {
                let fID =  folderList.selectedItem.value;
                let folder = tbSync.db.getFolder(account, fID, true);

                //only trashed folders can be purged (for example O365 does not show deleted folders but also does not allow to purge them)
                if (!eas.parentIsTrash(account, folder.parentID)) return;
                
                if (folder.selected == "1") document.defaultView.alert(tbSync.getString("deletefolder.notallowed::" + folder.name, "eas"));
                else if (document.defaultView.confirm(tbSync.getString("deletefolder.confirm::" + folder.name, "eas"))) {
                tbSync.syncAccount("deletefolder", account, fID);
                } 
            }            
        },
    }
    
};
    
