/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */

"use strict";

// Every object in here will be loaded into tbSync.providers.<providername>.
const eas = tbSync.providers.eas;

eas.prefs = Services.prefs.getBranch("extensions.eas4tbsync.");

//use flags instead of strings to avoid errors due to spelling errors
eas.flags = Object.freeze({
    allowEmptyResponse: true, 
});

eas.windowsTimezoneMap = {};
eas.cachedTimezoneData = null;
eas.defaultTimezoneInfo = null;
eas.defaultTimezone = null;
eas.utcTimezone = null;


/**
 * Implementing the TbSync interface for external provider extensions.
 */
var Base = class {
    /**
     * Called during load of external provider extension to init provider.
     */
    static async load() {
        // Set default prefs
        let branch = Services.prefs.getDefaultBranch("extensions.eas4tbsync.");
        branch.setIntPref("timeout", 90000);
        branch.setIntPref("maxitems", 50);
        branch.setCharPref("clientID.type", "TbSync");
        branch.setCharPref("clientID.useragent", "Thunderbird ActiveSync");    

        eas.defaultTimezone = null;
        eas.utcTimezone = null;
        eas.defaultTimezoneInfo = null;
        eas.windowsTimezoneMap = {};
        eas.openWindows = {};

        eas.overlayManager = new OverlayManager({verbose: 0});
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/abNewCardDialog.xul", "chrome://eas4tbsync/content/overlays/abNewCardWindow.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/abNewCardDialog.xul", "chrome://eas4tbsync/content/overlays/abCardWindow.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/abEditCardDialog.xul", "chrome://eas4tbsync/content/overlays/abCardWindow.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/addressbook.xul", "chrome://eas4tbsync/content/overlays/addressbookoverlay.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/addressbook.xul", "chrome://eas4tbsync/content/overlays/addressbookdetailsoverlay.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/addressbook.xul", "chrome://eas4tbsync/content/overlays/abServerSearch.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/abContactsPanel.xul", "chrome://eas4tbsync/content/overlays/abServerSearch.xul");

        // The abCSS.xul overlay is just adding a CSS file.
        await eas.overlayManager.registerOverlay("chrome://messenger/content/messengercompose/messengercompose.xul", "chrome://eas4tbsync/content/overlays/abCSS.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/abNewCardDialog.xul", "chrome://eas4tbsync/content/overlays/abCSS.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/addressbook.xul", "chrome://eas4tbsync/content/overlays/abCSS.xul");

        eas.overlayManager.startObserving();
                
        try {
            // Create a basic error info (no accountname or foldername, just the provider)
            let eventLogInfo = new tbSync.EventLogInfo("eas");
            
            if (tbSync.lightning.isAvailable()) {
                
                //get timezone info of default timezone (old cal. without dtz are depricated)
                eas.defaultTimezone = (tbSync.lightning.cal.dtz && tbSync.lightning.cal.dtz.defaultTimezone) ? tbSync.lightning.cal.dtz.defaultTimezone : tbSync.lightning.cal.calendarDefaultTimezone();
                eas.utcTimezone = (tbSync.lightning.cal.dtz && tbSync.lightning.cal.dtz.UTC) ? tbSync.lightning.cal.dtz.UTC : tbSync.lightning.cal.UTC();
                if (eas.defaultTimezone && eas.defaultTimezone.icalComponent) {
                    tbSync.eventlog.add("info", eventLogInfo, "Default timezone has been found.");                    
                } else {
                    tbSync.eventlog.add("info", eventLogInfo, "Default timezone is not defined, using UTC!");
                    eas.defaultTimezone = eas.utcTimezone;
                }

                eas.defaultTimezoneInfo = eas.tools.getTimezoneInfo(eas.defaultTimezone);
                if (!eas.defaultTimezoneInfo) {
                    tbSync.eventlog.add("info", eventLogInfo, "Could not create defaultTimezoneInfo");
                }
                
                //get windows timezone data from CSV
                let csvData = await eas.tools.fetchFile("chrome://eas4tbsync/content/timezonedata/WindowsTimezone.csv");
                for (let i = 0; i<csvData.length; i++) {
                    let lData = csvData[i].split(",");
                    if (lData.length<3) continue;
                    
                    let windowsZoneName = lData[0].toString().trim();
                    let zoneType = lData[1].toString().trim();
                    let ianaZoneName = lData[2].toString().trim();
                    
                    if (zoneType == "001") eas.windowsTimezoneMap[windowsZoneName] = ianaZoneName;
                    if (ianaZoneName == eas.defaultTimezoneInfo.std.id) eas.defaultTimezoneInfo.std.windowsZoneName = windowsZoneName;
                }


                //If an EAS calendar is currently NOT associated with an email identity, try to associate, 
                //but do not change any explicitly set association
                // - A) find email identity and accociate (which sets organizer to that user identity)
                // - B) overwrite default organizer with current best guess
                //TODO: Do this after email accounts changed, not only on restart? 
                let providerData = new tbSync.ProviderData("eas");
                let folders = providerData.getFolders({"selected": true, "type": ["8","13"]});
                for (let folder of folders) {
                    let calendar = tbSync.lightning.cal.getCalendarManager().getCalendarById(folder.getFolderProperty("target"));
                    if (calendar && calendar.getProperty("imip.identity.key") == "") {
                        //is there an email identity for this eas account?
                        let authData = eas.network.getAuthData(folder.accountData);

                        let key = eas.tools.getIdentityKey(authData.user);
                        if (key === "") { //TODO: Do this even after manually switching to NONE, not only on restart?
                            //set transient calendar organizer settings based on current best guess and 
                            calendar.setProperty("organizerId", tbSync.lightning.cal.email.prependMailTo(authData.user));
                            calendar.setProperty("organizerCN",  calendar.getProperty("fallbackOrganizerName"));
                        } else {                      
                            //force switch to found identity
                            calendar.setProperty("imip.identity.key", key);
                        }
                    }
                }
            } else {
                    tbSync.eventlog.add("info", eventLogInfo, "Lightning was not loaded, creation of timezone objects has been skipped.");
            }
        } catch(e) {
            Components.utils.reportError(e);        
        }        
    }



    /**
     * Called during unload of external provider extension to unload provider.
     */
    static async unload() {
        eas.overlayManager.stopObserving();	

        // Close all open windows of this provider.
        for (let id in eas.openWindows) {
          if (eas.openWindows.hasOwnProperty(id)) {
            eas.openWindows[id].close();
          }
        }
    }



    /**
     * Returns nice string for the name of provider for the add account menu.
     */
    static getNiceProviderName() {
        return "Exchange ActiveSync";
    }



    /**
     * Returns location of a provider icon.
     *
     * @param size       [in] size of requested icon
     * @param accountData  [in] optional AccountData
     *
     */
    static getProviderIcon(size, accountData = null) {
        switch (size) {
            case 16:
                return "chrome://eas4tbsync/skin/eas16.png";
            case 32:
                return "chrome://eas4tbsync/skin/eas32.png";
            default :
                return "chrome://eas4tbsync/skin/eas64.png";
        }
    }



    /**
     * Returns a list of sponsors, they will be sorted by the index
     */
    static getSponsors() {
        return {
            "Schiessl, Michael 1" : {name: "Michael Schiessl", description: "Tine 2.0", icon: "", link: "" },
            "Schiessl, Michael 2" : {name: "Michael Schiessl", description: " Exchange 2007", icon: "", link: "" },
            "netcup GmbH" : {name: "netcup GmbH", description : "SOGo", icon: "chrome://eas4tbsync/skin/sponsors/netcup.png", link: "http://www.netcup.de/" },
            "nethinks GmbH" : {name: "nethinks GmbH", description : "Zarafa", icon: "chrome://eas4tbsync/skin/sponsors/nethinks.png", link: "http://www.nethinks.com/" },
            "Jau, Stephan" : {name: "Stephan Jau", description: "Horde", icon: "", link: "" },
            "Zavar " : {name: "Zavar", description: "Zoho", icon: "", link: "" },
        };
    }



    /**
     * Returns the email address of the maintainer (used for bug reports).
     */
    static getMaintainerEmail() {
        return "john.bieling@gmx.de";
    }



    /**
     * Returns the URL of the string bundle file of this provider, it can be
     * accessed by tbSync.getString(<key>, <provider>)
     */
    static getStringBundleUrl() {
        return "chrome://eas4tbsync/locale/eas.strings";
    }

    

    /**
     * Returns URL of the new account window.
     *
     * The URL will be opened via openDialog(), when the user wants to create a
     * new account of this provider.
     */
    static getCreateAccountWindowUrl() {
        return "chrome://eas4tbsync/content/manager/createAccount.xul";
    }



    /**
     * Returns overlay XUL URL of the edit account dialog
     * (chrome://tbsync/content/manager/editAccount.xul)
     *
     * This overlay must (!) implement:
     *
     *    tbSyncEditAccountOverlay.onload(window, accountData)
     *
     * which is called each time an account of this provider is viewed/selected
     * in the manager and provides the tbSync.AccountData of the corresponding
     * account.
     */
    static getEditAccountOverlayUrl() {
        return "chrome://eas4tbsync/content/manager/editAccountOverlay.xul";
    }



    /**
     * Return object which contains all possible fields of a row in the
     * accounts database with the default value if not yet stored in the 
     * database.
     */
    static getDefaultAccountEntries() {
        let row = {
            "policykey" : "0", 
            "foldersynckey" : "0",
            "deviceId" : eas.tools.getNewDeviceId(),
            "asversionselected" : "auto",
            "asversion" : "",
            "host" : "",
            "user" : "",
            "servertype" : "",
            "seperator" : "10",
            "https" : true,
            "provision" : false,
            "displayoverride" : false, 
            "lastEasOptionsUpdate":"0",
            "allowedEasVersions": "",
            "allowedEasCommands": "",
            "useragent": eas.prefs.getCharPref("clientID.useragent"),
            "devicetype": eas.prefs.getCharPref("clientID.type"),
            "galautocomplete": true, 
            "synclimit" : "7",
            }; 
        return row;
    }



    /**
     * Return object which contains all possible fields of a row in the folder 
     * database with the default value if not yet stored in the database.
     */
    static getDefaultFolderEntries() {
        let folder = {
            "type" : "",
            "synckey" : "",
            "targetColor" : "",
            "parentID" : "0",
            "serverID" : "",
            };
        return folder;
    }



    /**
     * Is called everytime an account of this provider is enabled in the
     * manager UI.
     *
     * @param accountData  [in] AccountData
     */
    static onEnableAccount(accountData) {
        accountData.resetAccountProperty("policykey");
        accountData.resetAccountProperty("foldersynckey");
        accountData.resetAccountProperty("lastEasOptionsUpdate");
        accountData.resetAccountProperty("lastsynctime");
    }



    /**
     * Is called everytime an account of this provider is disabled in the
     * manager UI.
     *
     * @param accountData  [in] AccountData
     */
    static onDisableAccount(accountData) {
    }



    /**
     * Is called everytime an new target is created, intended to set a clean
     * sync status.
     *
     * @param accountData  [in] FolderData
     */
    static onResetTarget(folderData) {
        folderData.resetFolderProperty("synckey");
        folderData.resetFolderProperty("lastsynctime");
    }



    /**
     * Implement this method, if this provider should add additional entries
     * to the autocomplete list while typing something into the address field
     * of the message composer.
     *
     * When creating directories, you can set:
     *
     *    directory.setBoolValue("enable_autocomplete", false);
     *
     * to disable the default autocomplete for this directory and have full
     * control over the autocomplete.
     *
     * @param accountData   [in] AccountData of the account which should be
     *                           searched
     * @param currentQuery  [in] search query
     *
     * Return arrary of AutoCompleteData entries.
     */
    static async abAutoComplete(accountData, currentQuery)  {
        let data = [];

        if (currentQuery.length < 3)
            return null;
        
        if (!accountData.getAccountProperty("allowedEasCommands").split(",").includes("Search")) {
            return null;
        }

        if (!accountData.getAccountProperty("galautocomplete")) {
            return null;
        }
            
        let response = await eas.network.getSearchResults(accountData, currentQuery);
        let wbxmlData = eas.network.getDataFromResponse(response);

        if (wbxmlData.Search && wbxmlData.Search.Response && wbxmlData.Search.Response.Store && wbxmlData.Search.Response.Store.Result) {
            let results = eas.xmltools.nodeAsArray(wbxmlData.Search.Response.Store.Result);
            let accountname = accountData.getAccountProperty("accountname");
        
            for (let count = 0; count < results.length; count++) {
                if (results[count].Properties) {
                    //tbSync.window.console.log('Found contact:' + results[count].Properties.DisplayName);
                    data.push({
                        value: results[count].Properties.DisplayName + " <" + results[count].Properties.EmailAddress + ">", 
                        comment: tbSync.getString("autocomplete.serverdirectory", "eas") + " ("+accountData.getAccountProperty("accountname")+")",
                        icon: eas.Base.getProviderIcon(16, accountData),
                        style: "EASGAL",
                    });
                }
            }
        }
        
        return data;
    }



    /**
     * Returns all folders of the account, sorted in the desired order.
     * The most simple implementation is to return accountData.getAllFolders();
     *
     * @param accountData         [in] AccountData for the account for which the 
     *                                 sorted folder should be returned
     */
    static getSortedFolders(accountData) {
        let allowedTypesOrder = ["9","14","8","13","7","15"];
        
        function getIdChain (aServerID) {
            let serverID = aServerID;
            let chain = [];
            let folder;
            let rootType = "";
            
            // create sort string so that child folders are directly below their parent folders
            do { 
                folder = accountData.getFolder("serverID", serverID);
                if (folder) {
                    chain.unshift(folder.getFolderProperty("foldername"));
                    serverID = folder.getFolderProperty("parentID");
                    rootType = folder.getFolderProperty("type");
                }
            } while (folder && serverID != "0")
            
            // different folder types are grouped and trashed folders at the end
            let pos = allowedTypesOrder.indexOf(rootType);
            chain.unshift(pos == -1 ? "ZZZ" : pos.toString().padStart(3,"0"));
                        
            return chain.join(".");
        };
        
        let toBeSorted = [];
        let folders = accountData.getAllFolders();
        for (let f of folders) {
            if (!allowedTypesOrder.includes(f.getFolderProperty("type"))) {
                continue;
            }
            toBeSorted.push({"key": getIdChain(f.getFolderProperty("serverID")), "folder": f});
        }
        
        //sort
        toBeSorted.sort(function(a,b) {
            return  a.key > b.key;
        });

        let sortedFolders = [];
        for (let sortObj of toBeSorted) {
            sortedFolders.push(sortObj.folder);
        }
        return sortedFolders;        
    }



    /**
     * Return the connection timeout for an active sync, so TbSync can append
     * a countdown to the connection timeout, while waiting for an answer from
     * the server. Only syncstates which start with "send." will trigger this.
     *
     * @param accountData      [in] AccountData
     *
     * return timeout in milliseconds
     */
    static getConnectionTimeout(accountData) {
        return eas.prefs.getIntPref("timeout");
    }
    


    /**
     * Is called if TbSync needs to synchronize the folder list.
     *
     * @param syncData      [in] SyncData
     * @param syncJob       [in] String with a specific sync job. Defaults to
     *                           "sync", but can be set via the syncDescription
     *                           of AccountData.sync() or FolderData.sync()
     * @param syncRunNr     [in] Indicates the n-th number the account is being synced.
     *                           It starts with 1 and is limited by 
     *                           syncDescription.maxAccountReruns. 
     *
     * !!! NEVER CALL THIS FUNCTION DIRECTLY BUT USE !!!
     *    tbSync.AccountData::sync()
     *
     * return StatusData
     */
    static async syncFolderList(syncData, syncJob, syncRunNr) {
        // Recommendation: Put the actual function call inside a try catch, to
        // ensure returning a proper StatusData object, regardless of what
        // happens inside that function. You may also throw custom errors
        // in that function, which have the StatusData obj attached, which
        // should be returned.
        
        try {
            await eas.sync.folderList(syncData);
        } catch (e) {
            if (e.name == "eas4tbsync") {
                return e.statusData;
            } else {
                Components.utils.reportError(e);
                // re-throw any other error and let TbSync handle it
                throw (e);
            }
        }

        // Fall through, if there was no error.
        return new tbSync.StatusData();        
    }
    


    /**
     * Is called if TbSync needs to synchronize a folder.
     *
     * @param syncData      [in] SyncData
     * @param syncJob       [in] String with a specific sync job. Defaults to
     *                           "sync", but can be set via the syncDescription
     *                           of AccountData.sync() or FolderData.sync()
     * @param syncRunNr     [in] Indicates the n-th number the folder is being synced.
     *                           It starts with 1 and is limited by 
     *                           syncDescription.maxFolderReruns.
     *
     * !!! NEVER CALL THIS FUNCTION DIRECTLY BUT USE !!!
     *    tbSync.AccountData::sync() or
     *    tbSync.FolderData::sync()
     *
     * return StatusData
     */
    static async syncFolder(syncData, syncJob, syncRunNr) {
        // Recommendation: Put the actual function call inside a try catch, to
        // ensure returning a proper StatusData object, regardless of what
        // happens inside that function. You may also throw custom errors
        // in that function, which have the StatusData obj attached, which
        // should be returned.
        
        try {
            switch (syncJob) {
                case "deletefolder":
                    await eas.sync.deleteFolder(syncData);
                    break;
                default:
                   await eas.sync.singleFolder(syncData);
            }
        } catch (e) {
            if (e.name == "eas4tbsync") {
                return e.statusData;
            } else {
                Components.utils.reportError(e);
                // re-throw any other error and let TbSync handle it
                throw (e);
            }
        }

        // Fall through, if there was no error.
        return new tbSync.StatusData();   
    }
}





var standardTargets = {
    // This provider is using the standard "addressbook" targetType, so it must
    // implement the addressbook object.
    addressbook : {
        
        // define an item property, which should be used for the changelog
        // basically your primary key for an abItem
        // UID will be used, if nothing specified
        primaryKeyField: "X-EAS-SERVERID",
        


        generatePrimaryKey: function (folderData) {
             return tbSync.generateUUID();
        },
        


        // enable or disable changelog
        logUserChanges: true,



        directoryObserver: function (aTopic, folderData) {
            switch (aTopic) {
                case "addrbook-removed":
                case "addrbook-updated":
                    //Services.console.logStringMessage("["+ aTopic + "] " + folderData.getFolderProperty("foldername"));
                    break;
            }
        },
        


        cardObserver: function (aTopic, folderData, abCardItem) {
            switch (aTopic) {
                case "addrbook-contact-updated":
                case "addrbook-contact-removed":
                    //Services.console.logStringMessage("["+ aTopic + "] " + abCardItem.getProperty("DisplayName"));
                    break;

                case "addrbook-contact-created":
                {
                    //Services.console.logStringMessage("["+ aTopic + "] "+ abCardItem.getProperty("DisplayName")+">");
                    break;
                }
            }
        },
        


        listObserver: function (aTopic, folderData, abListItem, abListMember) {
            switch (aTopic) {
                case "addrbook-list-member-added":
                case "addrbook-list-member-removed":
                    //Services.console.logStringMessage("["+ aTopic + "] MemberName: " + abListMember.getProperty("DisplayName"));
                    break;
                
                case "addrbook-list-removed":
                case "addrbook-list-updated":
                    //Services.console.logStringMessage("["+ aTopic + "] ListName: " + abListItem.getProperty("ListName"));
                    break;
                
                case "addrbook-list-created": 
                    //Services.console.logStringMessage("["+ aTopic + "] ListName: "+abListItem.getProperty("ListName")+">");
                    break;
            }
        },
        


        /**
         * Is called by TargetData::getTarget() if  the standard "addressbook"
         * targetType is used, and a new addressbook needs to be created.
         *
         * @param newname       [in] name of the new address book
         * @param folderData  [in] FolderData
         *
         * return the new directory
         */
        createAddressBook: function (newname, folderData) {
            let dirPrefId = MailServices.ab.newAddressBook(newname, "", 2);  /* kPABDirectory - return abManager.newAddressBook(name, "moz-abmdbdirectory://", 2); */
            let directory = MailServices.ab.getDirectoryFromId(dirPrefId);

            if (directory && directory instanceof Components.interfaces.nsIAbDirectory && directory.dirPrefId == dirPrefId) {
                directory.setStringValue("tbSyncIcon", "eas");
                return directory;		
            }
            return null;
        },    
    },



    // This provider is using the standard "calendar" targetType, so it must
    // implement the calendar object.
    calendar : {
        
        // The calendar target does not support a custom primaryKeyField, because
        // the lightning implementation only allows to search for items via UID.
        // Like the addressbook target, the calendar target item element has a
        // primaryKey getter/setter which - however - only works on the UID.
        
        // enable or disable changelog
        logUserChanges: true,
        


        calendarObserver: function (aTopic, folderData, tbCalendar, aPropertyName, aPropertyValue, aOldPropertyValue) {
            switch (aTopic) {
                case "onCalendarPropertyChanged":
                    //Services.console.logStringMessage("["+ aTopic + "] " + tbCalendar.calendar.name + " : " + aPropertyName);
                    break;
                
                case "onCalendarDeleted":
                case "onCalendarPropertyDeleted":
                    //Services.console.logStringMessage("["+ aTopic + "] " +tbCalendar.calendar.name);
                    break;
            }
        },
        


        itemObserver: function (aTopic, folderData, tbItem, tbOldItem) {
            switch (aTopic) {
                case "onAddItem":
                case "onModifyItem":
                case "onDeleteItem":
                    //Services.console.logStringMessage("["+ aTopic + "] " + tbItem.nativeItem.title);
                    break;
            }
        },



        /**
         * Is called by TargetData::getTarget() if  the standard "calendar" targetType is used, and a new calendar needs to be created.
         *
         * @param newname       [in] name of the new calendar
         * @param folderData  [in] folderData
         *
         * return the new calendar
         */
        createCalendar: function(newname, folderData) {
            let calManager = tbSync.lightning.cal.getCalendarManager();
            //Alternative calendar, which uses calTbSyncCalendar
            //let newCalendar = calManager.createCalendar("TbSync", Services.io.newURI('tbsync-calendar://'));

            //Create the new standard calendar with a unique name
            let newCalendar = calManager.createCalendar("storage", Services.io.newURI("moz-storage-calendar://"));
            newCalendar.id = tbSync.lightning.cal.getUUID();
            newCalendar.name = newname;

            newCalendar.setProperty("color", folderData.getFolderProperty("targetColor"));
            newCalendar.setProperty("relaxedMode", true); //sometimes we get "generation too old for modifyItem", check can be disabled with relaxedMode
            newCalendar.setProperty("calendar-main-in-composite",true);
            newCalendar.setProperty("readOnly", folderData.getFolderProperty("downloadonly") == "1");
            calManager.registerCalendar(newCalendar);

            let authData = eas.network.getAuthData(folderData.accountData);
            
            //is there an email identity we can associate this calendar to? 
            //getIdentityKey returns "" if none found, which removes any association
            let key = eas.tools.getIdentityKey(authData.user);
            newCalendar.setProperty("fallbackOrganizerName", newCalendar.getProperty("organizerCN"));
            newCalendar.setProperty("imip.identity.key", key);
            if (key === "") {
                //there is no matching email identity - use current default value as best guess and remove association
                //use current best guess 
                newCalendar.setProperty("organizerCN", newCalendar.getProperty("fallbackOrganizerName"));
                newCalendar.setProperty("organizerId", tbSync.lightning.cal.email.prependMailTo(authData.user));
            }
            
            return newCalendar;
        },
    }
}





/**
 * This provider is using the StandardFolderList (instead of this it could also
 * implement the full folderList object).
 *
 * The DOM of the folderlist can be accessed by
 * 
 *    let list = document.getElementById("tbsync.accountsettings.folderlist");
 * 
 * and the folderData of each entry is attached to each row:
 * 
 *    let folderData = folderList.selectedItem.folderData;
 *
 */
var StandardFolderList = class {
    /**
     * Is called before the context menu of the folderlist is shown, allows to
     * show/hide custom menu options based on selected folder. During an active
     * sync, folderData will be null.
     *
     * @param window        [in] window object of the account settings window
     * @param folderData    [in] FolderData of the selected folder
     */
    static onContextMenuShowing(window, folderData) {
        let hideContextMenuDelete = true;
        if (folderData !== null) {
            //if a folder in trash is selected, also show ContextMenuDelete (but only if FolderDelete is allowed)
            if (eas.tools.parentIsTrash(folderData) && folderData.accountData.getAccountProperty("allowedEasCommands").split(",").includes("FolderDelete")) {
                hideContextMenuDelete = false;
                window.document.getElementById("TbSync.eas.FolderListContextMenuDelete").label = tbSync.getString("deletefolder.menuentry::" + folderData.getFolderProperty("foldername"), "eas");
            }                
        }
        window.document.getElementById("TbSync.eas.FolderListContextMenuDelete").hidden = hideContextMenuDelete;
    }



    /**
     * Return the icon used in the folderlist to represent the different folder
     * types.
     *
     * @param folderData         [in] FolderData of the selected folder
     */
    static getTypeImage(folderData) {
        let src = "";
        switch (folderData.getFolderProperty("type")) {
            case "9": 
            case "14": 
                src = "contacts16.png";
                break;
            case "8":
            case "13":
                src = "calendar16.png";
                break;
            case "7":
            case "15":
                src = "todo16.png";
                break;
        }
        return "chrome://tbsync/skin/" + src;
    }
    


    /**
     * Return the name of the folder shown in the folderlist.
     *
     * @param folderData         [in] FolderData of the selected folder
     */ 
    static getFolderDisplayName(folderData) {
        let folderName = folderData.getFolderProperty("foldername");
        if (eas.tools.parentIsTrash(folderData)) folderName = tbSync.getString("recyclebin", "eas") + " | " + folderName;
        return folderName;
    }
    


    /**
     * Return the attributes for the ACL RO (readonly menu element per folder.
    * (label, disabled, hidden, style, ...)
     *
     * @param folderData         [in] FolderData of the selected folder
     *
     * Return a list of attributes and their values If both (RO+RW) do
     * not return any attributes, the ACL menu is not displayed at all.
     */ 
    static getAttributesRoAcl(folderData) {
        return {
            label: tbSync.getString("acl.readonly", "eas"),
        };
    }
    


    /**
     * Return the attributes for the ACL RW (readwrite) menu element per folder.
    * (label, disabled, hidden, style, ...)
     *
     * @param folderData         [in] FolderData of the selected folder
     *
     * Return a list of attributes and their values. If both (RO+RW) do
     * not return any attributes, the ACL menu is not displayed at all.
     */ 
    static getAttributesRwAcl(folderData) {
        return {
            label: tbSync.getString("acl.readwrite", "eas"),
        }             
    }
}

Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/includes/network.js", this, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/includes/wbxmltools.js", this, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/includes/xmltools.js", this, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/includes/tools.js", this, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/includes/sync.js", this, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/includes/contactsync.js", this.sync, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/includes/calendarsync.js", this.sync, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/includes/tasksync.js", this.sync, "UTF-8");
