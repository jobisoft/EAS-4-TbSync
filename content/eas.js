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
    syncNextFolder: "syncNextFolder",
    resyncFolder: "resyncFolder", //will take down target and do a fresh sync
    resyncAccount: "resyncAccount", //will loop once more, but will not do any special actions
    abortWithError: "abortWithError",
    abortWithServerError: "abortWithServerError",
});

eas.windowsTimezoneMap = {};
eas.cachedTimezoneData = null;
eas.defaultTimezoneInfo = null;
eas.defaultTimezone = null;
eas.utcTimezone = null;


/**
 * Implementation the TbSync interfaces for external provider extensions.
 */    
    
/* TODO: 
 - convert account properties to native types (int, bool)
- getAttributesRoAcl
- getAttributesRwAcl
- getSortedFolders
            if (tbSync.lightning.isAvailable()) {

*/


var base = {
    /**
     * Called during load of external provider extension to init provider.
     */
    load: async function () {
        eas.overlayManager = new OverlayManager({verbose: 0});
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/abNewCardDialog.xul", "chrome://eas4tbsync/content/overlays/abNewCardWindow.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/abNewCardDialog.xul", "chrome://eas4tbsync/content/overlays/abCardWindow.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/abEditCardDialog.xul", "chrome://eas4tbsync/content/overlays/abCardWindow.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/addressbook.xul", "chrome://eas4tbsync/content/overlays/addressbookoverlay.xul");
        await eas.overlayManager.registerOverlay("chrome://messenger/content/addressbook/addressbook.xul", "chrome://eas4tbsync/content/overlays/addressbookdetailsoverlay.xul");
        eas.overlayManager.startObserving();

        eas.openWindows = {};
        
        // Create a basic error info (no accountname or foldername, just the provider)
        let errorInfo = new tbSync.ErrorInfo("eas");
        
        try {
            if (tbSync.lightning.isAvailable() && 1==2) {
                
                //get timezone info of default timezone (old cal. without dtz are depricated)
                eas.defaultTimezone = (cal.dtz && cal.dtz.defaultTimezone) ? cal.dtz.defaultTimezone : cal.calendarDefaultTimezone();
                eas.utcTimezone = (cal.dtz && cal.dtz.UTC) ? cal.dtz.UTC : cal.UTC();
                if (eas.defaultTimezone && eas.defaultTimezone.icalComponent) {
                    tbSync.errorlog.add("info", errorInfo, "Default timezone has been found.");                    
                } else {
                    tbSync.errorlog.add("info", errorInfo, "Default timezone is not defined, using UTC!");
                    eas.defaultTimezone = eas.utcTimezone;
                }

                eas.defaultTimezoneInfo = eas.tools.getTimezoneInfo(eas.defaultTimezone);
                if (!eas.defaultTimezoneInfo) {
                    tbSync.errorlog.add("info", errorInfo, "Could not create defaultTimezoneInfo");
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
                let folders = tbSync.db.findFoldersWithSetting(["selected","type"], ["1","8,13"], "provider", "eas");
                for (let f=0; f < folders.length; f++) {
                    let calendar = cal.getCalendarManager().getCalendarById(folders[f].target);
                    if (calendar && calendar.getProperty("imip.identity.key") == "") {
                        //is there an email identity for this eas account?
                        let key = tbSync.getIdentityKey(tbSync.db.getAccountSetting(folders[f].account, "user"));
                        if (key === "") { //TODO: Do this even after manually switching to NONE, not only on restart?
                            //set transient calendar organizer settings based on current best guess and 
                            calendar.setProperty("organizerId", cal.email.prependMailTo(tbSync.db.getAccountSetting(folders[f].account, "user")));
                            calendar.setProperty("organizerCN",  calendar.getProperty("fallbackOrganizerName"));
                        } else {                      
                            //force switch to found identity
                            calendar.setProperty("imip.identity.key", key);
                        }
                    }
                }
            } else {
                    tbSync.errorlog.add("info", errorInfo, "Lightning was not loaded, creation of timezone objects has been skipped.");
            }
        } catch(e) {
                    Components.utils.reportError(e);        
        }        
    },



    /**
     * Called during unload of external provider extension to unload provider.
     */
    unload: async function () {
        eas.overlayManager.stopObserving();	

        // Close all open windows of this provider.
        for (let id in eas.openWindows) {
          if (eas.openWindows.hasOwnProperty(id)) {
            eas.openWindows[id].close();
          }
        }
    },





    /**
     * Returns nice string for the name of provider for the add account menu.
     */
    getNiceProviderName: function () {
        return "Exchange ActiveSync";
    },


    /**
     * Returns location of a provider icon.
     *
     * @param size       [in] size of requested icon
     * @param accountData  [in] optional AccountData
     *
     */
    getProviderIcon: function (size, accountData = null) {
        switch (size) {
            case 16:
                return "chrome://eas4tbsync/skin/eas16.png";
            case 32:
                return "chrome://eas4tbsync/skin/eas32.png";
            default :
                return "chrome://eas4tbsync/skin/eas64.png";
        }
    },



    /**
     * Returns a list of sponsors, they will be sorted by the index
     */
    getSponsors: function () {
        return {
            "Schiessl, Michael 1" : {name: "Michael Schiessl", description: "Tine 2.0", icon: "", link: "" },
            "Schiessl, Michael 2" : {name: "Michael Schiessl", description: " Exchange 2007", icon: "", link: "" },
            "netcup GmbH" : {name: "netcup GmbH", description : "SOGo", icon: "chrome://eas4tbsync/skin/sponsors/netcup.png", link: "http://www.netcup.de/" },
            "nethinks GmbH" : {name: "nethinks GmbH", description : "Zarafa", icon: "chrome://eas4tbsync/skin/sponsors/nethinks.png", link: "http://www.nethinks.com/" },
            "Jau, Stephan" : {name: "Stephan Jau", description: "Horde", icon: "", link: "" },
            "Zavar " : {name: "Zavar", description: "Zoho", icon: "", link: "" },
        };
    },



    /**
     * Returns the email address of the maintainer (used for bug reports).
     */
    getMaintainerEmail: function () {
        return "john.bieling@gmx.de";
    },


    /**
     * Returns the URL of the string bundle file of this provider, it can be
     * accessed by tbSync.getString(<key>, <provider>)
     */
    getStringBundleUrl: function () {
        return "chrome://eas4tbsync/locale/eas.strings";
    },

    
    /**
     * Returns URL of the new account window.
     *
     * The URL will be opened via openDialog() and the tbSync.ProviderData of this
     * provider will be passed as first argument. It can be accessed via:
     *
     *    providerData = window.arguments[0];
     */
    getCreateAccountWindowUrl: function () {
        return "chrome://eas4tbsync/content/manager/createAccount.xul";
    },


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
    getEditAccountOverlayUrl: function () {
        return "chrome://eas4tbsync/content/manager/editAccountOverlay.xul";
    },



    /**
     * Return object which contains all possible fields of a row in the
     * accounts database with the default value if not yet stored in the 
     * database.
     */
    getDefaultAccountEntries: function () {
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
            "https" : "1",
            "provision" : "0",
            "birthday" : "0",
            "displayoverride" : "0", 
            "horde" : "0",
            "lastEasOptionsUpdate":"0",
            "allowedEasVersions": "",
            "allowedEasCommands": "",
            "useragent": eas.prefs.getCharPref("clientID.useragent"),
            "devicetype": eas.prefs.getCharPref("clientID.type"),
            "galautocomplete": "1", 
            }; 
        return row;
    },


    /**
     * Return object which contains all possible fields of a row in the folder 
     * database with the default value if not yet stored in the database.
     */
    getDefaultFolderEntries: function () { //TODO: shadow more standard entries
        let folder = {
            //"folderID" : "",
            //"useChangeLog" : "1", //log changes into changelog
            "type" : "",
            "synckey" : "",
            "targetColor" : "",
            "parentID" : "",
            };
        return folder;
    },



    /**
     * Is called everytime an account of this provider is enabled in the
     * manager UI.
     *
     * @param accountData  [in] AccountData
     */
    onEnableAccount: function (accountData) {
        accountData.resetAccountProperty("policykey");
        accountData.resetAccountProperty("foldersynckey");
        accountData.resetAccountProperty("lastEasOptionsUpdate");
        accountData.resetAccountProperty("lastsynctime");
    },



    /**
     * Is called everytime an account of this provider is disabled in the
     * manager UI.
     *
     * @param accountData  [in] AccountData
     */
    onDisableAccount: function (accountData) {
    },



    /**
     * Is called everytime an new target is created, intended to set a clean
     * sync status.
     *
     * @param accountData  [in] FolderData
     */
    onResetTarget: function (folderData) {
        folderData.resetFolderProperty("synckey");
        folderData.resetFolderProperty("lastsynctime");
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
     * @param accountID       [in] id of the account which should be searched
     * @param currentQuery  [in] search query
     * @param caller        [in] "autocomplete" or "search" //TODO
     */
    abServerSearch: async function (accountID, currentQuery, caller)  {
        return null;
    },



    /**
     * Returns all folders of the account, sorted in the desired order.
     * The most simple implementation is to return accountData.getAllFolders();
     *
     * @param accountData         [in] AccountData for the account for which the 
     *                                 sorted folder should be returned
     */
    getSortedFolders: function (accountData) {
        let folders = accountData.getAllFolders();

/*            let folderData = [];
            let folders = tbSync.db.getFolders(account);
            let allowedTypesOrder = ["9","14","8","13","7","15"];
            let folderIDs = Object.keys(folders).filter(f => allowedTypesOrder.includes(folders[f].type)).sort((a, b) => (tbSync.eas.folderList.getIdChain(allowedTypesOrder, account, a).localeCompare(tbSync.eas.folderList.getIdChain(allowedTypesOrder, account, b))));
            
            for (let i=0; i < folderIDs.length; i++) {
                folderData.push(tbSync.eas.folderList.getRowData(folders[folderIDs[i]]));
            }
            return folderData; */
        
        return folders;
    },


    /**
     * Return the connection timeout for an active sync, so TbSync can append
     * a countdown to the connection timeout, while waiting for an answer from
     * the server. Only syncstates which start with "send." will trigger this.
     *
     * @param syncData      [in] SyncData
     *
     * return timeout in milliseconds
     */
    getConnectionTimeout: function (syncData) {
        return eas.prefs.getIntPref("timeout");
    },
    
    /**
     * Is called if TbSync needs to synchronize the folder list.
     *
     * @param syncData      [in] SyncData
     *
     * return StatusData
     */
    syncFolderList: async function (syncData) {
        // update folders avail on server and handle added, removed and renamed
        // folders

        try {
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
                        throw eas.sync.finishSync("InvalidServerOptions", eas.flags.abortWithError);
                    } else {
                        throw eas.sync.finishSync("nosupportedeasversion::"+allowedVersionsArray.join(", "), eas.flags.abortWithError);
                    }
                } else if (allowedVersionsString != "" && !allowedVersionsArray.includes(asversionselected)) {
                    throw eas.sync.finishSync("notsupportedeasversion::"+asversionselected+"::"+allowedVersionsArray.join(", "), eas.flags.abortWithError);
                } else {
                    //just use the value set by the user
                    syncData.accountData.setAccountProperty("asversion", asversionselected);
                }
            }
            
            //do we need to get a new policy key?
            if (syncData.accountData.getAccountProperty("provision") == "1" && syncData.accountData.getAccountProperty("policykey") == "0") {
                await eas.network.getPolicykey(syncData);
            }
            
        } catch (e) {
            if (e.name == "eas4tbsync") {
                return e.statusData;
            } else {
                Components.utils.reportError(e);
                return new tbSync.StatusData(tbSync.StatusData.WARNING, "JavaScriptError", e.message + "\n\n" + e.stack);
            }
        }
        // we fall through, if there was no error
        return new tbSync.StatusData();

    },
    
    /**
     * Is called if TbSync needs to synchronize a folder.
     *
     * @param syncData      [in] SyncData
     *
     * return StatusData
     */
    syncFolder: async function (syncData) {
        //process a single folder
        return new tbSync.StatusData(tbSync.StatusData.SUCCESS, "Haha");
        //return await eas.sync.folder(syncData);
    },    
}

// This provider is using the standard "addressbook" targetType, so it must
// implement the addressbook object.
var addressbook = {

    // define a card property, which should be used for the changelog
    // basically your primary key for the abItem properties
    // UID will be used, if nothing specified
    primaryKeyField: "X-DAV-HREF",
    
    generatePrimaryKey: function (folderData) {
         return folderData.getFolderProperty("href") + tbSync.generateUUID() + ".vcf";
    },
    
    // enable or disable changelog
    logUserChanges: true,

    directoryObserver: function (aTopic, folderData) {
        switch (aTopic) {
            case "addrbook-removed":
            case "addrbook-updated":
                //Services.console.logStringMessage("["+ aTopic + "] " + folderData.getFolderProperty("name"));
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
                //Services.console.logStringMessage("["+ aTopic + "] Created new X-DAV-UID for Card <"+ abCardItem.getProperty("DisplayName")+">");
                abCardItem.setProperty("X-DAV-UID", tbSync.generateUUID());
                // the card is tagged with "_by_user" so it will not be changed to "_by_server" by the following modify
                abCardItem.abDirectory.modify(abCardItem);
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
                //Services.console.logStringMessage("["+ aTopic + "] Created new X-DAV-UID for List <"+abListItem.getProperty("ListName")+">");
                abListItem.setProperty("X-DAV-UID", tbSync.generateUUID());
                // custom props of lists get updated directly, no need to call .modify()            
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
        // this is the standard target, should it not be created it like this?
        let dirPrefId = MailServices.ab.newAddressBook(newname, "", 2);
        let directory = MailServices.ab.getDirectoryFromId(dirPrefId);

        if (directory && directory instanceof Components.interfaces.nsIAbDirectory && directory.dirPrefId == dirPrefId) {
            let serviceprovider = folderData.accountData.getAccountProperty("serviceprovider");
            let icon = "custom";
            if (eas.sync.serviceproviders.hasOwnProperty(serviceprovider)) {
                icon = eas.sync.serviceproviders[serviceprovider].icon;
            }
            directory.setStringValue("tbSyncIcon", "dav" + icon);
            return directory;
        }
        return null;
    },    
}



// This provider is using the standard "calendar" targetType, so it must
// implement the calendar object.
var calendar = {
    
    // define a card property, which should be used for the changelog
    // basically your primary key for the abItem properties
    // UID will be used, if nothing specified
    //primaryKeyField: "",
    
    // enable or disable changelog
    //logUserChanges: false,

    // The calendarObserver::onCalendarReregistered needs to know, which field
    // of the folder is used to store the full url of a calendar, to be able to
    // find calendars, which could be connected to other accounts.
    calendarUrlField: "url",
    
    calendarObserver: function (aTopic, folderData, aCalendar, aPropertyName, aPropertyValue, aOldPropertyValue) {
        switch (aTopic) {
            case "onCalendarPropertyChanged":
            {
                switch (aPropertyName) {
                    case "color":
                        if (aOldPropertyValue.toString().toUpperCase() != aPropertyValue.toString().toUpperCase()) {
                            //prepare connection data
                            let connection = new eas.network.ConnectionData(folderData);
                            //update color on server
                            eas.network.sendRequest("<d:propertyupdate "+eas.tools.xmlns(["d","apple"])+"><d:set><d:prop><apple:calendar-color>"+(aPropertyValue + "FFFFFFFF").slice(0,9)+"</apple:calendar-color></d:prop></d:set></d:propertyupdate>", folderData.getFolderProperty("href"), "PROPPATCH", connection);
                        }
                        break;
                }
            }
            break;

            case "onCalendarReregistered": 
            {
                folderData.setFolderProperty("selected", true);
                folderData.setFolderProperty("status", tbSync.StatusData.SUCCESS);
                //add target to re-take control
                folderData.setFolderProperty("target", aCalendar.id);
                //update settings window, if open
                Services.obs.notifyObservers(null, "tbsync.observer.manager.updateSyncstate", folderData.accountID);
            }
            break;
            
            case "onCalendarDeleted":
            case "onCalendarPropertyDeleted":
                //Services.console.logStringMessage("["+ aTopic + "] " + aCalendar.name);
                break;
        }
    },
    
    itemObserver: function (aTopic, folderData, aItem, aOldItem) {
        switch (aTopic) {
            case "onAddItem":
            case "onModifyItem":
            case "onDeleteItem":
                //Services.console.logStringMessage("["+ aTopic + "] " + aItem.title);
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
        let authData = eas.network.getAuthData(folderData.accountData);
        let password = authData.password;
        let username =  authData.user;
      
        let caltype = folderData.getFolderProperty("type");

        let baseUrl = "";
        if (caltype != "ics") {
            baseUrl =  "http" + (folderData.accountData.getAccountProperty("https") ? "s" : "") + "://" + folderData.getFolderProperty("fqdn");
        }

        let url = eas.tools.parseUri(baseUrl + folderData.getFolderProperty("href"));        
        folderData.setFolderProperty("url", url.spec);

        //check if that calendar already exists
        let cals = calManager.getCalendars({});
        let newCalendar = null;
        let found = false;
        for (let calendar of calManager.getCalendars({})) {
            if (calendar.uri.spec == url.spec) {
                newCalendar = calendar;
                found = true;
                break;
            }
        }

        if (!found) {
            newCalendar = calManager.createCalendar(caltype, url); //caldav or ics
            newCalendar.id = tbSync.lightning.cal.getUUID();
            newCalendar.name = newname;

            newCalendar.setProperty("username", username);
            newCalendar.setProperty("color", folderData.getFolderProperty("targetColor"));
            newCalendar.setProperty("calendar-main-in-composite", true);
            newCalendar.setProperty("cache.enabled", folderData.accountData.getAccountProperty("useCalendarCache"));
        }
        
        if (folderData.getFolderProperty("downloadonly")) newCalendar.setProperty("readOnly", true);

        // ICS urls do not need a password
        if (caltype != "ics") {
            tbSync.dump("Searching CalDAV authRealm for", url.host);
            let realm = (eas.network.listOfRealms.hasOwnProperty(url.host)) ? eas.network.listOfRealms[url.host] : "";
            if (realm !== "") {
                tbSync.dump("Found CalDAV authRealm",  realm);
                //manually create a lightning style entry in the password manager
                tbSync.passwordManager.updateLoginInfo(url.prePath, realm, /* old */ username, /* new */ username, password);
            }
        }

        if (!found) {
            calManager.registerCalendar(newCalendar);
        }
        return newCalendar;
    },
}


// This provider is using the standardFolderList (instead of this it could also
// implement the full folderList object).
var standardFolderList = {
    /**
     * Is called before the context menu of the folderlist is shown, allows to
     * show/hide custom menu options based on selected folder.
     *
     * @param document       [in] document object of the account settings window
     * @param folderData         [in] FolderData of the selected folder
     */
    onContextMenuShowing: function (document, folderData) {
    },

    /**
     * Return the icon used in the folderlist to represent the different folder
     * types.
     *
     * @param folderData         [in] FolderData of the selected folder
     */
    getTypeImage: function (folderData) {
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
    },
    
    getAttributesRoAcl: function (folderData) {
        return {
            label: tbSync.getString("acl.readonly", "eas"),
        };
    },
    
    getAttributesRwAcl: function (folderData) {
        let acl = parseInt(folderData.getFolderProperty("acl"));
        let acls = [];
        if (acl & 0x2) acls.push(tbSync.getString("acl.modify", "eas"));
        if (acl & 0x4) acls.push(tbSync.getString("acl.add", "eas"));
        if (acl & 0x8) acls.push(tbSync.getString("acl.delete", "eas"));
        if (acls.length == 0)  acls.push(tbSync.getString("acl.none", "eas"));

        return {
            label: tbSync.getString("acl.readwrite::"+acls.join(", "), "eas"),
            disabled: (acl & 0x7) != 0x7,
        }             
    },
}

Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/network.js", this, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/wbxmltools.js", this, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/xmltools.js", this, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/tools.js", this, "UTF-8");
Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/sync.js", this, "UTF-8");
//Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/tasksync.js", this, "UTF-8");
//Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/calendarsync.js", this, "UTF-8");
//Services.scriptloader.loadSubScript("chrome://eas4tbsync/content/contactsync.js", this, "UTF-8");
