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
        let wbxmlData = eas.network.getDataFromResponse(response);
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

                if (accountReSyncs > 3) {
                    throw eas.sync.finishSync("resync-loop", eas.flags.abortWithError);
                }

              
    },



  






    


        

    


    

    


    
    


    
    
    
    /**
     * Functions used by the folderlist in the main account settings tab
     */
    folderList: {


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
    }
};
    
