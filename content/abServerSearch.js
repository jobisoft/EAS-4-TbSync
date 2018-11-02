/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

Components.utils.import("resource://gre/modules/Task.jsm");

//this is used in multiple places (addressbook + contactsidebar) so we cannot use objects of tbSync to store states, but need to store distinct variables
//in the window scope -> window.tbSync_XY

var serverSearch = {};

serverSearch.eventHandlerWindowReference = function (window) {
	this.window = window;
    
    this.removeEventListener = function (element, type, bubble) {
        element.removeEventListener(type, this, bubble);
    };

    this.addEventListener = function (element, type, bubble) {
        element.addEventListener(type, this, bubble);
    };
    
    this.handleEvent = function(event) {
        switch(event.type) {
            case 'input':
                serverSearch.onSearchInputChanged(this.window);
                break;
            case "select":
                {
                    serverSearch.clearServerSearchResults(this.window);
                    let searchbox =  window.document.getElementById("peopleSearchInput");
                    let target = window.GetSelectedDirectory();
                    if (searchbox && target) {
                        let folders = tbSync.db.findFoldersWithSetting("target", target);
                        if (folders.length == 1 && tbSync[tbSync.db.getAccountSetting(folders[0].account, "provider")].abServerSearch) {
                            searchbox.setAttribute("placeholder", tbSync.getLocalizedMessage("addressbook.searchgal::" + tbSync.db.getAccountSetting(folders[0].account, "accountname")));
                        } else {
                            searchbox.setAttribute("placeholder", tbSync.getLocalizedMessage((target == "moz-abdirectory://?") ? "addressbook.searchall" : "addressbook.searchthis"));
                        }
                    }
                }
                break;
        }
    };
    return this;
}

serverSearch.searchValuePoll = function (window, searchbox) {
    let value = searchbox.value;
    if (window.tbSync_searchValue != "" && value == "") {
        serverSearch.clearServerSearchResults(window);
    }
    window.tbSync_searchValue = value;
}

serverSearch.onInjectIntoAddressbook = function (window) {
    window.tbSync_eventHandler = serverSearch.eventHandlerWindowReference(window);
	
    let searchbox =  window.document.getElementById("peopleSearchInput");
    if (searchbox) {
        window.tbSync_searchValue = searchbox.value;
        window.tbSync_searchValuePollHandler = window.setInterval(function(){serverSearch.searchValuePoll(window, searchbox)}, 200);
        window.tbSync_eventHandler.addEventListener(searchbox, "input", false);
    }
    
    let dirtree = window.document.getElementById("dirTree");
    if (dirtree) {
        window.tbSync_eventHandler.addEventListener(dirtree, "select", false);        
    }
}

serverSearch.onRemoveFromAddressbook = function (window) {
    let searchbox =  window.document.getElementById("peopleSearchInput");
    if (searchbox) {
        window.tbSync_eventHandler.removeEventListener(searchbox, "input", false);
        window.clearInterval(window.tbSync_searchValuePollHandler);
    }

    let dirtree = window.document.getElementById("dirTree");
    if (dirtree) {
        window.tbSync_eventHandler.removeEventListener(dirtree, "select", false);        
    }
}

serverSearch.clearServerSearchResults = function (window) {
    let target = window.GetSelectedDirectory();
    if (target == "moz-abdirectory://?") return; //global search not yet(?) supported
    
    let addressbook = tbSync.getAddressBookObject(target);
    if (addressbook) {
        try {
            let oldresults = addressbook.getCardsFromProperty("X-Server-Searchresult", "TbSync", true);
            let cardsToDelete = Components.classes["@mozilla.org/array;1"].createInstance(Components.interfaces.nsIMutableArray);
            while (oldresults.hasMoreElements()) {
                cardsToDelete.appendElement(oldresults.getNext(), false);
            }
            addressbook.deleteCards(cardsToDelete);
        } catch (e) {
            //if  getCardsFromProperty is not implemented, do nothing
        }
    }
}

serverSearch.onSearchInputChanged = Task.async (function* (window) {
    let target = window.GetSelectedDirectory();
    if (target == "moz-abdirectory://?") return; //global search not yet(?) supported
	
    let folders = tbSync.db.findFoldersWithSetting("target", target);
    if (folders.length == 1) {
        let searchbox = window.document.getElementById("peopleSearchInput");
        let query = searchbox.value;        
        let addressbook = tbSync.getAddressBookObject(target);

        let account = folders[0].account;
        let provider = tbSync.db.getAccountSetting(account, "provider");
        let accountname = tbSync.db.getAccountSetting(account, "accountname");
        if (tbSync[provider].abServerSearch) {

            if (query.length<3) {
                //delete all old results
                serverSearch.clearServerSearchResults(window);
                window.onEnterInSearchBar();
            } else {
                window.tbSync_serverSearchNextQuery = query;                
                if (window.tbSync_serverSearchBusy) {
                } else {
                    window.tbSync_serverSearchBusy = true;
                    while (window.tbSync_serverSearchBusy) {

                        yield tbSync.sleep(1000);
                        let currentQuery = window.tbSync_serverSearchNextQuery;
                        window.tbSync_serverSearchNextQuery = "";
                        let results = yield tbSync[provider].abServerSearch (account, currentQuery);

                        //delete all old results
                        serverSearch.clearServerSearchResults(window);

                        for (let count = 0; count < results.length; count++) {
                            let newItem = Components.classes["@mozilla.org/addressbook/cardproperty;1"].createInstance(Components.interfaces.nsIAbCard);
                            for (var prop in results[count].properties) {
                                if (results[count].properties.hasOwnProperty(prop)) {
                                    newItem.setProperty(prop, results[count].properties[prop]);
                                }
                            }
                            newItem.setProperty("X-Server-Searchresult", "TbSync");
                            addressbook.addCard(newItem);
                        }   
                        window.onEnterInSearchBar();
                        if (window.tbSync_serverSearchNextQuery == "") window.tbSync_serverSearchBusy = false;
                    }
                }
            }            
        }
    }
})
