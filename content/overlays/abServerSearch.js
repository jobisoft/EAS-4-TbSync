/*
 * This file is part of TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

var { Services } = ChromeUtils.import("resource://gre/modules/Services.jsm");
var { MailServices } = ChromeUtils.import("resource:///modules/MailServices.jsm");
var { tbSync } = ChromeUtils.import("chrome://tbsync/content/tbsync.jsm");

const eas = tbSync.providers.eas;

var tbSyncAbServerSearch = {

  request: async function (accountData, currentQuery)  {
    if (!accountData.getAccountProperty("allowedEasCommands").split(",").includes("Search")) {
        return null;
    }
        
    let response = await eas.network.getSearchResults(accountData, currentQuery);
    let wbxmlData = eas.network.getDataFromResponse(response);
    let data = [];

    if (wbxmlData.Search && wbxmlData.Search.Response && wbxmlData.Search.Response.Store && wbxmlData.Search.Response.Store.Result) {
      let results = eas.xmltools.nodeAsArray(wbxmlData.Search.Response.Store.Result);
      let accountname = accountData.getAccountProperty("accountname");
  
      for (let result of results) {
        if (result.hasOwnProperty("Properties")) {
          //console.log(" RAW : " + JSON.stringify(result.Properties));
          let resultset = {};
          resultset["FirstName"] = result.Properties.FirstName;
          resultset["LastName"] = result.Properties.LastName;
          resultset["DisplayName"] = result.Properties.DisplayName;
          resultset["PrimaryEmail"] = result.Properties.EmailAddress;
          resultset["CellularNumber"] = result.Properties.MobilePhone;
          resultset["HomePhone"] = result.Properties.HomePhone;
          resultset["WorkPhone"] = result.Properties.Phone;
          resultset["Company"] = accountname; //result.Properties.Company;
          resultset["Department"] = result.Properties.Title;
          resultset["JobTitle"] = result.Properties.Office;
          
          data.push(resultset);
        }
      }
    }
    return data;
},

  
  onInject: function (window) {
    this._eventHandler = tbSyncAbServerSearch.eventHandlerWindowReference(window);
    
    let searchbox =  window.document.getElementById("peopleSearchInput");
    if (searchbox) {
      this._searchValue = searchbox.value;
      this._searchValuePollHandler = window.setInterval(function(){tbSyncAbServerSearch.searchValuePoll(window, searchbox)}, 200);
      this._eventHandler.addEventListener(searchbox, "input", false);
    }
    
    let dirtree = window.document.getElementById("dirTree");
    if (dirtree) {
      this._eventHandler.addEventListener(dirtree, "select", false);        
    }
  },
  
  onRemove: function (window) {
    let searchbox =  window.document.getElementById("peopleSearchInput");
    if (searchbox) {
      this._eventHandler.removeEventListener(searchbox, "input", false);
      window.clearInterval(this._searchValuePollHandler);
    }

    let dirtree = window.document.getElementById("dirTree");
    if (dirtree) {
      this._eventHandler.removeEventListener(dirtree, "select", false);        
    }
  },    

  eventHandlerWindowReference: function (window) {
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
          tbSyncAbServerSearch.onSearchInputChanged(this.window);
          break;
        case "select":
          {
            tbSyncAbServerSearch.clearServerSearchResults(this.window);
            let searchbox =  window.document.getElementById("peopleSearchInput");
            let selectedDirectoryURI = window.GetSelectedDirectory();
            
            if (searchbox && selectedDirectoryURI) {
              let addressbook = MailServices.ab.getDirectory(selectedDirectoryURI);
              
              let folders = tbSync.db.findFolders({"target": addressbook.UID}, {"provider": "eas"});
              if (folders.length == 1) {
                searchbox.setAttribute("placeholder", tbSync.getString("addressbook.searchgal::" + tbSync.db.getAccountProperty(folders[0].accountID, "accountname")));
              } else {
                searchbox.setAttribute("placeholder", tbSync.getString((selectedDirectoryURI == "moz-abdirectory://?") ? "addressbook.searchall" : "addressbook.searchthis"));
              }
            }
          }
          break;
      }
    };
    return this;
  },

  searchValuePoll: function (window, searchbox) {
    let value = searchbox.value;
    if (this._searchValue != "" && value == "") {
      tbSyncAbServerSearch.clearServerSearchResults(window);
    }
    this._searchValue = value;
  },

  clearServerSearchResults: function (window) {
    let selectedDirectoryURI = window.GetSelectedDirectory();
    if (selectedDirectoryURI == "moz-abdirectory://?") return; //global search not yet(?) supported
    
    let addressbook = MailServices.ab.getDirectory(selectedDirectoryURI);
    if (addressbook) {
      try {
        let oldresults = addressbook.getCardsFromProperty("X-Server-Searchresult", "TbSync/EAS", true);
        let cardsToDelete = Components.classes["@mozilla.org/array;1"].createInstance(Components.interfaces.nsIMutableArray);
        while (oldresults.hasMoreElements()) {
          cardsToDelete.appendElement(oldresults.getNext(), false);
        }
        addressbook.deleteCards(cardsToDelete);
      } catch (e) {
        //if  getCardsFromProperty is not implemented, do nothing
      }
    }
  },

  onSearchInputChanged: async function (window) {
    let selectedDirectoryURI = window.GetSelectedDirectory();
    if (selectedDirectoryURI == "moz-abdirectory://?") return; //global search not yet(?) supported    
    let addressbook = MailServices.ab.getDirectory(selectedDirectoryURI);
    
    let folders = tbSync.db.findFolders({"target": addressbook.UID}, {"provider": "eas"});
    if (folders.length == 1) {
      let searchbox = window.document.getElementById("peopleSearchInput");
      let query = searchbox.value;        

      let accountData = new tbSync.AccountData(folders[0].accountID);
      let accountname = accountData.getAccountProperty("accountname");
      if (true) { // we may want to disable this

        if (query.length<3) {
          //delete all old results
          tbSyncAbServerSearch.clearServerSearchResults(window);
          window.onEnterInSearchBar();
        } else {          
          this._serverSearchNextQuery = query;                
          if (this._serverSearchBusy) {
            //NOOP
          } else {
            this._serverSearchBusy = true;
            while (this._serverSearchBusy) {

              await tbSync.tools.sleep(1000);
              let currentQuery = this._serverSearchNextQuery;
              this._serverSearchNextQuery = "";
              let results = await tbSyncAbServerSearch.request(accountData, currentQuery);
              //delete all old results
              tbSyncAbServerSearch.clearServerSearchResults(window);

              for (let result of results) {
                let newItem = Components.classes["@mozilla.org/addressbook/cardproperty;1"].createInstance(Components.interfaces.nsIAbCard);
                for (var prop in result) {
                  if (result.hasOwnProperty(prop)) {
                    newItem.setProperty(prop, result[prop]);
                  }
                }
                newItem.setProperty("X-Server-Searchresult", "TbSync/EAS");
                addressbook.addCard(newItem);
              }   
              window.onEnterInSearchBar();
              if (this._serverSearchNextQuery == "") this._serverSearchBusy = false;
            }
          }
        }            
      }
    }
  }
}