/*
 * This file is part of TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

Components.utils.import("chrome://tbsync/content/tbsync.jsm");

var tbSyncEasAddressBook = {

    onInject: function (window) {
        if (window.document.getElementById("dirTree")) {
            window.document.getElementById("dirTree").addEventListener("select", tbSyncEasAddressBook.onAbDirectorySelectionChanged, false);
        }
    },

    onRemove: function (window) {
        if (window.document.getElementById("dirTree")) {
            window.document.getElementById("dirTree").removeEventListener("select", tbSyncEasAddressBook.onAbDirectorySelectionChanged, false);
        }
    },
    
    onAbDirectorySelectionChanged: function () {
        //TODO: Do not do this, if provider did not change
        //remove our details injection (if injected)
        tbSync.eas.overlayManager.removeOverlay(window, "chrome://eas4tbsync/content/provider/eas/overlays/addressbookdetailsoverlay.xul");
        //inject our details injection (if the new selected book is us)
        tbSync.eas.overlayManager.injectOverlay(window, "chrome://eas4tbsync/content/provider/eas/overlays/addressbookdetailsoverlay.xul");
    }
}



var tbSyncEasAddressBookDetails = {
    
    onBeforeInject: function (window) {
        let cardProvider = "";
        
        try {
            let aParentDirURI = window.GetSelectedDirectory();
            let abManager = Components.classes["@mozilla.org/abmanager;1"].getService(Components.interfaces.nsIAbManager);
            let selectedBook = abManager.getDirectory(aParentDirURI);
            if (selectedBook.isMailList) {
                aParentDirURI = aParentDirURI.substring(0, aParentDirURI.lastIndexOf("/"));
            }

            if (aParentDirURI) {
                let folders = tbSync.db.findFoldersWithSetting("target", aParentDirURI);
                if (folders.length == 1) {
                    cardProvider = tbSync.db.getAccountSetting(folders[0].account, "provider");
                }
            }
        } catch (e) {
            //if the window / gDirTree is not yet avail 
        }
        
        //returning false will prevent injection
        return (cardProvider == "eas");
    },

    onInject: function (window) {
        if (window.document.getElementById("abResultsTree")) {
            window.document.getElementById("abResultsTree").addEventListener("select", tbSyncEasAddressBookDetails.onAbResultSelectionChanged, false);
            tbSyncEasAddressBookDetails.onAbResultSelectionChanged();
        }
    },

    onRemove: function (window) {
        if (window.document.getElementById("abResultsTree")) {
            window.document.getElementById("abResultsTree").removeEventListener("select", tbSyncEasAddressBookDetails.onAbResultSelectionChanged, false);
        }
    },
    
    onAbResultSelectionChanged: function () {
        let cards = window.GetSelectedAbCards();
        if (cards.length == 1) {
            let aCard = cards[0];
            
            let email3Box = window.document.getElementById("cvEmail3Box");
            if (email3Box) {
                let email3Value = aCard.getProperty("Email3Address","");
                if (email3Value) {
                email3Box.collapsed = false;
                let email3Element = window.document.getElementById("cvEmail3");
                window.HandleLink(email3Element, window.zSecondaryEmail, email3Value, email3Box, "mailto:" + email3Value);
                }
            }
            
            let phoneNumbers = {
                easPhWork2: "Business2PhoneNumber",
                easPhWorkFax: "BusinessFaxNumber",
                easPhCompany: "CompanyMainPhone",
                easPhAssistant: "AssistantPhoneNumber",
                easPhHome2: "Home2PhoneNumber",
                easPhCar: "CarPhoneNumber",
                easPhRadio: "RadioPhoneNumber"
            };
            
            let phoneFound = false;
            for (let field in phoneNumbers) {
                if (phoneNumbers.hasOwnProperty(field)) {
                let element = window.document.getElementById(field);
                if (element) {
                    let value = aCard.getProperty(phoneNumbers[field],"");
                    if (value) {
                    element.collapsed = false;
                    element.textContent = element.getAttribute("labelprefix") + " " + value;
                    phoneFound = true;
                    }
                }
                }
            }

            if (phoneFound) {
                window.document.getElementById("cvbPhone").collapsed = false;
                window.document.getElementById("cvhPhone").collapsed = false;
            }

        }
    },
    
}

