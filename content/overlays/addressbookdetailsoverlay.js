/*
 * This file is part of TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

Components.utils.import("chrome://tbsync/content/tbsync.jsm");

var tbSyncEasAddressBookDetails = {
    
    onBeforeInject: function (window) {
        //we inject always now and let onAbResultSelectionChanged handle our custom display
        return true;

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

            //function to get correct uri of current card for global book as well for mailLists
            let abUri = TbSync.providers.eas.tools.getSelectedUri(window.GetSelectedDirectory(), aCard);
            let show = (MailServices.ab.getDirectory(abUri).getStringValue("tbSyncProvider", "") == "eas");
        
            let email3Box = window.document.getElementById("cvEmail3Box");
            if (email3Box) {
                if (show) {
                    let email3Value = aCard.getProperty("Email3Address","");
                    if (email3Value) {
                        email3Box.collapsed = false;
                        let email3Element = window.document.getElementById("cvEmail3");
                        window.HandleLink(email3Element, window.zSecondaryEmail, email3Value, email3Box, "mailto:" + email3Value);
                    }
                } else {
                        email3Box.collapsed = true;                    
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
                        if (show) {
                            let value = aCard.getProperty(phoneNumbers[field],"");
                            if (value) {
                                element.collapsed = false;
                                element.textContent = element.getAttribute("labelprefix") + " " + value;
                                phoneFound = true;
                            }
                        } else {
                            element.collapsed = true;                            
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
