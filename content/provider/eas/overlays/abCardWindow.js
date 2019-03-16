/*
 * This file is part of TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

Components.utils.import("chrome://tbsync/content/tbsync.jsm");
Components.utils.import("resource://gre/modules/osfile.jsm");

var tbSyncAbEasCardWindow = {
    
    onBeforeInject: function (window) {
        let cardProvider = "";
        let aParentDirURI  = "";
        
        if (window.location.href=="chrome://messenger/content/addressbook/abNewCardDialog.xul") {
            aParentDirURI = window.document.getElementById("abPopup").value;
        } else {
            aParentDirURI = tbSyncAbEasCardWindow.getSelectedAbFromArgument(window.arguments[0]);
        }

        if (aParentDirURI) {
            let folders = tbSync.db.findFoldersWithSetting("target", aParentDirURI);
            if (folders.length == 1) {
                cardProvider = tbSync.db.getAccountSetting(folders[0].account, "provider");
            }
        }
        
        //returning false will prevent injection
        return (cardProvider == "eas");
    },

    getSelectedAbFromArgument: function (arg) {
        let abURI = "";
        if (arg.hasOwnProperty("abURI")) {
            abURI = arg.abURI;
        } else if (arg.hasOwnProperty("selectedAB")) {
            abURI = arg.selectedAB;
        }
        
        if (abURI) {
            let abManager = Components.classes["@mozilla.org/abmanager;1"].getService(Components.interfaces.nsIAbManager);
            let ab = abManager.getDirectory(abURI);
            if (ab.isMailList) {
                let parts = abURI.split("/");
                parts.pop();
                return parts.join("/");
            }
        }
        return abURI;
    },


    
    onInject: function (window) {
        //keep track of default elements we hide/disable, so it can be undone during overlay remove
        tbSyncAbEasCardWindow.elementsToHide = [];
        tbSyncAbEasCardWindow.elementsToDisable = [];
        
        //hide stuff from gContactSync *grrrr* - I cannot hide all because he adds them via javascript :-(
        tbSyncAbEasCardWindow.elementsToHide.push(window.document.getElementById("gContactSyncTab"));

        //hide registered default elements
        for (let i=0; i < tbSyncAbEasCardWindow.elementsToHide.length; i++) {
            if (tbSyncAbEasCardWindow.elementsToHide[i]) {
                tbSyncAbEasCardWindow.elementsToHide[i].collapsed = true;
            }
        }

        //disable registered default elements
        for (let i=0; i < tbSyncAbEasCardWindow.elementsToDisable.length; i++) {
            if (tbSyncAbEasCardWindow.elementsToDisable[i]) {
                tbSyncAbEasCardWindow.elementsToDisable[i].disabled = true;
            }
        }

        if (window.location.href=="chrome://messenger/content/addressbook/abNewCardDialog.xul") {
            window.sizeToContent(); 
            window.RegisterSaveListener(tbSyncAbEasCardWindow.onSaveCard);        
        } else {            
            window.RegisterLoadListener(tbSyncAbEasCardWindow.onLoadCard);
            window.RegisterSaveListener(tbSyncAbEasCardWindow.onSaveCard);

            //if this window was open during inject, load the extra fields
            if (gEditCard) tbSyncAbEasCardWindow.onLoadCard(gEditCard.card, window.document);
        }
    },

    onRemove: function (window) {
        if (window.location.href=="chrome://messenger/content/addressbook/abNewCardDialog.xul") {
            window.UnregisterSaveListener(tbSyncAbEasCardWindow.onSaveCard);
        } else {
            window.UnregisterLoadListener(tbSyncAbEasCardWindow.onLoadCard);
            window.UnregisterSaveListener(tbSyncAbEasCardWindow.onSaveCard);
        }
          
        //unhide elements hidden by this provider
        for (let i=0; i < tbSyncAbEasCardWindow.elementsToHide.length; i++) {
            if (tbSyncAbEasCardWindow.elementsToHide[i]) {
                tbSyncAbEasCardWindow.elementsToHide[i].collapsed = false;
            }
        }

        //re-enable elements disabled by this provider
        for (let i=0; i < tbSyncAbEasCardWindow.elementsToDisable.length; i++) {
            if (tbSyncAbEasCardWindow.elementsToDisable[i]) {
                tbSyncAbEasCardWindow.elementsToDisable[i].disabled = false;
            }
        }
    },
    

    
    onLoadCard: function (aCard, aDocument) {                
        //load properties
        let items = aDocument.getElementsByClassName("easProperty");
        for (let i=0; i < items.length; i++) {
            items[i].value = aCard.getProperty(items[i].id, "");
        }
    },
    
    onSaveCard: function (aCard, aDocument) {
        let items = aDocument.getElementsByClassName("easProperty");
        for (let i=0; i < items.length; i++) {
            aCard.setProperty(items[i].id, items[i].value);
        }
    }
    
}
