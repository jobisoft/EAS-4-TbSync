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
        let aParentDirURI  = "";

        if (window.location.href=="chrome://messenger/content/addressbook/abNewCardDialog.xul") {
            //get provider via uri from drop down
            aParentDirURI = window.document.getElementById("abPopup").value;
        } else {
            //function to get correct uri of current card for global book as well for mailLists
            aParentDirURI = tbSync.providers.eas.tools.getSelectedUri(window.arguments[0].abURI, window.arguments[0].card);
        }
        
        //returning false will prevent injection
        return (MailServices.ab.getDirectory(aParentDirURI).getStringValue("tbSyncProvider", "") == "eas");
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
            window.RegisterSaveListener(tbSyncAbEasCardWindow.onSaveCard);        
        } else {            
            window.RegisterLoadListener(tbSyncAbEasCardWindow.onLoadCard);
            window.RegisterSaveListener(tbSyncAbEasCardWindow.onSaveCard);

            //if this window was open during inject, load the extra fields
            if (gEditCard) tbSyncAbEasCardWindow.onLoadCard(gEditCard.card, window.document);
        }
        window.sizeToContent();
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
