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