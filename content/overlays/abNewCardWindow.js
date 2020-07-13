/*
 * This file is part of TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

Components.utils.import("chrome://tbsync/content/tbsync.jsm");

var tbSyncEasAbNewCardWindow = {

    onInject: function (window) {
        window.document.getElementById("abPopup").addEventListener("select", tbSyncEasAbNewCardWindow.onAbSelectChangeNewCard, false);
    },

    onRemove: function (window) {
        window.document.getElementById("abPopup").removeEventListener("select", tbSyncEasAbNewCardWindow.onAbSelectChangeNewCard, false);
    },

    onAbSelectChangeNewCard: function () {        
        //remove our overlay (if injected)
        TbSync.providers.eas.overlayManager.removeOverlay(window, "chrome://eas4tbsync/content/overlays/abCardWindow.xhtml");
        //inject our overlay (if our card)
        TbSync.providers.eas.overlayManager.injectOverlay(window, "chrome://eas4tbsync/content/overlays/abCardWindow.xhtml");
    },
        
}
