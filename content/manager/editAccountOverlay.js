/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

var tbSyncEasEditAccount = {

    onInject: function (window) {
        //for some unknown reason, my OverlayManager cannot create menulists, so I need to do that
        //manually and append the already loaded menupopus into the manually created menulists

        let asversionPopup = window.document.getElementById('asversion.popup');
        let asversionHook = window.document.getElementById('asversion.hook');
        let asversionMenuList = window.document.createElement("menulist");
        asversionMenuList.setAttribute("id", "tbsync.accountsettings.pref.asversionselected");
        asversionMenuList.setAttribute("class", "lockIfConnected");
        asversionMenuList.appendChild(asversionPopup);
        //add after the hook element
        asversionHook.parentNode.insertBefore(asversionMenuList, asversionHook.nextSibling);

        let separatorPopup = window.document.getElementById('separator.popup');
        let separatorHook = window.document.getElementById('separator.hook');
        let separatorMenuList = window.document.createElement("menulist");
        separatorMenuList.setAttribute("id", "tbsync.accountsettings.pref.seperator");
        separatorMenuList.setAttribute("class", "lockIfConnected");
        separatorMenuList.appendChild(separatorPopup);
        //add before the hook element
        separatorHook.parentNode.insertBefore(separatorMenuList, separatorHook);        
    },

};
