/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */

"use strict";

const eas = TbSync.providers.eas;

var tbSyncEditAccountOverlay = {

    onload: function (window, accountData) {
        this.accountData = accountData;
        
        // special treatment for configuration label, which is a permanent setting and will not change by switching modes
        let configlabel = window.document.getElementById("tbsync.accountsettings.label.config");
        if (configlabel) {
            configlabel.setAttribute("value", TbSync.getString("config.custom", "eas"));
        }
    },

    stripHost: function (document) {
        let host = document.getElementById('tbsync.AccountPropertys.pref.host').value;
        if (host.indexOf("https://") == 0) {
            host = host.replace("https://","");
            document.getElementById('tbsync.AccountPropertys.pref.https').checked = true;
            this.accountData.setAccountProperty("https", true);
        } else if (host.indexOf("http://") == 0) {
            host = host.replace("http://","");
            document.getElementById('tbsync.AccountPropertys.pref.https').checked = false;
           this.accountData.setAccountProperty("https", false);
        }
        
        while (host.endsWith("/")) { host = host.slice(0,-1); }        
        document.getElementById('tbsync.AccountPropertys.pref.host').value = host
       this.accountData.setAccountProperty("host", host);
    },
    
    deleteFolder: function() {
        let folderList = document.getElementById("tbsync.accountsettings.folderlist");
        if (folderList.selectedItem !== null && !folderList.disabled) {
            let folderData = folderList.selectedItem.folderData;

            //only trashed folders can be purged (for example O365 does not show deleted folders but also does not allow to purge them)
            if (!eas.tools.parentIsTrash(folderData)) return;
            
            if (folderData.getFolderProperty("selected")) window.alert(TbSync.getString("deletefolder.notallowed::" + folderData.getFolderProperty("foldername"), "eas"));
            else if (window.confirm(TbSync.getString("deletefolder.confirm::" + folderData.getFolderProperty("foldername"), "eas"))) {
                folderData.sync({syncList: false, syncJob: "deletefolder"});
            } 
        }            
    }    
};
