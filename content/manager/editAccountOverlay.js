/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */

"use strict";

const eas = tbSync.providers.eas;

var tbSyncEasEditAccount = {

    stripHost: function (document, account) {
        let host = document.getElementById('tbsync.AccountPropertys.pref.host').value;
        if (host.indexOf("https://") == 0) {
            host = host.replace("https://","");
            document.getElementById('tbsync.AccountPropertys.pref.https').checked = true;
            tbSync.db.setAccountProperty(account, "https", "1");
        } else if (host.indexOf("http://") == 0) {
            host = host.replace("http://","");
            document.getElementById('tbsync.AccountPropertys.pref.https').checked = false;
            tbSync.db.setAccountProperty(account, "https", "0");
        }
        
        while (host.endsWith("/")) { host = host.slice(0,-1); }        
        document.getElementById('tbsync.AccountPropertys.pref.host').value = host
        tbSync.db.setAccountProperty(account, "host", host);
    },

    onload: function (window, accountID) {
        // special treatment for configuration label, which is a permanent setting and will not change by switching modes
        let configlabel = window.document.getElementById("tbsync.accountsettings.label.config");
        if (configlabel) {
            configlabel.setAttribute("value", tbSync.getString("config.custom", "eas"));
        }
    },
};
