/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

Components.utils.import("resource://gre/modules/Task.jsm");
Components.utils.import("chrome://tbsync/content/tbsync.jsm");

var tbSyncEasNewAccount = {

    startTime: 0,
    maxTimeout: 30,

    onClose: function () {
        //tbSync.dump("onClose", tbSync.addAccountWindowOpen);
        return !document.documentElement.getButton("cancel").disabled;
    },

    onLoad: function () {
        this.elementName = document.getElementById('tbsync.newaccount.name');
        this.elementUser = document.getElementById('tbsync.newaccount.user');
        this.elementUrl = document.getElementById('tbsync.newaccount.url');
        this.elementPass = document.getElementById('tbsync.newaccount.password');
        this.elementServertype = document.getElementById('tbsync.newaccount.servertype');
        
        document.documentElement.getButton("extra1").disabled = true;
        document.documentElement.getButton("extra1").label = tbSync.getLocalizedMessage("newaccount.add_auto","eas");
        document.getElementById('tbsync.newaccount.autodiscoverlabel').hidden = true;
        document.getElementById('tbsync.newaccount.autodiscoverstatus').hidden = true;

        document.getElementById('tbsync.newaccount.url.box').style.visibility =  (this.elementServertype.value != "custom") ? "hidden" : "visible";
        document.getElementById("tbsync.newaccount.name").focus();
    },

    onUnload: function () {
    },

    onUserTextInput: function () {
        if (this.elementServertype.value != "custom") {
            document.documentElement.getButton("extra1").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "" || this.elementPass.value == "");
        } else {
            document.documentElement.getButton("extra1").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "" || this.elementPass.value == "" ||  this.elementUrl.value.trim() == "");
        }
    },

    onUserDropdown: function () {
        document.documentElement.getButton("extra1").label = tbSync.getLocalizedMessage("newaccount.add_" + this.elementServertype.value,"eas");
        document.getElementById('tbsync.newaccount.url.box').style.visibility = (this.elementServertype.value != "custom") ? "hidden" : "visible";
        this.onUserTextInput();
    },

    onAdd: Task.async (function* () {
        if (document.documentElement.getButton("extra1").disabled == false) {
            let user = this.elementUser.value;
            let password = this.elementPass.value;
            let servertype = this.elementServertype.value;
            let accountname = this.elementName.value.trim();
            let url = this.elementUrl.value.trim();

            if (servertype == "custom") {
                tbSyncEasNewAccount.addAccount(user, password, servertype, accountname, url);                
            }
            
            if (servertype == "auto") {

                if (user.split("@").length != 2) {
                    alert(tbSync.getLocalizedMessage("autodiscover.NeedEmail","eas"))
                    return
                }

                document.documentElement.getButton("cancel").disabled = true;
                document.documentElement.getButton("extra1").disabled = true;
                document.getElementById("tbsync.newaccount.name").disabled = true;
                document.getElementById("tbsync.newaccount.user").disabled = true;
                document.getElementById("tbsync.newaccount.password").disabled = true;
                document.getElementById("tbsync.newaccount.servertype").disabled = true;

                let updateTimer = Components.classes["@mozilla.org/timer;1"].createInstance(Components.interfaces.nsITimer);
                updateTimer.initWithCallback({notify : function () {tbSyncEasNewAccount.updateAutodiscoverStatus()}}, 1000, 3);

                tbSyncEasNewAccount.startTime = Date.now();
                tbSyncEasNewAccount.updateAutodiscoverStatus();

                let result = yield tbSync.eas.getServerConnectionViaAutodiscover(user, password, tbSyncEasNewAccount.maxTimeout*1000);
                updateTimer.cancel();

                document.getElementById('tbsync.newaccount.autodiscoverlabel').hidden = true;
                document.getElementById('tbsync.newaccount.autodiscoverstatus').hidden = true;

                if (result.server) {
                    alert(tbSync.getLocalizedMessage("autodiscover.Ok","eas"));
                    //add account with found server url
                    tbSyncEasNewAccount.addAccount(result.user, password, servertype, accountname, result.server);                
                } else {                    
                    alert(tbSync.getLocalizedMessage("autodiscover.Failed","eas").replace("##user##", result.user) + "\n\n" + result.error);
                }

                document.getElementById("tbsync.newaccount.name").disabled = false;
                document.getElementById("tbsync.newaccount.user").disabled = false;
                document.getElementById("tbsync.newaccount.password").disabled = false;
                document.getElementById("tbsync.newaccount.servertype").disabled = false;

                document.documentElement.getButton("cancel").disabled = false;
                document.documentElement.getButton("extra1").disabled = false;
            }

        }
    }),

    updateAutodiscoverStatus: function () {
        document.getElementById('tbsync.newaccount.autodiscoverstatus').hidden = false;
        let offset = Math.round(((Date.now()-tbSyncEasNewAccount.startTime)/1000));
        let timeout = (offset>2) ? " (" + (tbSyncEasNewAccount.maxTimeout - offset) + ")" : "";

        document.getElementById('tbsync.newaccount.autodiscoverstatus').value  = tbSync.getLocalizedMessage("autodiscover.Querying","eas") + timeout;
    },

    addAccount (user, password, servertype, accountname, url) {
        let newAccountEntry = tbSync.eas.getDefaultAccountEntries();
        newAccountEntry.accountname = accountname;
        newAccountEntry.user = user;
        newAccountEntry.servertype = servertype;

        if (url) {
            //if no protocoll is given, prepend "https://"
            if (url.substring(0,4) != "http" || url.indexOf("://") == -1) url = "https://" + url.split("://").join("/");
            newAccountEntry.host = tbSync.eas.stripAutodiscoverUrl(url);
            newAccountEntry.https = (url.substring(0,5) == "https") ? "1" : "0";
            //also update password in PasswordManager (only works if url is present)
            tbSync.eas.setPassword (newAccountEntry, password);
        }

        //create a new EAS account and pass its id to updateAccountsList, which will select it
        //the onSelect event of the List will load the selected account
        window.opener.tbSyncAccounts.updateAccountsList(tbSync.db.addAccount(newAccountEntry));

        window.close();
    }
};
