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
    validating: false,

    onClose: function () {
        //disallow closing of wizard while validating
        return !this.validating;
    },

    onCancel: function () {
        //disallow closing of wizard while validating
        return !this.validating;
    },

    onLoad: function () {
        this.elementName = document.getElementById('tbsync.newaccount.name');
        this.elementUser = document.getElementById('tbsync.newaccount.user');
        this.elementUrl = document.getElementById('tbsync.newaccount.url');
        this.elementPass = document.getElementById('tbsync.newaccount.password');
        this.elementServertype = document.getElementById('tbsync.newaccount.servertype');
        
        document.documentElement.getButton("back").hidden = true;
        document.documentElement.getButton("finish").disabled = true;
        document.documentElement.getButton("finish").label = tbSync.getLocalizedMessage("newaccount.add_auto","eas");

        document.getElementById("tbsync.error").hidden = true;
        document.getElementById("tbsync.spinner").hidden = true;

        document.getElementById('tbsync.newaccount.url.box').style.visibility =  (this.elementServertype.value != "custom") ? "hidden" : "visible";
        document.getElementById("tbsync.newaccount.name").focus();
    },

    onUnload: function () {
    },

    onUserTextInput: function () {
        document.getElementById("tbsync.error").hidden = true;
        if (this.elementServertype.value != "custom") {
            document.documentElement.getButton("finish").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "" || this.elementPass.value == "");
        } else {
            document.documentElement.getButton("finish").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "" || this.elementPass.value == "" ||  this.elementUrl.value.trim() == "");
        }
    },

    onUserDropdown: function () {
        document.documentElement.getButton("finish").label = tbSync.getLocalizedMessage("newaccount.add_" + this.elementServertype.value,"eas");
        document.getElementById('tbsync.newaccount.url.box').style.visibility = (this.elementServertype.value != "custom") ? "hidden" : "visible";
        this.onUserTextInput();
    },

    onFinish: function () {
        if (document.documentElement.getButton("finish").disabled == false) {
            //initiate validation of server connection
            this.validate();
        }
        return false;
    },

    validate: Task.async (function* () {
        let user = this.elementUser.value;
        let password = this.elementPass.value;
        let servertype = this.elementServertype.value;
        let accountname = this.elementName.value.trim();
        let url = this.elementUrl.value.trim();

        if (servertype == "auto" &&  user.split("@").length != 2) {
            alert(tbSync.getLocalizedMessage("autodiscover.NeedEmail","eas"))
            return;
        }
        
        this.validating = true;
        let error = "";
        
        //document.getElementById("tbsync.newaccount.wizard").canRewind = false;        
        document.getElementById("tbsync.error").hidden = true;
        document.getElementById("tbsync.spinner").hidden = false;
        document.documentElement.getButton("cancel").disabled = true;
        document.documentElement.getButton("finish").disabled = true;
        document.getElementById("tbsync.newaccount.name").disabled = true;
        document.getElementById("tbsync.newaccount.user").disabled = true;
        document.getElementById("tbsync.newaccount.password").disabled = true;
        document.getElementById("tbsync.newaccount.servertype").disabled = true;
        
        //do autodiscover
        if (servertype == "auto") {
            let updateTimer = Components.classes["@mozilla.org/timer;1"].createInstance(Components.interfaces.nsITimer);
            updateTimer.initWithCallback({notify : function () {tbSyncEasNewAccount.updateAutodiscoverStatus()}}, 1000, 3);

            tbSyncEasNewAccount.startTime = Date.now();
            tbSyncEasNewAccount.updateAutodiscoverStatus();

            let result = yield tbSync.eas.getServerConnectionViaAutodiscover(user, password, tbSyncEasNewAccount.maxTimeout*1000);
            updateTimer.cancel();
    
            if (result.server) {
                user = result.user;
                url = result.server;
            } else {                    
                error = result.error;
            }
        }

        //now validate the information
        if (!error) {
        }

        //add if valid
        if (!error) {
            tbSyncEasNewAccount.addAccount(user, password, servertype, accountname, url);
        }
        
        //end validation
        document.getElementById("tbsync.newaccount.name").disabled = false;
        document.getElementById("tbsync.newaccount.user").disabled = false;
        document.getElementById("tbsync.newaccount.password").disabled = false;
        document.getElementById("tbsync.newaccount.servertype").disabled = false;
        document.documentElement.getButton("cancel").disabled = false;
        document.documentElement.getButton("finish").disabled = false;
        document.getElementById("tbsync.spinner").hidden = true;
        //document.getElementById("tbsync.newaccount.wizard").canRewind = true;
        
        this.validating = false;
        
        //close wizard, if done
        if (!error) {
            document.getElementById("tbsync.newaccount.wizard").cancel();        
        } else {
            document.getElementById("tbsync.error.message").textContent = error;
            document.getElementById("tbsync.error").hidden = false;
        }            
    }),

    updateAutodiscoverStatus: function () {
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
