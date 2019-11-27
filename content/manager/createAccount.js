/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

var { TbSync } = ChromeUtils.import("chrome://tbsync/content/tbsync.jsm");

const eas = TbSync.providers.eas;

var tbSyncEasNewAccount = {

    startTime: 0,
    maxTimeout: 30,
    validating: false,

    onClose: function () {
        //disallow closing of wizard while validating
        return !this.validating;
    },

    onCancel: function (event) {
        //disallow closing of wizard while validating
        if (this.validating) {
            event.preventDefault();
        }
    },

    onLoad: function () {
        this.providerData = new TbSync.ProviderData("eas");

        this.elementName = document.getElementById('tbsync.newaccount.name');
        this.elementUser = document.getElementById('tbsync.newaccount.user');
        this.elementUrl = document.getElementById('tbsync.newaccount.url');
        this.elementPass = document.getElementById('tbsync.newaccount.password');
        this.elementServertype = document.getElementById('tbsync.newaccount.servertype');
        
        document.documentElement.getButton("back").hidden = true;
        this.onUserDropdown();

        document.getElementById("tbsync.error").hidden = true;
        document.getElementById("tbsync.spinner").hidden = true;

        document.addEventListener("wizardfinish", tbSyncEasNewAccount.onFinish.bind(this));
        document.addEventListener("wizardcancel", tbSyncEasNewAccount.onCancel.bind(this));
    },

    onUnload: function () {
    },

    onUserTextInput: function () {
        document.getElementById("tbsync.error").hidden = true;
        switch (this.elementServertype.value) {
            case "select":            
                document.documentElement.getButton("finish").disabled = true;
                break;

            case "auto":            
                document.documentElement.getButton("finish").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "" || this.elementPass.value == "");
                break;
            
            case "office365":            
                document.documentElement.getButton("finish").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "");
                break;

            case "custom":
            default:
                document.documentElement.getButton("finish").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "" || this.elementPass.value == "" ||  this.elementUrl.value.trim() == "");
                break;
        }
    },

    onUserDropdown: function () {
        switch (this.elementServertype.value) {
            case "select":            
                document.getElementById('tbsync.newaccount.user.box').style.visibility = "hidden";
                document.getElementById('tbsync.newaccount.url.box').style.visibility = "hidden";
                document.getElementById('tbsync.newaccount.password.box').style.visibility = "hidden";
                document.documentElement.getButton("finish").label = TbSync.getString("newaccount.add_custom","eas");
                break;

            case "auto":            
                document.getElementById('tbsync.newaccount.user.box').style.visibility = "visible";
                document.getElementById('tbsync.newaccount.url.box').style.visibility = "hidden";
                document.getElementById('tbsync.newaccount.password.box').style.visibility = "visible";
                document.documentElement.getButton("finish").label = TbSync.getString("newaccount.add_auto","eas");
                break;
            
            case "office365":            
                document.getElementById('tbsync.newaccount.user.box').style.visibility = "visible";
                document.getElementById('tbsync.newaccount.url.box').style.visibility = "hidden";
                document.getElementById('tbsync.newaccount.password.box').style.visibility = "hidden";
                document.documentElement.getButton("finish").label = TbSync.getString("newaccount.add_custom","eas");
                break;

            case "custom":
            default:
                document.getElementById('tbsync.newaccount.user.box').style.visibility = "visible";
                document.getElementById('tbsync.newaccount.url.box').style.visibility = "visible";
                document.getElementById('tbsync.newaccount.password.box').style.visibility = "visible";
                document.documentElement.getButton("finish").label = TbSync.getString("newaccount.add_custom","eas");
                break;
        }
        this.onUserTextInput();
        //document.getElementById("tbsync.newaccount.name").focus();        
    },

    onFinish: function (event) {
        if (document.documentElement.getButton("finish").disabled == false) {
            //initiate validation of server connection
            this.validate();
        }
        event.preventDefault();
    },

    validate: async function () {
        let user = this.elementUser.value;
        let servertype = this.elementServertype.value;
        let accountname = this.elementName.value.trim();

        let url = (servertype == "custom") ?this.elementUrl.value.trim() : "";
        let password = (servertype == "auto" || servertype == "custom") ? this.elementPass.value : "";

        if ((servertype == "auto" || servertype == "office365") && user.split("@").length != 2) {
            alert(TbSync.getString("autodiscover.NeedEmail","eas"))
            return;
        }
        
        this.validating = true;
        let error = "";
        
        //document.getElementById("tbsync.newaccount.wizard").canRewind = false;        
        document.getElementById("tbsync.error").hidden = true;
        document.documentElement.getButton("cancel").disabled = true;
        document.documentElement.getButton("finish").disabled = true;
        document.getElementById("tbsync.newaccount.name").disabled = true;
        document.getElementById("tbsync.newaccount.user").disabled = true;
        document.getElementById("tbsync.newaccount.password").disabled = true;
        document.getElementById("tbsync.newaccount.servertype").disabled = true;

        tbSyncEasNewAccount.startTime = Date.now();
        tbSyncEasNewAccount.updateAutodiscoverStatus();
        document.getElementById("tbsync.spinner").hidden = false;
        
        //do autodiscover
        if (servertype == "office365" || servertype == "auto") {
            let updateTimer = Components.classes["@mozilla.org/timer;1"].createInstance(Components.interfaces.nsITimer);
            updateTimer.initWithCallback({notify : function () {tbSyncEasNewAccount.updateAutodiscoverStatus()}}, 1000, 3);

            if (servertype == "office365") {
                let v2 = await eas.network.getServerConnectionViaAutodiscoverV2JsonRequest("https://autodiscover-s.outlook.com/autodiscover/autodiscover.json?Email="+encodeURIComponent(user)+"&Protocol=ActiveSync");
                let oauthData = eas.network.getOAuthObj({ host: v2.server, user, accountname });
                if (oauthData) {
                    // ask for token
                    document.getElementById("tbsync.spinner").hidden = true;
                    let _rv = {};
                    if (await oauthData.asyncConnect(_rv)) {
                        password = _rv.tokens;
                    } else {
                        error = TbSync.getString("status." + _rv.error, "eas");
                    }
                    document.getElementById("tbsync.spinner").hidden = false;                
                    url=v2.server;
                } else {
                    error = TbSync.getString("status.404", "eas");
                }
            } else {
                let result = await eas.network.getServerConnectionViaAutodiscover(user, password, tbSyncEasNewAccount.maxTimeout*1000);
                if (result.server) {
                    user = result.user;
                    url = result.server;
                } else {                    
                    error = result.error; // is a localized string
                }
            }

            updateTimer.cancel();
        }

        //now validate the information
        if (!error) {
            if (!password) error = TbSync.getString("status.401", "eas");
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
    },

    updateAutodiscoverStatus: function () {
        let offset = Math.round(((Date.now() - tbSyncEasNewAccount.startTime)/1000));
        let timeout = (offset>2) ? " (" + (tbSyncEasNewAccount.maxTimeout - offset) + ")" : "";

        document.getElementById('tbsync.newaccount.autodiscoverstatus').value  = TbSync.getString("autodiscover.Querying","eas") + timeout;
    },

    addAccount (user, password, servertype, accountname, url) {
        let newAccountEntry = this.providerData.getDefaultAccountEntries();
        newAccountEntry.user = user;
        newAccountEntry.servertype = servertype;

        if (url) {
            //if no protocoll is given, prepend "https://"
            if (url.substring(0,4) != "http" || url.indexOf("://") == -1) url = "https://" + url.split("://").join("/");
            newAccountEntry.host = eas.network.stripAutodiscoverUrl(url);
            newAccountEntry.https = (url.substring(0,5) == "https");
        }

        // Add the new account.
        let newAccountData = this.providerData.addAccount(accountname, newAccountEntry);
        eas.network.getAuthData(newAccountData).updateLoginData(user, password);

        window.close();
    }
};
