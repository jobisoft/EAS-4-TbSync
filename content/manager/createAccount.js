/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */

"use strict";

var { ExtensionParent } = ChromeUtils.importESModule(
    "resource://gre/modules/ExtensionParent.sys.mjs"
);

var tbsyncExtension = ExtensionParent.GlobalManager.getExtension(
    "tbsync@jobisoft.de"
);
var { TbSync } = ChromeUtils.importESModule(
    `chrome://tbsync/content/tbsync.sys.mjs?${tbsyncExtension.manifest.version}`
);

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

        document.getElementById("tbsync.newaccount.wizard").getButton("back").hidden = true;
        this.onUserDropdown();

        document.getElementById("tbsync.error").hidden = true;
        document.getElementById("tbsync.spinner").hidden = true;

        document.addEventListener("wizardfinish", tbSyncEasNewAccount.onFinish.bind(this));
        document.addEventListener("wizardcancel", tbSyncEasNewAccount.onCancel.bind(this));
        // bug https://bugzilla.mozilla.org/show_bug.cgi?id=1618252
        document.getElementById('tbsync.newaccount.wizard')._adjustWizardHeader();
    },

    onUnload: function () {
    },

    onUserTextInput: function () {
        document.getElementById("tbsync.error").hidden = true;
        switch (this.elementServertype.value) {
            case "select":
                document.getElementById("tbsync.newaccount.wizard").getButton("finish").disabled = true;
                break;

            case "auto":
                document.getElementById("tbsync.newaccount.wizard").getButton("finish").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "" || this.elementPass.value == "");
                break;

            case "office365":
                document.getElementById("tbsync.newaccount.wizard").getButton("finish").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "");
                break;

            case "custom":
            default:
                document.getElementById("tbsync.newaccount.wizard").getButton("finish").disabled = (this.elementName.value.trim() == "" || this.elementUser.value == "" || this.elementPass.value == "" || this.elementUrl.value.trim() == "");
                break;
        }
    },

    onUserDropdown: function () {
        if (this.elementServertype) {
            switch (this.elementServertype.value) {
                case "select":
                    document.getElementById('tbsync.newaccount.user.box').hidden = true;
                    document.getElementById('tbsync.newaccount.url.box').hidden = true;
                    document.getElementById('tbsync.newaccount.password.box').hidden = true;
                    document.getElementById("tbsync.newaccount.wizard").getButton("finish").label = TbSync.getString("newaccount.add_custom", "eas");
                    break;

                case "auto":
                    document.getElementById('tbsync.newaccount.user.box').hidden = false;
                    document.getElementById('tbsync.newaccount.url.box').hidden = true;
                    document.getElementById('tbsync.newaccount.password.box').hidden = false;
                    document.getElementById("tbsync.newaccount.wizard").getButton("finish").label = TbSync.getString("newaccount.add_auto", "eas");
                    break;

                case "office365":
                    document.getElementById('tbsync.newaccount.user.box').hidden = false;
                    document.getElementById('tbsync.newaccount.url.box').hidden = true;
                    document.getElementById('tbsync.newaccount.password.box').hidden = true;
                    document.getElementById("tbsync.newaccount.wizard").getButton("finish").label = TbSync.getString("newaccount.add_custom", "eas");
                    break;

                case "custom":
                default:
                    document.getElementById('tbsync.newaccount.user.box').hidden = false;
                    document.getElementById('tbsync.newaccount.url.box').hidden = false;
                    document.getElementById('tbsync.newaccount.password.box').hidden = false;
                    document.getElementById("tbsync.newaccount.wizard").getButton("finish").label = TbSync.getString("newaccount.add_custom", "eas");
                    break;
            }
            this.onUserTextInput();
            //document.getElementById("tbsync.newaccount.name").focus();
        }
    },

    onFinish: function (event) {
        if (document.getElementById("tbsync.newaccount.wizard").getButton("finish").disabled == false) {
            //initiate validation of server connection
            this.validate();
        }
        event.preventDefault();
    },

    validate: async function () {
        let user = this.elementUser.value;
        let servertype = this.elementServertype.value;
        let accountname = this.elementName.value.trim();

        let url = (servertype == "custom") ? this.elementUrl.value.trim() : "";
        let password = (servertype == "auto" || servertype == "custom") ? this.elementPass.value : "";

        if ((servertype == "auto" || servertype == "office365") && user.split("@").length != 2) {
            alert(TbSync.getString("autodiscover.NeedEmail", "eas"))
            return;
        }

        this.validating = true;
        let error = "";

        //document.getElementById("tbsync.newaccount.wizard").canRewind = false;        
        document.getElementById("tbsync.error").hidden = true;
        document.getElementById("tbsync.newaccount.wizard").getButton("cancel").disabled = true;
        document.getElementById("tbsync.newaccount.wizard").getButton("finish").disabled = true;
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
            updateTimer.initWithCallback({ notify: function () { tbSyncEasNewAccount.updateAutodiscoverStatus() } }, 1000, 3);

            if (servertype == "office365") {
                let v2 = await eas.network.getServerConnectionViaAutodiscoverV2JsonRequest(
                    accountname,
                    user,
                    "https://autodiscover-s.outlook.com/autodiscover/autodiscover.json?Email=" + encodeURIComponent(user) + "&Protocol=ActiveSync",
                );
                let oauthData = eas.network.getOAuthObj({ host: v2.server, user, accountname, servertype });
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
                    url = v2.server;
                } else {
                    error = TbSync.getString("status.404", "eas");
                }
            } else {
                let result = await eas.network.getServerConnectionViaAutodiscover(
                    accountname,
                    user,
                    password,
                    tbSyncEasNewAccount.maxTimeout * 1000
                );
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
            await tbSyncEasNewAccount.addAccount(user, password, servertype, accountname, url);
        }

        //end validation
        document.getElementById("tbsync.newaccount.name").disabled = false;
        document.getElementById("tbsync.newaccount.user").disabled = false;
        document.getElementById("tbsync.newaccount.password").disabled = false;
        document.getElementById("tbsync.newaccount.servertype").disabled = false;
        document.getElementById("tbsync.newaccount.wizard").getButton("cancel").disabled = false;
        document.getElementById("tbsync.newaccount.wizard").getButton("finish").disabled = false;
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
        window.sizeToContent();
    },

    updateAutodiscoverStatus: function () {
        let offset = Math.round(((Date.now() - tbSyncEasNewAccount.startTime) / 1000));
        let timeout = (offset > 2) ? " (" + (tbSyncEasNewAccount.maxTimeout - offset) + ")" : "";

        document.getElementById('tbsync.newaccount.autodiscoverstatus').value = TbSync.getString("autodiscover.Querying", "eas") + timeout;
    },

    async addAccount(user, password, servertype, accountname, url) {
        let newAccountEntry = this.providerData.getDefaultAccountEntries();
        newAccountEntry.user = user;
        newAccountEntry.servertype = servertype;

        if (url) {
            //if no protocoll is given, prepend "https://"
            if (url.substring(0, 4) != "http" || url.indexOf("://") == -1) url = "https://" + url.split("://").join("/");
            newAccountEntry.host = eas.network.stripAutodiscoverUrl(url);
            newAccountEntry.https = (url.substring(0, 5) == "https");
        }

        // Add the new account.
        let newAccountData = this.providerData.addAccount(accountname, newAccountEntry);
        await eas.network.getAuthData(newAccountData).updateLoginData(user, password);

        window.close();
    }
};
