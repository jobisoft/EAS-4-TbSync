/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */

//no need to create namespace, we are in a sandbox

Components.utils.import("resource://gre/modules/Services.jsm");

function install(data, reason) {
}

function uninstall(data, reason) {
}

function startup(data, reason) {
    //possible reasons: APP_STARTUP, ADDON_ENABLE, ADDON_INSTALL, ADDON_UPGRADE, or ADDON_DOWNGRADE.

    //set default prefs
    let branch = Services.prefs.getDefaultBranch("extensions.tbsync.");
    branch.setIntPref("eas.synclimit", 7);
    branch.setIntPref("eas.maxitems", 50);
    branch.setCharPref("eas.clientID.type", "TbSync");
    branch.setCharPref("eas.clientID.useragent", "Thunderbird ActiveSync");    
    branch.setBoolPref("eas.fix4freedriven", false);
    
    //during APP_STARTUP, TbSync will find auto load all active providers, if this provider gets enabled later, load it dynamically 
    if (reason != APP_STARTUP) {
        Services.obs.notifyObservers(null, "tbsync.registerProvider", JSON.stringify({provider: "eas", js: "//eas4tbsync/content/eas.js"}));
    }
}

function shutdown(data, reason) {
    //unload this provider from TbSync
    Services.obs.notifyObservers(null, "tbsync.unregisterProvider", "eas");
    Services.obs.notifyObservers(null, "chrome-flush-caches", null);
}
