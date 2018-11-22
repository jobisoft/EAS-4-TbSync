/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */

//no need to create namespace, we are in a sandbox

Components.utils.import("resource://gre/modules/Services.jsm");
Components.utils.import("resource://gre/modules/Task.jsm");

let thisID = "";

let onInitDoneObserver = {
    observe: Task.async (function* (aSubject, aTopic, aData) {        
        //it is now safe to import tbsync.jsm
        Components.utils.import("chrome://tbsync/content/tbsync.jsm");
        
        //load all providers of this provider Add-on into TbSync (one at a time, obey order)
       try {
            yield tbSync.loadProvider(thisID, "eas", "//eas4tbsync/content/provider/eas/eas.js");
        } catch (e) {}
    })
}

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
    
    thisID = data.id;
    Services.obs.addObserver(onInitDoneObserver, "tbsync.init.done", false);
    
    //during app startup, the load of the provider will be triggered by a "tbsync.init.done" notification, 
    //if load happens later, we need load manually 
    if (reason != APP_STARTUP) {
        onInitDoneObserver.observe();
    }    
}

function shutdown(data, reason) {
    Services.obs.removeObserver(onInitDoneObserver, "tbsync.init.done");

    //unload this provider Add-On and all its loaded providers from TbSync
    try {
        tbSync.unloadProviderAddon(data.id);
    } catch (e) {}
    Services.obs.notifyObservers(null, "chrome-flush-caches", null);    
}
