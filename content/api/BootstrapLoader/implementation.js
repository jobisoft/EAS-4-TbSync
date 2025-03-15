/*
 * This file is provided by the addon-developer-support repository at
 * https://github.com/thundernest/addon-developer-support
 *
 * Version: 1.21
 *
 * Author: John Bieling (john@thunderbird.net)
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

// Get various parts of the WebExtension framework that we need.
var { ExtensionCommon } = ChromeUtils.importESModule("resource://gre/modules/ExtensionCommon.sys.mjs");
var { ExtensionSupport } = ChromeUtils.importESModule("resource:///modules/ExtensionSupport.sys.mjs");
var { AddonManager } = ChromeUtils.importESModule("resource://gre/modules/AddonManager.sys.mjs");

function getMessenger(context) {
  let apis = ["storage", "runtime", "extension", "i18n"];

  function getStorage() {
    let localstorage = null;
    try {
      localstorage = context.apiCan.findAPIPath("storage");
      localstorage.local.get = (...args) =>
        localstorage.local.callMethodInParentProcess("get", args);
      localstorage.local.set = (...args) =>
        localstorage.local.callMethodInParentProcess("set", args);
      localstorage.local.remove = (...args) =>
        localstorage.local.callMethodInParentProcess("remove", args);
      localstorage.local.clear = (...args) =>
        localstorage.local.callMethodInParentProcess("clear", args);
    } catch (e) {
      console.info("Storage permission is missing");
    }
    return localstorage;
  }

  let messenger = {};
  for (let api of apis) {
    switch (api) {
      case "storage":
        ChromeUtils.defineLazyGetter(messenger, "storage", () =>
          getStorage()
        );
        break;

      default:
        ChromeUtils.defineLazyGetter(messenger, api, () =>
          context.apiCan.findAPIPath(api)
        );
    }
  }
  return messenger;
}

// Removed all extra code for backward compatibility for better maintainability.
var BootstrapLoader = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    this.uniqueRandomID = "AddOnNS" + context.extension.instanceId;
    this.menu_addonPrefs_id = "addonPrefs";


    this.pathToBootstrapScript = null;
    this.chromeHandle = null;
    this.chromeData = null;
    this.bootstrappedObj = {};

    // make the extension object and the messenger object available inside
    // the bootstrapped scope
    this.bootstrappedObj.extension = context.extension;
    this.bootstrappedObj.messenger = getMessenger(this.context);

    this.BOOTSTRAP_REASONS = {
      APP_STARTUP: 1,
      APP_SHUTDOWN: 2,
      ADDON_ENABLE: 3,
      ADDON_DISABLE: 4,
      ADDON_INSTALL: 5,
      ADDON_UNINSTALL: 6, // not supported
      ADDON_UPGRADE: 7,
      ADDON_DOWNGRADE: 8,
    };

    const aomStartup = Cc["@mozilla.org/addons/addon-manager-startup;1"].getService(Ci.amIAddonManagerStartup);
    const resProto = Cc["@mozilla.org/network/protocol;1?name=resource"].getService(Ci.nsISubstitutingProtocolHandler);

    let self = this;

    // TabMonitor to detect opening of tabs, to setup the options button in the add-on manager.
    this.tabMonitor = {
      onTabTitleChanged(tab) { },
      onTabClosing(tab) { },
      onTabPersist(tab) { },
      onTabRestored(tab) { },
      onTabSwitched(aNewTab, aOldTab) { },
      async onTabOpened(tab) {
        if (tab.browser && tab.mode.name == "contentTab") {
          let { setTimeout } = Services.wm.getMostRecentWindow("mail:3pane");
          // Instead of registering a load observer, wait until its loaded. Not nice,
          // but gets aroud a lot of edge cases.
          while (!tab.pageLoaded) {
            await new Promise(r => setTimeout(r, 150));
          }
          self.setupAddonManager(self.getAddonManagerFromTab(tab));
        }
      },
    };

    return {
      BootstrapLoader: {
        registerChromeUrl(data) {
          let chromeData = [];
          for (let entry of data) {
            chromeData.push(entry)
          }

          if (chromeData.length > 0) {
            const manifestURI = Services.io.newURI(
              "manifest.json",
              null,
              context.extension.rootURI
            );
            self.chromeHandle = aomStartup.registerChrome(manifestURI, chromeData);
          }

          self.chromeData = chromeData;
        },

        registerBootstrapScript: async function (aPath) {
          self.pathToBootstrapScript = aPath.startsWith("chrome://")
            ? aPath
            : context.extension.rootURI.resolve(aPath);

          // Get the addon object belonging to this extension.
          let addon = await AddonManager.getAddonByID(context.extension.id);
          console.log(addon.id);
          //make the addon globally available in the bootstrapped scope
          self.bootstrappedObj.addon = addon;

          // add BOOTSTRAP_REASONS to scope
          for (let reason of Object.keys(self.BOOTSTRAP_REASONS)) {
            self.bootstrappedObj[reason] = self.BOOTSTRAP_REASONS[reason];
          }

          // Load registered bootstrap scripts and execute its startup() function.
          try {
            if (self.pathToBootstrapScript) Services.scriptloader.loadSubScript(self.pathToBootstrapScript, self.bootstrappedObj, "UTF-8");
            if (self.bootstrappedObj.startup) self.bootstrappedObj.startup.call(self.bootstrappedObj, self.extension.addonData, self.BOOTSTRAP_REASONS[self.extension.startupReason]);
          } catch (e) {
            Components.utils.reportError(e)
          }

        }
      }
    };
  }

  onShutdown(isAppShutdown) {
    if (isAppShutdown) {
      return; // the application gets unloaded anyway
    }


    // Execute registered shutdown()
    try {
      if (this.bootstrappedObj.shutdown) {
        this.bootstrappedObj.shutdown(
          this.extension.addonData,
          isAppShutdown
            ? this.BOOTSTRAP_REASONS.APP_SHUTDOWN
            : this.BOOTSTRAP_REASONS.ADDON_DISABLE);
      }
    } catch (e) {
      Components.utils.reportError(e)
    }

    if (this.chromeHandle) {
      this.chromeHandle.destruct();
      this.chromeHandle = null;
    }
    // Flush all caches
    Services.obs.notifyObservers(null, "startupcache-invalidate");
    console.log("BootstrapLoader for " + this.extension.id + " unloaded!");
  }
};