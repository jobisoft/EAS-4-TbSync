/*
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

"use strict";

// Using a closure to not leak anything but the API to the outside world.
(function (exports) {

  var EAS4TbSync = class extends ExtensionCommon.ExtensionAPI {
    async observeTbSyncInitialized (aSubject, aTopic, aData) {
      let valid = false;
      try {
        var { TbSync } = ChromeUtils.importESModule("chrome://tbsync/content/tbsync.sys.mjs");
        valid = TbSync.enabled;
      } catch (e) {
        // If this fails, TbSync is not loaded yet and we will get the notification later again.
      }

      //load this provider add-on into TbSync
      if (valid) {
        await TbSync.providers.loadProvider(this.extension, "eas", "chrome://eas4tbsync/content/provider.js");
      }
    }

    getAPI(context) {
      const that = this;

      return {
        EAS4TbSync: {
          async load() {
            Services.obs.addObserver(that.observeTbSyncInitialized, "tbsync.observer.initialized", false);

            // Did we miss the observer?
            that.observeTbSyncInitialized();
          }
        },
      };
    }

    onShutdown(isAppShutdown) {
      if (isAppShutdown) {
        return; // the application gets unloaded anyway
      }

      Services.obs.removeObserver(this.observeTbSyncInitialized, "tbsync.observer.initialized");
      //unload this provider add-on from TbSync
      try {
        var { TbSync } = ChromeUtils.importESModule("chrome://tbsync/content/tbsync.sys.mjs");
        TbSync.providers.unloadProvider("eas");
      } catch (e) {
        //if this fails, TbSync has been unloaded already and has unloaded this addon as well
      }
    }
  };
  exports.EAS4TbSync = EAS4TbSync;
})(this);