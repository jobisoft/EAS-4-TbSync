/* global Services, ExtensionCommon */

"use strict";

var LegacyLoginManager = class extends ExtensionCommon.ExtensionAPI {
  getAPI(_context) {
    return {
      LegacyLoginManager: {
        /** Mirrors the legacy TbSync `passwordManager.getLoginInfo` call
         *  (TbSync/content/modules/passwordManager.js) which passes the bare
         *  hostname as the login origin and matches on httpRealm + username.
         *  Returns the password string or null when no entry matches. */
        getLoginInfo: async ({ origin, httpRealm, username }) => {
          const logins = Services.logins.findLogins(origin, null, httpRealm);
          for (const login of logins) {
            if (login.username === username) return login.password;
          }
          return null;
        },
      },
    };
  }
};
