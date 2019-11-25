/* This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this file,
 * You can obtain one at http://mozilla.org/MPL/2.0/. */

/**
 * Provides OAuth 2.0 authentication.
 * @see RFC 6749
 */
var EXPORTED_SYMBOLS = ["OAuth2_1"];

const { Services } = ChromeUtils.import("resource://gre/modules/Services.jsm");
const { Log4Moz } = ChromeUtils.import("resource:///modules/gloda/log4moz.js");

Cu.importGlobalProperties(["fetch"]);

// Only allow one connecting window per endpoint.
var gConnecting = {};

function OAuth2_1(aBaseURI, aScope, aAppKey, aAppSecret) {
  this.authURI = aBaseURI + "oauth2/auth";
  this.tokenURI = aBaseURI + "oauth2/token";
  this.consumerKey = aAppKey;
  this.consumerSecret = aAppSecret;
  this.scope = aScope;
  this.extraAuthParams = [];

  this.log = Log4Moz.getConfiguredLogger("TBOAuth");
}

OAuth2_1.prototype = {
  consumerKey: null,
  consumerSecret: null,
  completionURI: "http://localhost",
  requestWindowURI: "chrome://messenger/content/browserRequest.xul",
  requestWindowFeatures: "chrome,private,centerscreen,width=980,height=750",
  requestWindowTitle: "",
  scope: null,

  accessToken: null,
  refreshToken: null,
  tokenExpires: 0,

  connect(aSuccess, aFailure, aWithUI, aRefresh) {
    this.connectSuccessCallback = aSuccess;
    this.connectFailureCallback = aFailure;

    if (!aRefresh && this.accessToken) {
      aSuccess();
    } else if (this.refreshToken) {
      this.requestAccessToken(this.refreshToken, true);
    } else {
      if (!aWithUI) {
        aFailure('{ "error": "auth_noui" }');
        return;
      }
      if (gConnecting[this.authURI]) {
        aFailure("Window already open");
        return;
      }
      this.requestAuthorization();
    }
  },

  requestAuthorization() {
    let params = new URLSearchParams({
      response_type: "code",
      client_id: this.consumerKey,
      redirect_uri: this.completionURI,
    });

    // The scope is optional.
    if (this.scope) {
      params.append("scope", this.scope);
    }

    for (let [name, value] of this.extraAuthParams) {
      params.append(name, value);
    }

    let authEndpointURI = this.authURI + "?" + params.toString();
    this.log.info(
      "Interacting with the resource owner to obtain an authorization grant " +
        "from the authorization endpoint: " +
        authEndpointURI
    );

    this._browserRequest = {
      account: this,
      url: authEndpointURI,
      _active: true,
      iconURI: "",
      cancelled() {
        if (!this._active) {
          return;
        }

        this.account.finishAuthorizationRequest();
        this.account.onAuthorizationFailed(
          Cr.NS_ERROR_ABORT,
          '{ "error": "cancelled"}'
        );
      },

      loaded(aWindow, aWebProgress) {
        if (!this._active) {
          return;
        }

        this._listener = {
          window: aWindow,
          webProgress: aWebProgress,
          _parent: this.account,

          QueryInterface: ChromeUtils.generateQI([
            Ci.nsIWebProgressListener,
            Ci.nsISupportsWeakReference,
          ]),

          _cleanUp() {
            this.webProgress.removeProgressListener(this);
            this.window.close();
            delete this.window;
          },

          _checkForRedirect(aURL) {
            if (aURL.indexOf(this._parent.completionURI) != 0) {
              return;
            }

            this._parent.finishAuthorizationRequest();
            this._parent.onAuthorizationReceived(aURL);
          },

          onStateChange(aWebProgress, aRequest, aStateFlags, aStatus) {
            const wpl = Ci.nsIWebProgressListener;
            if (aStateFlags & (wpl.STATE_START | wpl.STATE_IS_NETWORK)) {
              this._checkForRedirect(aRequest.name);
            }
          },
          onLocationChange(aWebProgress, aRequest, aLocation) {
            this._checkForRedirect(aLocation.spec);
          },
          onProgressChange() {},
          onStatusChange() {},
          onSecurityChange() {},
        };
        aWebProgress.addProgressListener(
          this._listener,
          Ci.nsIWebProgress.NOTIFY_ALL
        );
        aWindow.document.title = this.account.requestWindowTitle;
      },
    };

    this.wrappedJSObject = this._browserRequest;
    gConnecting[this.authURI] = true;
    Services.ww.openWindow(
      null,
      this.requestWindowURI,
      null,
      this.requestWindowFeatures,
      this
    );
  },
  finishAuthorizationRequest() {
    gConnecting[this.authURI] = false;
    if (!("_browserRequest" in this)) {
      return;
    }

    this._browserRequest._active = false;
    if ("_listener" in this._browserRequest) {
      this._browserRequest._listener._cleanUp();
    }
    delete this._browserRequest;
  },

  // @see RFC 6749 section 4.1.2: Authorization Response
  onAuthorizationReceived(aURL) {
    this.log.info("OAuth2_1 authorization received: url=" + aURL);
    let params = new URLSearchParams(aURL.split("?", 2)[1]);
    if (params.has("code")) {
      this.requestAccessToken(params.get("code"), false);
    } else {
      this.onAuthorizationFailed(null, aURL);
    }
  },

  onAuthorizationFailed(aError, aData) {
    this.connectFailureCallback(aData);
  },

  /**
   * Request a new access token, or refresh an existing one.
   * @param {string} aCode - The token issued to the client.
   * @param {boolean} aRefresh - Whether it's a refresh of a token or not.
   */
  requestAccessToken(aCode, aRefresh) {
    // @see RFC 6749 section 4.1.3. Access Token Request
    // @see RFC 6749 section 6. Refreshing an Access Token

    let data = new URLSearchParams();
    data.append("client_id", this.consumerKey);
    if (this.consumerSecret) data.append("client_secret", this.consumerSecret);

    if (aRefresh) {
      this.log.info(
        `Making a refresh request to the token endpoint: ${this.tokenURI}`
      );
      data.append("grant_type", "refresh_token");
      data.append("refresh_token", aCode);
    } else {
      this.log.info(
        `Making access token request to the token endpoint: ${this.tokenURI}`
      );
      data.append("grant_type", "authorization_code");
      data.append("code", aCode);
      data.append("redirect_uri", this.completionURI);
    }

    fetch(this.tokenURI, {
      method: "POST",
      cache: "no-cache",
      body: data,
    })
      .then(response => response.json())
      .then(result => {
        if ("error" in result) {
          // RFC 6749 section 5.2. Error Response
          this.log.info(
            `The authorization server returned an error response: ${JSON.stringify(
              result
            )}`
          );
          this.connectFailureCallback(result);
          return;
        }

        // RFC 6749 section 5.1. Successful Response
        this.log.info("The authorization server issued an access token.");
        this.accessToken = result.access_token;
        if ("refresh_token" in result) {
          this.refreshToken = result.refresh_token;
        }
        if ("expires_in" in result) {
          this.tokenExpires = new Date().getTime() + result.expires_in * 1000;
        } else {
          this.tokenExpires = Number.MAX_VALUE;
        }
        this.connectSuccessCallback();
      })
      .catch(err => {
        this.log.info(`Connection to authorization server failed: ${err}`);
        this.connectFailureCallback(err);
      });
  },
};
