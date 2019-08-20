/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

var network = {  
    
    getEasURL: function(accountData) {
        let protocol = (accountData.getAccountProperty("https")) ? "https://" : "http://";
        let h = protocol + accountData.getAccountProperty("host"); 
        while (h.endsWith("/")) { h = h.slice(0,-1); }

        if (h.endsWith("Microsoft-Server-ActiveSync")) return h;
        return h + "/Microsoft-Server-ActiveSync"; 
    },
    
    getAuthData: function(accountData) {
        let authData = {
            // This is the host for the password manager, which could be different from
            // the actual host property of the account. For EAS we want to couple the password
            // with the ACCOUNT and not any sort of url, which could change via autodiscover
            // at any time.
            get host() { 
                return "TbSync#" + accountData.accountID;
            },

            get user() {
                return accountData.getAccountProperty("user");
            },

            get password() {
                return tbSync.passwordManager.getLoginInfo(this.host, "TbSync/EAS", this.user);
            },

            updateLoginData: function(newUsername, newPassword) {
                let oldUsername = this.user;
                tbSync.passwordManager.updateLoginInfo(this.host, "TbSync/EAS", oldUsername, newUsername, newPassword);
                // Also update the username of this account. Add dedicated username setter?
                accountData.setAccountProperty("user", newUsername);
            },          
        };
        return authData;
    },  










    sendRequest: async function (wbxml, command, syncData, allowSoftFail = false) {
        let ALLOWED_RETRIES = {
            PasswordPrompt : 3,
            NetworkError : 1,
        }
        
        let rv = {};
        for (;;) {

            if (rv.errorType) {                
                let retry = false;
                
                if (ALLOWED_RETRIES[rv.errorType] > 0) {
                    ALLOWED_RETRIES[rv.errorType]--;
                    
                    switch (rv.errorType) {
                        
                        case "PasswordPrompt": 
                        {
                            let authData = eas.network.getAuthData(syncData.accountData);
                            let promptData = {
                                windowID: "auth:" + syncData.accountData.accountID,
                                accountname: syncData.accountData.getAccountProperty("accountname"),
                                usernameLocked: syncData.accountData.isConnected(),
                                username: authData.user
                            }
                            
                            let syncState = syncData.getSyncState(); 
                            syncData.setSyncState("passwordprompt");
                            let credentials = await tbSync.passwordManager.asyncPasswordPrompt(promptData, eas.openWindows);
                            if (credentials) {
                                // Update login data and try again.
                                authData.updateLoginData(credentials.username, credentials.password);
                                syncData.setSyncState(syncState);
                                retry = true;
                            }
                        }
                        break;
                        
                        case "NetworkError":
                        {
                            // Could not connect to server. Can we rerun autodiscover?       
                            if (syncData.accountData.getAccountProperty( "servertype") == "auto") {
                                let errorcode = await eas.network.updateServerConnectionViaAutodiscover(syncData);
                                console.log("ERR: " + errorcode);
                                if (errorcode == 200) {                       
                                    // autodiscover succeeded, retry with new data
                                    retry = true;                            
                                } else if (errorcode == 401) {
                                    // manipulate rv to run password prompt
                                    ALLOWED_RETRIES[rv.errorType]++;
                                    rv.errorType = "PasswordPrompt";
                                    rv.errorObj = eas.sync.finish("error", "401");
                                    continue; // with the next loop, skip connection to the server
                                }
                            }
                        }
                        break;
                        
                    }
                } 
                
                if (!retry) throw rv.errorObj;
            }
            
            rv = await this.sendRequestPromise(wbxml, command, syncData, allowSoftFail);
            
            if (rv.errorType) {
                // make sure, there is a valid ALLOWED_RETRIES setting for the returned error
                if (rv.errorType && !ALLOWED_RETRIES.hasOwnProperty(rv.errorType)) {
                    ALLOWED_RETRIES[rv.errorType] = 1;
                }
            } else {
                return rv;
            }
        }        
    },

    sendRequestPromise: function (wbxml, command, syncData, allowSoftFail = false) {
        let msg = "Sending data <" + syncData.getSyncState() + "> for " + syncData.accountData.getAccountProperty("accountname");
        if (syncData.currentFolderData) msg += " (" + syncData.currentFolderData.getFolderProperty("foldername") + ")";
        syncData.request = eas.network.logXML(wbxml, msg);
        syncData.response = "";

        let connection = eas.network.getAuthData(syncData.accountData);
        let userAgent = syncData.accountData.getAccountProperty("useragent"); //plus calendar.useragent.extra = Lightning/5.4.5.2
        let deviceType = syncData.accountData.getAccountProperty("devicetype");
        let deviceId = syncData.accountData.getAccountProperty("deviceId");

        tbSync.dump("Sending (EAS v"+syncData.accountData.getAccountProperty("asversion") +")", "POST " + eas.network.getEasURL(syncData.accountData) + '?Cmd=' + command + '&User=' + encodeURIComponent(connection.user) + '&DeviceType=' +deviceType + '&DeviceId=' + deviceId, true);
        
        return new Promise(function(resolve,reject) {
            // Create request handler - API changed with TB60 to new XMKHttpRequest()
            syncData.req = new XMLHttpRequest();
            syncData.req.mozBackgroundRequest = true;
            syncData.req.open("POST", eas.network.getEasURL(syncData.accountData) + '?Cmd=' + command + '&User=' + encodeURIComponent(connection.user) + '&DeviceType=' +encodeURIComponent(deviceType) + '&DeviceId=' + deviceId, true);
            syncData.req.overrideMimeType("text/plain");
            syncData.req.setRequestHeader("User-Agent", userAgent);
            syncData.req.setRequestHeader("Content-Type", "application/vnd.ms-sync.wbxml");
            syncData.req.setRequestHeader("Authorization", 'Basic ' + tbSync.tools.b64encode(connection.user + ':' + connection.password));
            if (syncData.accountData.getAccountProperty("asversion") == "2.5") {
                syncData.req.setRequestHeader("MS-ASProtocolVersion", "2.5");
            } else {
                syncData.req.setRequestHeader("MS-ASProtocolVersion", "14.0");
            }
            syncData.req.setRequestHeader("Content-Length", wbxml.length);
            if (syncData.accountData.getAccountProperty("provision")) {
                syncData.req.setRequestHeader("X-MS-PolicyKey", syncData.accountData.getAccountProperty("policykey"));
                tbSync.dump("PolicyKey used", syncData.accountData.getAccountProperty("policykey"));
            }

            syncData.req.timeout = eas.Base.getConnectionTimeout();

            syncData.req.ontimeout = function () {
                if (allowSoftFail) {
                    resolve("");
                } else {
                    reject(eas.sync.finish("error", "timeout"));
                }
            };

            syncData.req.onerror = function () {
                if (allowSoftFail) {
                    resolve("");
                } else {
                    let error = tbSync.network.createTCPErrorFromFailedXHR(syncData.req) || "networkerror";
                    let rv = {};
                    rv.errorObj = eas.sync.finish("error", error);
                    rv.errorType = "NetworkError";
                    resolve(rv);
                }
            };

            syncData.req.onload = function() {
                let response = syncData.req.responseText;
                switch(syncData.req.status) {

                    case 200: //OK
                        let msg = "Receiving data <" + syncData.getSyncState() + "> for " + syncData.accountData.getAccountProperty("accountname");
                        if (syncData.currentFolderData) msg += " (" + syncData.currentFolderData.getFolderProperty("foldername") + ")";
                        syncData.response = eas.network.logXML(response, msg);

                        //What to do on error? IS this an error? Yes!
                        if (!allowSoftFail && response.length !== 0 && response.substr(0, 4) !== String.fromCharCode(0x03, 0x01, 0x6A, 0x00)) {
                            tbSync.dump("Recieved Data", "Expecting WBXML but got junk (request status = " + syncData.req.status + ", ready state = " + syncData.req.readyState + "\n>>>>>>>>>>\n" + response + "\n<<<<<<<<<<\n");
                            reject(eas.sync.finish("warning", "invalid"));
                        } else {
                            resolve(response);
                        }
                        break;

                    case 401: // AuthError
                    case 403: // Forbiddden (some servers send forbidden on AuthError, like Freenet)
                        let rv = {};
                        rv.errorObj = eas.sync.finish("error", "401");
                        rv.errorType = "PasswordPrompt";
                        resolve(rv);
                        break;

                    case 449: // Request for new provision (enable it if needed)
                        //enable provision
                        syncData.accountData.setAccountProperty("provision", true);
                        syncData.accountData.resetAccountProperty("policykey");
                        reject(eas.sync.finish("resyncAccount", syncData.req.status));
                        break;

                    case 451: // Redirect - update host and login manager 
                        let header = syncData.req.getResponseHeader("X-MS-Location");
                        let newHost = header.slice(header.indexOf("://") + 3, header.indexOf("/M"));

                        tbSync.dump("redirect (451)", "header: " + header + ", oldHost: " +syncData.accountData.getAccountProperty("host") + ", newHost: " + newHost);

                        syncData.accountData.setAccountProperty("host", newHost);
                        reject(eas.sync.finish("resyncAccount", syncData.req.status));
                        break;
                        
                    default:
                        if (allowSoftFail) {
                            resolve("");
                        } else {
                            reject(eas.sync.finish("error", "httperror::" + syncData.req.status));
                        }
                }
            };

            syncData.req.send(wbxml);
            
        });
    },










    // RESPONSE EVALUATION
    
    logXML : function (wbxml, what) {
        let rawxml = eas.wbxmltools.convert2xml(wbxml);
        let xml = null;
        if (rawxml)  {
            xml = rawxml.split('><').join('>\n<');
        }
        
        //include xml in log, if userdatalevel 2 or greater
        if ((tbSync.prefs.getBoolPref("log.toconsole") || tbSync.prefs.getBoolPref("log.tofile")) && tbSync.prefs.getIntPref("log.userdatalevel")>1) {

            //log raw wbxml if userdatalevel is 3 or greater
            if (tbSync.prefs.getIntPref("log.userdatalevel")>2) {
                let charcodes = [];
                for (let i=0; i< wbxml.length; i++) charcodes.push(wbxml.charCodeAt(i).toString(16));
                let bytestring = charcodes.join(" ");
                tbSync.dump("WBXML: " + what, "\n" + bytestring);
            }

            if (xml) {
                //raw xml is save xml with all special chars in user data encoded by encodeURIComponent - KEEP that in order to be able to analyze logged XML 
                //let xml = decodeURIComponent(rawxml.split('><').join('>\n<'));
                //if userdatalevel is 3 or greater print everything, otherwise exclude application data
                if (tbSync.prefs.getIntPref("log.userdatalevel")<3) {
                    let rx = new RegExp("<ApplicationData[\\d\\D]*?\/ApplicationData>", "g");
                    tbSync.dump("XML: " + what, "\n" + xml.replace(rx, ""));
                } else {
                    tbSync.dump("XML: " + what, "\n" + xml);
                }
            } else {
                tbSync.dump("XML: " + what, "\nFailed to convert WBXML to XML!\n");
            }
        }
    
    return xml;
    },
    
    //returns false on parse error and null on empty response (if allowed)
    getDataFromResponse: function (wbxml, allowEmptyResponse = !eas.flags.allowEmptyResponse) {        
        //check for empty wbxml
        if (wbxml.length === 0) {
            if (allowEmptyResponse) return null;
            else throw eas.sync.finish("warning", "empty-response");
        }

        //convert to save xml (all special chars in user data encoded by encodeURIComponent) and check for parse errors
        let xml = eas.wbxmltools.convert2xml(wbxml);
        if (xml === false) {
            throw eas.sync.finish("warning", "wbxml-parse-error");
        }
        
        //retrieve data and check for empty data (all returned data fields are already decoded by decodeURIComponent)
        let wbxmlData = eas.xmltools.getDataFromXMLString(xml);
        if (wbxmlData === null) {
            if (allowEmptyResponse) return null;
            else throw eas.sync.finish("warning", "response-contains-no-data");
        }
        
        //debug
        eas.xmltools.printXmlData(wbxmlData, false); //do not include ApplicationData in log
        return wbxmlData;
    },  
  
    updateSynckey: function (syncData, wbxmlData) {
        let synckey = eas.xmltools.getWbxmlDataField(wbxmlData,"Sync.Collections.Collection.SyncKey");

        if (synckey) {
            // This COULD be a cause of problems... 
            syncData.synckey = synckey;
            syncData.currentFolderData.setFolderProperty("synckey", synckey);
        } else {
            throw eas.sync.finish("error", "wbxmlmissingfield::Sync.Collections.Collection.SyncKey");
        }
    },

    checkStatus : function (syncData, wbxmlData, path, rootpath="", allowSoftFail = false) {
        //path is relative to wbxmlData
        //rootpath is the absolute path and must be specified, if wbxml is not the root node and thus path is not the rootpath	    
        let status = eas.xmltools.getWbxmlDataField(wbxmlData,path);
        let fullpath = (rootpath=="") ? path : rootpath;
        let elements = fullpath.split(".");
        let type = elements[0];

        //check if fallback to main class status: the answer could just be a "Sync.Status" instead of a "Sync.Collections.Collections.Status"
        if (status === false) {
            let mainStatus = eas.xmltools.getWbxmlDataField(wbxmlData, type + "." + elements[elements.length-1]);
            if (mainStatus === false) {
                //both possible status fields are missing, abort
                throw eas.sync.finish("warning", "wbxmlmissingfield::" + fullpath, "Request:\n" + syncData.request + "\n\nResponse:\n" + syncData.response);
            } else {
                //the alternative status could be extracted
                status = mainStatus;
                fullpath = type + "." + elements[elements.length-1];
            }
        }

        //check if all is fine (not bad)
        if (status == "1") {
            return "";
        }

        tbSync.dump("wbxml status check", type + ": " + fullpath + " = " + status);

        //handle errrors based on type
        let statusType = type+"."+status;
        switch (statusType) {
            case "Sync.3": /*
                        MUST return to SyncKey element value of 0 for the collection. The client SHOULD either delete any items that were added 
                        since the last successful Sync or the client MUST add those items back to the server after completing the full resynchronization
                        */
                tbSync.eventlog.add("warning", syncData.eventLogInfo, "Forced Folder Resync", "Request:\n" + syncData.request + "\n\nResponse:\n" + syncData.response);
                syncData.currentFolderData.remove();
                throw eas.sync.finish("resyncFolder", statusType);
            
            case "Sync.4": //Malformed request
            case "Sync.5": //Temporary server issues or invalid item
            case "Sync.6": //Invalid item
            case "Sync.8": //Object not found
                if (allowSoftFail) return statusType;
                throw eas.sync.finish("warning", statusType, "Request:\n" + syncData.request + "\n\nResponse:\n" + syncData.response);

            case "Sync.7": //The client has changed an item for which the conflict policy indicates that the server's changes take precedence.
            case "Sync.9": //User account could be out of disk space, also send if no write permission (TODO)
                return "";

            case "FolderDelete.3": // special system folder - fatal error
            case "FolderDelete.6": // error on server
                throw eas.sync.finish("warning", statusType, "Request:\n" + syncData.request + "\n\nResponse:\n" + syncData.response);

            case "FolderDelete.4": // folder does not exist - resync ( we allow delete only if folder is not subscribed )
            case "FolderDelete.9": // invalid synchronization key - resync
            case "FolderSync.9": // invalid synchronization key - resync
            case "Sync.12": // folder hierarchy changed
                {
                    let folders = syncData.accountData.getAllFoldersIncludingCache();
                    for (let folder of folders) {
                        folder.remove();
                    }		    
                    // reset account
                    eas.Base.onEnableAccount(syncData.accountData);
                    throw eas.sync.finish("resyncAccount", statusType, "Request:\n" + syncData.request + "\n\nResponse:\n" + syncData.response);
                }
        }
        
        //handle global error (https://msdn.microsoft.com/en-us/library/ee218647(v=exchg.80).aspx)
        let descriptions = {};
        switch(status) {
            case "101": //invalid content
            case "102": //invalid wbxml
            case "103": //invalid xml
                throw eas.sync.finish("error", "global." + status, "Request:\n" + syncData.request + "\n\nResponse:\n" + syncData.response);
            
            case "109": descriptions["109"]="DeviceTypeMissingOrInvalid";
            case "112": descriptions["112"]="ActiveDirectoryAccessDenied";
            case "126": descriptions["126"]="UserDisabledForSync";
            case "127": descriptions["127"]="UserOnNewMailboxCannotSync";
            case "128": descriptions["128"]="UserOnLegacyMailboxCannotSync";
            case "129": descriptions["129"]="DeviceIsBlockedForThisUser";
            case "130": descriptions["120"]="AccessDenied";
            case "131": descriptions["131"]="AccountDisabled";
                throw eas.sync.finish("error", "global.clientdenied"+ "::" + status + "::" + descriptions[status]);

            case "110": //server error - resync
                {
                    let folders = syncData.accountData.getAllFoldersIncludingCache();
                    for (let folder of folders) {
                        folder.remove();
                    }		    
                    // reset account
                    eas.Base.onEnableAccount(syncData.accountData);
                    throw eas.sync.finish("resyncAccount", statusType, "Request:\n" + syncData.request + "\n\nResponse:\n" + syncData.response);
                }
                
            case "141": // The device is not provisionable
            case "142": // DeviceNotProvisioned
            case "143": // PolicyRefresh
            case "144": // InvalidPolicyKey
                //enable provision
                syncData.accountData.setAccountProperty("provision", true);
                syncData.accountData.resetAccountProperty("policykey");
                throw eas.sync.finish("resyncAccount", statusType);
            
            default:
                if (allowSoftFail) return statusType;
                throw eas.sync.finish("error", statusType, "Request:\n" + syncData.request + "\n\nResponse:\n" + syncData.response);

        }		
    },
    









    // WBXML COMM STUFF
    
    setDeviceInformation: async function (syncData)  {
        if (syncData.accountData.getAccountProperty("asversion") == "2.5" || !syncData.accountData.getAccountProperty("allowedEasCommands").split(",").includes("Settings")) {
            return;
        }
            
        syncData.setSyncState("prepare.request.setdeviceinfo");

        let wbxml = wbxmltools.createWBXML();
        wbxml.switchpage("Settings");
        wbxml.otag("Settings");
            wbxml.otag("DeviceInformation");
                wbxml.otag("Set");
                    wbxml.atag("Model", "Computer");
                    wbxml.atag("FriendlyName", "TbSync on Device " + syncData.accountData.getAccountProperty("deviceId").substring(4));
                    wbxml.atag("OS", OS.Constants.Sys.Name);
                    wbxml.atag("UserAgent", syncData.accountData.getAccountProperty("useragent"));
                wbxml.ctag();
            wbxml.ctag();
        wbxml.ctag();

        syncData.setSyncState("send.request.setdeviceinfo");
        let response = await eas.network.sendRequest(wbxml.getBytes(), "Settings", syncData);

        syncData.setSyncState("eval.response.setdeviceinfo");
        let wbxmlData = eas.network.getDataFromResponse(response);

        eas.network.checkStatus(syncData, wbxmlData,"Settings.Status");
    },

    getPolicykey: async function (syncData)  {
        //build WBXML to request provision
       syncData.setSyncState("prepare.request.provision");
        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.switchpage("Provision");
        wbxml.otag("Provision");
            wbxml.otag("Policies");
                wbxml.otag("Policy");
                    wbxml.atag("PolicyType", (syncData.accountData.getAccountProperty("asversion") == "2.5") ? "MS-WAP-Provisioning-XML" : "MS-EAS-Provisioning-WBXML" );
                wbxml.ctag();
            wbxml.ctag();
        wbxml.ctag();

        for (let loop=0; loop < 2; loop++) {
           syncData.setSyncState("send.request.provision");
            let response = await eas.network.sendRequest(wbxml.getBytes(), "Provision", syncData);

            syncData.setSyncState("eval.response.provision");
            let wbxmlData = eas.network.getDataFromResponse(response);
            let policyStatus = eas.xmltools.getWbxmlDataField(wbxmlData, "Provision.Policies.Policy.Status");
            let provisionStatus = eas.xmltools.getWbxmlDataField(wbxmlData, "Provision.Status");
            if (provisionStatus === false) {
                throw eas.sync.finish("error", "wbxmlmissingfield::Provision.Status");
            } else if (provisionStatus != "1") {
                //dump policy status as well
                if (policyStatus) tbSync.dump("PolicyKey","Received policy status: " + policyStatus);
                throw eas.sync.finish("error", "provision::" + provisionStatus);
            }

            //reaching this point: provision status was ok
            let policykey = eas.xmltools.getWbxmlDataField(wbxmlData,"Provision.Policies.Policy.PolicyKey");
            switch (policyStatus) {
                case false:
                    throw eas.sync.finish("error", "wbxmlmissingfield::Provision.Policies.Policy.Status");

                case "2":
                    //server does not have a policy for this device: disable provisioning
                    syncData.accountData.setAccountProperty("provision", false)
                    syncData.accountData.resetAccountProperty("policykey");
                    throw eas.sync.finish("resyncAccount", "NoPolicyForThisDevice");

                case "1":
                    if (policykey === false) {
                        throw eas.sync.finish("error", "wbxmlmissingfield::Provision.Policies.Policy.PolicyKey");
                    } 
                    tbSync.dump("PolicyKey","Received policykey (" + loop + "): " + policykey);
                    syncData.accountData.setAccountProperty("policykey", policykey);
                    break;

                default:
                    throw eas.sync.finish("error", "policy." + policyStatus);
            }

            //build WBXML to acknowledge provision
           syncData.setSyncState("prepare.request.provision");
            wbxml = eas.wbxmltools.createWBXML();
            wbxml.switchpage("Provision");
            wbxml.otag("Provision");
                wbxml.otag("Policies");
                    wbxml.otag("Policy");
                        wbxml.atag("PolicyType",(syncData.accountData.getAccountProperty("asversion") == "2.5") ? "MS-WAP-Provisioning-XML" : "MS-EAS-Provisioning-WBXML" );
                        wbxml.atag("PolicyKey", policykey);
                        wbxml.atag("Status", "1");
                    wbxml.ctag();
                wbxml.ctag();
            wbxml.ctag();
            
            //this wbxml will be used by Send at the top of this loop
        }
    },

    getSynckey: async function (syncData) {
        syncData.setSyncState("prepare.request.synckey");
        //build WBXML to request a new syncKey
        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.otag("Sync");
            wbxml.otag("Collections");
                wbxml.otag("Collection");
                    if (syncData.accountData.getAccountProperty("asversion") == "2.5") wbxml.atag("Class", syncData.type);
                    wbxml.atag("SyncKey","0");
                    wbxml.atag("CollectionId", syncData.currentFolderData.getFolderProperty("serverID"));
                wbxml.ctag();
            wbxml.ctag();
        wbxml.ctag();
        
        syncData.setSyncState("send.request.synckey");
        let response = await eas.network.sendRequest(wbxml.getBytes(), "Sync", syncData);

        syncData.setSyncState("eval.response.synckey");
        // get data from wbxml response
        let wbxmlData = eas.network.getDataFromResponse(response);
        //check status
        eas.network.checkStatus(syncData, wbxmlData,"Sync.Collections.Collection.Status");
        //update synckey
        eas.network.updateSynckey(syncData, wbxmlData);
    },

    getItemEstimate: async function (syncData)  {
        syncData.progressData.reset();
        
        if (!syncData.accountData.getAccountProperty("allowedEasCommands").split(",").includes("GetItemEstimate")) {
            return; //do not throw, this is optional
        }
        
        syncData.setSyncState("prepare.request.estimate");
        
        // BUILD WBXML
        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.switchpage("GetItemEstimate");
        wbxml.otag("GetItemEstimate");
            wbxml.otag("Collections");
                wbxml.otag("Collection");
                    if (syncData.accountData.getAccountProperty("asversion") == "2.5") { //got this order for 2.5 directly from Microsoft support
                        wbxml.atag("Class", syncData.type); //only 2.5
                        wbxml.atag("CollectionId", syncData.currentFolderData.getFolderProperty("serverID"));
                        wbxml.switchpage("AirSync");
                        // required !
                        // https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-ascmd/ffbefa62-e315-40b9-9cc6-f8d74b5f65d4
                        if (syncData.type == "Calendar") wbxml.atag("FilterType", syncData.currentFolderData.accountData.getAccountProperty("synclimit"));
                        else wbxml.atag("FilterType", "0"); // we may filter incomplete tasks
                        
                        wbxml.atag("SyncKey", syncData.synckey);
                        wbxml.switchpage("GetItemEstimate");
                    } else { //14.0
                        wbxml.switchpage("AirSync");
                        wbxml.atag("SyncKey", syncData.synckey);
                        wbxml.switchpage("GetItemEstimate");
                        wbxml.atag("CollectionId", syncData.currentFolderData.getFolderProperty("serverID"));
                        wbxml.switchpage("AirSync");
                        wbxml.otag("Options");
                            // optional
                            if (syncData.type == "Calendar") wbxml.atag("FilterType", syncData.currentFolderData.accountData.getAccountProperty("synclimit"));
                            wbxml.atag("Class", syncData.type);
                        wbxml.ctag();
                        wbxml.switchpage("GetItemEstimate");
                    }
                wbxml.ctag();
            wbxml.ctag();
        wbxml.ctag();

        //SEND REQUEST
        syncData.setSyncState("send.request.estimate");
        let response = await eas.network.sendRequest(wbxml.getBytes(), "GetItemEstimate", syncData, /* allowSoftFail */ true);

        //VALIDATE RESPONSE
        syncData.setSyncState("eval.response.estimate");

        // get data from wbxml response, some servers send empty response if there are no changes, which is not an error
        let wbxmlData = eas.network.getDataFromResponse(response, eas.flags.allowEmptyResponse);
        if (wbxmlData === null) return;

        let status = eas.xmltools.getWbxmlDataField(wbxmlData, "GetItemEstimate.Response.Status");
        let estimate = eas.xmltools.getWbxmlDataField(wbxmlData, "GetItemEstimate.Response.Collection.Estimate");

        if (status && status == "1") { //do not throw on error, with EAS v2.5 I get error 2 for tasks and calendars ???
            syncData.progressData.reset(0, estimate);
        }
    },

    getUserInfo: async function (syncData)  {
        if (!syncData.accountData.getAccountProperty("allowedEasCommands").split(",").includes("Settings")) {
            return;
        }

        syncData.setSyncState("prepare.request.getuserinfo");

        let wbxml = eas.wbxmltools.createWBXML();
        wbxml.switchpage("Settings");
        wbxml.otag("Settings");
            wbxml.otag("UserInformation");
                wbxml.atag("Get");
            wbxml.ctag();
        wbxml.ctag();

        syncData.setSyncState("send.request.getuserinfo");
        let response = await eas.network.sendRequest(wbxml.getBytes(), "Settings", syncData);


        syncData.setSyncState("eval.response.getuserinfo");
        let wbxmlData = eas.network.getDataFromResponse(response);

        eas.network.checkStatus(syncData, wbxmlData,"Settings.Status");
    },



    






    // SEARCH

    getSearchResults: async function (accountData, currentQuery) {

        let _wbxml = eas.wbxmltools.createWBXML();
        _wbxml.switchpage("Search");
        _wbxml.otag("Search");
            _wbxml.otag("Store");
                _wbxml.atag("Name", "GAL");
                _wbxml.atag("Query", currentQuery);
                _wbxml.otag("Options");
                    _wbxml.atag("Range", "0-99"); //Z-Push needs a Range
                    //Not valid for GAL: https://msdn.microsoft.com/en-us/library/gg675461(v=exchg.80).aspx
                    //_wbxml.atag("DeepTraversal");
                    //_wbxml.atag("RebuildResults");
                _wbxml.ctag();
            _wbxml.ctag();
        _wbxml.ctag();

        let wbxml = _wbxml.getBytes();
        
        eas.network.logXML(wbxml, "Send (GAL Search)");
        let command = "Search";
        
        let authData = eas.network.getAuthData(accountData);
        let userAgent = accountData.getAccountProperty("useragent"); //plus calendar.useragent.extra = Lightning/5.4.5.2
        let deviceType = accountData.getAccountProperty("devicetype");
        let deviceId = accountData.getAccountProperty("deviceId");

        tbSync.dump("Sending (EAS v" + accountData.getAccountProperty("asversion") +")", "POST " + eas.network.getEasURL(accountData) + '?Cmd=' + command + '&User=' + encodeURIComponent(authData.user) + '&DeviceType=' +deviceType + '&DeviceId=' + deviceId, true);
        
        try {
            let response = await new Promise(function(resolve, reject) {
                // Create request handler - API changed with TB60 to new XMKHttpRequest()
                let req = new XMLHttpRequest();
                req.mozBackgroundRequest = true;
                req.open("POST", eas.network.getEasURL(accountData) + '?Cmd=' + command + '&User=' + encodeURIComponent(authData.user) + '&DeviceType=' +encodeURIComponent(deviceType) + '&DeviceId=' + deviceId, true);
                req.overrideMimeType("text/plain");
                req.setRequestHeader("User-Agent", userAgent);
                req.setRequestHeader("Content-Type", "application/vnd.ms-sync.wbxml");
                req.setRequestHeader("Authorization", 'Basic ' + tbSync.tools.b64encode(authData.user + ':' + authData.password));
                if (accountData.getAccountProperty("asversion") == "2.5") {
                    req.setRequestHeader("MS-ASProtocolVersion", "2.5");
                } else {
                    req.setRequestHeader("MS-ASProtocolVersion", "14.0");
                }
                req.setRequestHeader("Content-Length", wbxml.length);
                if (accountData.getAccountProperty("provision")) {
                    req.setRequestHeader("X-MS-PolicyKey", accountData.getAccountProperty("policykey"));
                    tbSync.dump("PolicyKey used", accountData.getAccountProperty("policykey"));
                }

                req.timeout = eas.Base.getConnectionTimeout();

                req.ontimeout = function () {
                    reject("GAL Search timeout");
                };

                req.onerror = function () {
                    reject("GAL Search Error");
                };

                req.onload = function() {
                    let response = req.responseText;
                    
                    switch(req.status) {

                        case 200: //OK
                            eas.network.logXML(response, "Received (GAL Search");

                            //What to do on error? IS this an error? Yes!
                            if (response.length !== 0 && response.substr(0, 4) !== String.fromCharCode(0x03, 0x01, 0x6A, 0x00)) {
                                tbSync.dump("Recieved Data", "Expecting WBXML but got junk (request status = " + req.status + ", ready state = " + req.readyState + "\n>>>>>>>>>>\n" + response + "\n<<<<<<<<<<\n");
                                reject("GAL Search Response Invalid");
                            } else {
                                resolve(response);
                            }
                            break;
                          
                        default:
                            reject("GAL Search Failed: " + req.status);
                    }
                };

                req.send(wbxml);
                
            });
            return response;
        } catch (e) {
            Components.utils.reportError(e);
            return;
        }
    },










       // OPTIONS

    getServerOptions: async function (syncData) {        
        syncData.setSyncState("prepare.request.options");
        let authData = eas.network.getAuthData(syncData.accountData);

        let userAgent = syncData.accountData.getAccountProperty("useragent"); //plus calendar.useragent.extra = Lightning/5.4.5.2
        tbSync.dump("Sending", "OPTIONS " + eas.network.getEasURL(syncData.accountData));
        
        let allowedRetries = 5;
        let retry;
        do {
            retry = false;
            
            let result = await new Promise(function(resolve,reject) {
                syncData.req = new XMLHttpRequest();
                syncData.req.mozBackgroundRequest = true;
                syncData.req.open("OPTIONS", eas.network.getEasURL(syncData.accountData), true);
                syncData.req.overrideMimeType("text/plain");
                syncData.req.setRequestHeader("User-Agent", userAgent);            
                syncData.req.setRequestHeader("Authorization", 'Basic ' + tbSync.tools.b64encode(authData.user + ':' + authData.password));
                syncData.req.timeout = eas.Base.getConnectionTimeout();

                syncData.req.ontimeout = function () {
                    resolve();
                };

                syncData.req.onerror = function () {
                    resolve();
                };

                syncData.req.onload = function() {
                    syncData.setSyncState("eval.request.options");
                    let responseData = {};

                    switch(syncData.req.status) {
                        case 401: // AuthError
                            let rv = {};
                            rv.errorObj = eas.sync.finish("error", "401");
                            rv.errorType = "PasswordPrompt";
                            resolve(rv);
                            break;

                        case 200:
                                responseData["MS-ASProtocolVersions"] =  syncData.req.getResponseHeader("MS-ASProtocolVersions");
                                responseData["MS-ASProtocolCommands"] =  syncData.req.getResponseHeader("MS-ASProtocolCommands");                        

                                tbSync.dump("EAS OPTIONS with response (status: 200)", "\n" +
                                "responseText: " + syncData.req.responseText + "\n" +
                                "responseHeader(MS-ASProtocolVersions): " + responseData["MS-ASProtocolVersions"]+"\n" +
                                "responseHeader(MS-ASProtocolCommands): " + responseData["MS-ASProtocolCommands"]);

                                if (responseData && responseData["MS-ASProtocolCommands"] && responseData["MS-ASProtocolVersions"]) {
                                    syncData.accountData.setAccountProperty("allowedEasCommands", responseData["MS-ASProtocolCommands"]);
                                    syncData.accountData.setAccountProperty("allowedEasVersions", responseData["MS-ASProtocolVersions"]);
                                    syncData.accountData.setAccountProperty("lastEasOptionsUpdate", Date.now());
                                }
                                resolve();
                            break;

                        default:
                                resolve();
                            break;

                    }
                };
                
                syncData.setSyncState("send.request.options");
                syncData.req.send();
                
            });
            
            if (result && result.hasOwnProperty("errorType") && result.errorType == "PasswordPrompt") {
                if (allowedRetries > 0) {
                    let authData = eas.network.getAuthData(syncData.accountData);
                    let promptData = {
                        windowID: "auth:" + syncData.accountData.accountID,
                        accountname: syncData.accountData.getAccountProperty("accountname"),
                        usernameLocked: syncData.accountData.isConnected(),
                        username: authData.user
                    }
                    
                    let syncState = syncData.getSyncState(); 
                    syncData.setSyncState("passwordprompt");
                    let credentials = await tbSync.passwordManager.asyncPasswordPrompt(promptData, eas.openWindows);
                    if (credentials) {
                        // Update login data and try again.
                        authData.updateLoginData(credentials.username, credentials.password);
                        syncData.setSyncState(syncState);
                        retry = true;
                    }
                }

                if (!retry) {
                    throw result.errorObj;
                }
            }
            
            allowedRetries--;
        } while (retry);
    },










    // AUTODISCOVER
    
    updateServerConnectionViaAutodiscover: async function (syncData) {
        syncData.setSyncState("prepare.request.autodiscover");
        let authData = eas.network.getAuthData(syncData.accountData);

        syncData.setSyncState("send.request.autodiscover");
        let result = await eas.network.getServerConnectionViaAutodiscover(authData.user, authData.password, 30*1000);

        syncData.setSyncState("eval.response.autodiscover");
        if (result.errorcode == 200) {
            //update account
            syncData.accountData.setAccountProperty("host", eas.network.stripAutodiscoverUrl(result.server)); 
            syncData.accountData.setAccountProperty("user", result.user);
            syncData.accountData.setAccountProperty("https", (result.server.substring(0,5) == "https"));
        }

        return result.errorcode;
    },
    
    stripAutodiscoverUrl: function(url) {
        let u = url;
        while (u.endsWith("/")) { u = u.slice(0,-1); }
        if (u.endsWith("/Microsoft-Server-ActiveSync")) u=u.slice(0, -28);
        else tbSync.dump("Received non-standard EAS url via autodiscover:", url);

        return u.split("//")[1]; //cut off protocol
    },

    getServerConnectionViaAutodiscover : async function (user, password, maxtimeout) {
        let urls = [];
        let parts = user.split("@");
        
        urls.push({"url":"http://autodiscover."+parts[1]+"/autodiscover/autodiscover.xml", "user":user});
        urls.push({"url":"http://"+parts[1]+"/autodiscover/autodiscover.xml", "user":user});
        urls.push({"url":"http://autodiscover."+parts[1]+"/Autodiscover/Autodiscover.xml", "user":user});
        urls.push({"url":"http://"+parts[1]+"/Autodiscover/Autodiscover.xml", "user":user});

        urls.push({"url":"https://autodiscover."+parts[1]+"/autodiscover/autodiscover.xml", "user":user});
        urls.push({"url":"https://"+parts[1]+"/autodiscover/autodiscover.xml", "user":user});
        urls.push({"url":"https://autodiscover."+parts[1]+"/Autodiscover/Autodiscover.xml", "user":user});
        urls.push({"url":"https://"+parts[1]+"/Autodiscover/Autodiscover.xml", "user":user});
        
        let requests = [];
        let responses = []; //array of objects {url, error, server}
        
        for (let i=0; i< urls.length; i++) {
            await tbSync.tools.sleep(200);
            requests.push( eas.network.getServerConnectionViaAutodiscoverRedirectWrapper(urls[i].url, urls[i].user, password, maxtimeout) );
        }

        try {
            responses = await Promise.all(requests); 
        } catch (e) {
            responses.push(e.result); //this is actually a success, see return value of getServerConnectionViaAutodiscoverRedirectWrapper()
        }
        
        let result;
        let log = [];        
        for (let r=0; r < responses.length; r++) {
            log.push("*  "+responses[r].url+" @ " + responses[r].user +" : " + (responses[r].server ? responses[r].server : responses[r].error));

            if (responses[r].server) {
                result = {"server": responses[r].server, "user": responses[r].user, "error": "", "errorcode": 200};
                break;
            }
            
            if (responses[r].error == 403 || responses[r].error == 401) {
                //we could still find a valid server, so just store this state
                result = {"server": "", "user": responses[r].user, "errorcode": responses[r].error, "error": tbSync.getString("status." + responses[r].error, "eas")};
            }
        } 
        
        //this is only reached on fail, if no result defined yet, use general error
        if (!result) { 
            result = {"server": "", "user": user, "error": tbSync.getString("autodiscover.Failed","eas").replace("##user##", user), "errorcode": 503};
        }

        tbSync.eventlog.add("error", new tbSync.EventLogInfo("eas"), result.error, log.join("\n"));
        return result;        
    },
       
    getServerConnectionViaAutodiscoverRedirectWrapper : async function (url, user, password, maxtimeout) {        
        //using HEAD to find URL redirects until response URL no longer changes 
        // * XHR should follow redirects transparently, but that does not always work, POST data could get lost, so we
        // * need to find the actual POST candidates (example: outlook.de accounts)
        let result = {};
        let method = "HEAD";
        let connection = { url, user };
        
        do {            
            await tbSync.tools.sleep(200);
            result = await eas.network.getServerConnectionViaAutodiscoverRequest(method, connection, password, maxtimeout);
            method = "";
            
            if (result.error == "redirect found") {
                tbSync.dump("EAS autodiscover URL redirect",  "\n" + connection.url + " @ " + connection.user + " => \n" + result.url + " @ " + result.user);
                connection.url = result.url;
                connection.user = result.user;
                method = "HEAD";
            } else if (result.error == "POST candidate found") {
                method = "POST";
            }

        } while (method);
        
        //invert reject and resolve, so we exit the promise group on success right away
        if (result.server) {
            let e = new Error("Not an error (early exit from promise group)");
            e.result = result;
            throw e;
        } else {
            return result;
        }
    },    
    
    getServerConnectionViaAutodiscoverRequest: function (method, connection, password, maxtimeout) {
        tbSync.dump("Querry EAS autodiscover URL", connection.url + " @ " + connection.user);
        
        return new Promise(function(resolve,reject) {
            
            let xml = '<?xml version="1.0" encoding="utf-8"?>\r\n';
            xml += '<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/mobilesync/requestschema/2006">\r\n';
            xml += '<Request>\r\n';
            xml += '<EMailAddress>' + connection.user + '</EMailAddress>\r\n';
            xml += '<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/mobilesync/responseschema/2006</AcceptableResponseSchema>\r\n';
            xml += '</Request>\r\n';
            xml += '</Autodiscover>\r\n';
            
            let userAgent = eas.prefs.getCharPref("clientID.useragent"); //plus calendar.useragent.extra = Lightning/5.4.5.2

            // Create request handler - API changed with TB60 to new XMKHttpRequest()
            let req = new XMLHttpRequest();
            req.mozBackgroundRequest = true;
            req.open(method, connection.url, true);
            req.timeout = maxtimeout;
            req.setRequestHeader("User-Agent", userAgent);
            
            let secure = (connection.url.substring(0,8).toLowerCase() == "https://");
            
            if (method == "POST") {
                req.setRequestHeader("Content-Length", xml.length);
                req.setRequestHeader("Content-Type", "text/xml");
                if (secure) req.setRequestHeader("Authorization", "Basic " + tbSync.tools.b64encode(connection.user + ":" + password));                
            }

            req.ontimeout = function () {
                tbSync.dump("EAS autodiscover with timeout", "\n" + connection.url + " => \n" + req.responseURL);
                resolve({"url":req.responseURL, "error":"timeout", "server":"", "user":connection.user});
            };
           
            req.onerror = function () {
                let error = tbSync.network.createTCPErrorFromFailedXHR(req);
                if (!error) error = req.responseText;
                tbSync.dump("EAS autodiscover with error ("+error+")",  "\n" + connection.url + " => \n" + req.responseURL);
                resolve({"url":req.responseURL, "error":error, "server":"", "user":connection.user});
            };

            req.onload = function() { 
                //initiate rerun on redirects
                if (req.responseURL != connection.url) {
                    resolve({"url":req.responseURL, "error":"redirect found", "server":"", "user":connection.user});
                    return;
                }

                //initiate rerun on HEAD request without redirect (rerun and do a POST on this)
                if (method == "HEAD") {
                    resolve({"url":req.responseURL, "error":"POST candidate found", "server":"", "user":connection.user});
                    return;
                }

                //ignore POST without autherization (we just do them to get redirect information)
                if (!secure) {
                    resolve({"url":req.responseURL, "error":"unsecure POST", "server":"", "user":connection.user});
                    return;
                }
                
                //evaluate secure POST requests which have not been redirected
                tbSync.dump("EAS autodiscover POST with status (" + req.status + ")",   "\n" + connection.url + " => \n" + req.responseURL  + "\n[" + req.responseText + "]");
                
                if (req.status === 200) {
                    let data = eas.xmltools.getDataFromXMLString(req.responseText);
            
                    if (!(data === null) && data.Autodiscover && data.Autodiscover.Response && data.Autodiscover.Response.Action) {
                        // "Redirect" or "Settings" are possible
                        if (data.Autodiscover.Response.Action.Redirect) {
                            // redirect, start again with new user
                            let newuser = action.Redirect;
                            resolve({"url":req.responseURL, "error":"redirect found", "server":"", "user":newuser});

                        } else if (data.Autodiscover.Response.Action.Settings) {
                            // get server settings
                            let server = eas.xmltools.nodeAsArray(data.Autodiscover.Response.Action.Settings.Server);

                            for (let count = 0; count < server.length; count++) {
                                if (server[count].Type == "MobileSync" && server[count].Url) {
                                    resolve({"url":req.responseURL, "error":"", "server":server[count].Url, "user":connection.user});
                                    return;
                                }
                            }
                        }
                    } else {
                        resolve({"url":req.responseURL, "error":"invalid", "server":"", "user":connection.user});
                    }
                } else {
                    resolve({"url":req.responseURL, "error":req.status, "server":"", "user":connection.user});                     
                }
            };
            
            if (method == "HEAD") req.send();
            else  req.send(xml);
            
        });
    }
}
