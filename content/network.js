/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

var network = {
    
    // AUTODISCOVER        
    updateServerConnectionViaAutodiscover: async function (syncdata) {
        tbSync.setSyncState("prepare.request.autodiscover", syncdata.account);
        let user = tbSync.db.getAccountSetting(syncdata.account, "user");
        let password = eas.auth.getPassword(tbSync.db.getAccount(syncdata.account));

        tbSync.setSyncState("send.request.autodiscover", syncdata.account);
        let result = await eas.network.getServerConnectionViaAutodiscover(user, password, 30*1000);

        tbSync.setSyncState("eval.response.autodiscover", syncdata.account);
        if (result.errorcode == 200) {
            //update account
            tbSync.db.setAccountSetting(syncdata.account, "host", eas.stripAutodiscoverUrl(result.server)); 
            tbSync.db.setAccountSetting(syncdata.account, "user", result.user);
            tbSync.db.setAccountSetting(syncdata.account, "https", (result.server.substring(0,5) == "https") ? "1" : "0");
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
        for (let i=0; i< urls.length; i++) {
            await tbSync.tools.sleep(200, false);
            requests.push( eas.network.getServerConnectionViaAutodiscoverRedirectWrapper(urls[i].url, urls[i].user, password, maxtimeout) );
        }
 
        let responses = []; //array of objects {url, error, server}
        try {
            responses = await Promise.all(requests); 
        } catch (e) {
            responses.push(e); //this is actually a success, see return value of getServerConnectionViaAutodiscoverRedirectWrapper()
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

        tbSync.errorlog.add("error", new tbSync.ErrorInfo("eas"), result.error, log.join("\n"));
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
            await tbSync.tools.sleep(200, false);
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
        if (result.server) throw result;
        else return result;
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
    },    

}
