/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

var xmltools = {

    isString : function (obj) {
        return (Object.prototype.toString.call(obj) === '[object String]');
    },
        
    checkString : function(d, fallback = "") {
        if (this.isString(d)) return d;
        else return fallback;
    },	    
        
    nodeAsArray : function (node) {
        let a = [];
        if (node) {
            //return, if already an array
            if (node instanceof Array) return node;

            //else push node into an array
            a.push(node);
        }
        return a;
    },

    hasWbxmlDataField: function(wbxmlData, path) {
        if (wbxmlData) {		
            let pathElements = path.split(".");
            let data = wbxmlData;
            for (let x = 0; x < pathElements.length; x++) {
                if (data[pathElements[x]]) data = data[pathElements[x]];
                else return false;
            }
            return true;
        }
        return false;
    },

    getWbxmlDataField: function(wbxmlData,path) {
        if (wbxmlData) {		
            let pathElements = path.split(".");
            let data = wbxmlData;
            let valid = true;
            for (let x = 0; valid && x < pathElements.length; x++) {
                if (data[pathElements[x]]) data = data[pathElements[x]];
                else valid = false;
            }
            if (valid) return data;
        }
        return false
    },

    //print content of xml data object (if debug output enabled)
    printXmlData : function (data, printApplicationData) {
        if ((tbSync.prefSettings.getBoolPref("log.toconsole") || tbSync.prefSettings.getBoolPref("log.tofile")) && printApplicationData) {
            let dump = JSON.stringify(data);
            tbSync.dump("Extracted XML data", "\n" + dump);
        }
    },

    getDataFromXMLString: function (str) {
        let data = null;
        let xml = "";        
        if (str == "") return data;
        
        let oParser = (Services.vc.compare(Services.appinfo.platformVersion, "61.*") >= 0) ? new DOMParser() : Components.classes["@mozilla.org/xmlextras/domparser;1"].createInstance(Components.interfaces.nsIDOMParser);
        try {
            xml = oParser.parseFromString(str, "application/xml");
        } catch (e) {
            //however, domparser does not throw an error, it returns an error document
            //https://developer.mozilla.org/de/docs/Web/API/DOMParser
            //just in case
            throw eas.finishSync("mailformed-xml", eas.flags.abortWithError);
        }

        //check if xml is error document
        if (xml.documentElement.nodeName == "parsererror") {
            tbSync.dump("BAD XML", "The above XML and WBXML could not be parsed correctly, something is wrong.");
            throw eas.finishSync("mailformed-xml", eas.flags.abortWithError);
        }

        try {
            data = this.getDataFromXML(xml);
        } catch (e) {
            throw eas.finishSync("mailformed-data", eas.flags.abortWithError);
        }
        
        return data;
    },
    
    //create data object from XML node
    getDataFromXML : function (nodes) {
        
        /*
         * The passed nodes value could be an entire document in a single node (type 9) or a 
         * single element node (type 1) as returned by getElementById. It could however also 
         * be an array of nodes as returned by getElementsByTagName or a nodeList as returned
         * by childNodes. In that case node.length is defined.
         */        
        
        // create the return object
        let obj = {};
        let nodeList = [];
        let multiplicity = {};
        
        if (nodes.length === undefined) nodeList.push(nodes);
        else nodeList = nodes;
        
        // nodelist contains all childs, if two childs have the same name, we cannot add the chils as an object, but as an array of objects
        for (let node of nodeList) { 
            if (node.nodeType == 1 || node.nodeType == 3) {
                if (!multiplicity.hasOwnProperty(node.nodeName)) multiplicity[node.nodeName] = 0;
                multiplicity[node.nodeName]++;
                //if this nodeName has multiplicity > 1, prepare obj  (but only once)
                if (multiplicity[node.nodeName]==2) obj[node.nodeName] = [];
            }
        }

        // process nodes
        for (let node of nodeList) { 
            switch (node.nodeType) {
                case 9: 
                    //document node, dive directly and process all children
                    if (node.hasChildNodes) obj = this.getDataFromXML(node.childNodes);
                    break;
                case 1: 
                    //element node
                    if (node.hasChildNodes) {
                        //if this is an element with only one text child, do not dive, but get text childs value
                        let o;
                        if (node.childNodes.length == 1 && node.childNodes.item(0).nodeType==3) {
                            //the passed xml is a save xml with all special chars in the user data encoded by encodeURIComponent
                            o = decodeURIComponent(node.childNodes.item(0).nodeValue);
                        } else {
                            o = this.getDataFromXML(node.childNodes);
                        }
                        //check, if we can add the object directly, or if we have to push it into an array
                        if (multiplicity[node.nodeName]>1) obj[node.nodeName].push(o)
                        else obj[node.nodeName] = o; 
                    }
                    break;
            }
        }
        return obj;
    }
    
};
