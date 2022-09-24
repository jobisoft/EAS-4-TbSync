/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

 var { XPCOMUtils } = ChromeUtils.import(
    "resource://gre/modules/XPCOMUtils.jsm"
);

XPCOMUtils.defineLazyModuleGetters(this, {
    newUID: "resource:///modules/AddrBookUtils.jsm",
    AddrBookCard: "resource:///modules/AddrBookCard.jsm",
    BANISHED_PROPERTIES: "resource:///modules/VCardUtils.jsm",
    VCardProperties: "resource:///modules/VCardUtils.jsm",
    VCardPropertyEntry: "resource:///modules/VCardUtils.jsm",
    VCardUtils: "resource:///modules/VCardUtils.jsm",
});

const eas = TbSync.providers.eas;

var Contacts = {

    // These functions handle categories compatible to the Category Manager add-on, which is compatible
    // to lots of other sync tools (sogo, carddav-sync, roundcube)
    categoriesFromString: function (catString) {
        let catsArray = [];
        if (catString.trim().length>0) catsArray = catString.trim().split("\u001A").filter(String);
        return catsArray;
    },

    categoriesToString: function (catsArray) {
        return catsArray.join("\u001A");
    },

    /* The following TB properties are not synced to the server:
       - only one WebPage
       - more than 3 emails
       - more than one fax, pager, mobile, work, home
       - position (in org)
    */

    vcard_array_fields : {
        n : 5,
        adr : 7,
        org : 2 
    },

    map_EAS_properties_to_vCard : {
        FileAs: {item: "fn", type: "text", params: {}}, /* DisplayName */ 

        Birthday: {item: "bday", type: "date", params: {}},
        Anniversary: {item: "anniversary", type: "date", params: {}},
        
        LastName: {item: "n", type: "text", params: {}, index: 0},
        FirstName: {item: "n", type: "text", params: {}, index: 1},
        MiddleName: {item: "n", type: "text", params: {}, index: 2},
        Title: {item: "n", type: "text", params: {}, index: 3},
        Suffix: {item: "n", type: "text", params: {}, index: 4},

        Notes: {item: "note", type: "text", params: {}},

        // What should we do with Email 4+ ?
        Email1Address: {item: "email", type: "text", params: {}},
        Email2Address: {item: "email", type: "text", params: {}, entry: 1},
        Email3Address: {item: "email", type: "text", params: {}, entry: 2},

        // WebPage has fallbackParams defined, to pick any url, if the specified
        // one is not found, and the user has created a home or work url. It will
        // come back as "Other".
        WebPage: {item: "url", type: "text", params: {}, fallbackParams: [{type: "home"}, {type: "work"}]},
        
        CompanyName: {item: "org", type: "text", params: {}, index: 0}, /* Company */
        Department: {item: "org", type: "text", params: {}, index: 1}, /* Department */
        JobTitle: { item: "title", type: "text", params: {} }, /* JobTitle */ 

        MobilePhoneNumber: { item: "tel", type: "text", params: {type: "cell" }},
        PagerNumber: { item: "tel", type: "text", params: {type: "pager" }},
        HomeFaxNumber: { item: "tel", type: "text", params: {type: "fax" }},
        HomePhoneNumber: { item: "tel", type: "text", params: {type: "home"}, fallbackParams: [{}]},
        BusinessPhoneNumber: { item: "tel", type: "text", params: {type: "work"}},
        Home2PhoneNumber: { item: "tel", type: "text", params: {type: "home"}, entry: 1 },
        Business2PhoneNumber: { item: "tel", type: "text", params: {type: "work"}, entry: 1 },

        HomeAddressStreet: {item: "adr", type: "text", params: {type: "home"}, index: 2},  // needs special handling
        HomeAddressCity: {item: "adr", type: "text", params: {type: "home"}, index: 3},
        HomeAddressState: {item: "adr", type: "text", params: {type: "home"}, index: 4},
        HomeAddressPostalCode: {item: "adr", type: "text", params: {type: "home"}, index: 5},
        HomeAddressCountry: {item: "adr", type: "text", params: {type: "home"}, index: 6},

        BusinessAddressStreet: {item: "adr", type: "text", params: {type: "work"}, index: 2},  // needs special handling
        BusinessAddressCity: {item: "adr", type: "text", params: {type: "work"}, index: 3},
        BusinessAddressState: {item: "adr", type: "text", params: {type: "work"}, index: 4},
        BusinessAddressPostalCode: {item: "adr", type: "text", params: {type: "work"}, index: 5},
        BusinessAddressCountry: {item: "adr", type: "text", params: {type: "work"}, index: 6},

        OtherAddressStreet: {item: "adr", type: "text", params: {}, index: 2},  // needs special handling
        OtherAddressCity: {item: "adr", type: "text", params: {}, index: 3},
        OtherAddressState: {item: "adr", type: "text", params: {}, index: 4},
        OtherAddressPostalCode: {item: "adr", type: "text", params: {}, index: 5},
        OtherAddressCountry: {item: "adr", type: "text", params: {}, index: 6},

        // Misusing this EAS field, so that "Custom1" is saved to the server.
        OfficeLocation: {item: "x-custom1", type: "text", params: {}},

        Picture: {item: "photo", params: {}, type: "uri"},

        // TB shows them as undefined, but showing them might be better, than not. Use a prefix.
        AssistantPhoneNumber: { item: "tel", type: "text", params: {type: "Assistant"}, prefix: true},
        CarPhoneNumber: { item: "tel", type: "text", params: {type: "Car"}, prefix: true},
        RadioPhoneNumber: { item: "tel", type: "text", params: {type: "Radio"}, prefix: true},
        BusinessFaxNumber: { item: "tel", type: "text", params: {type: "WorkFax"}, prefix: true},
    },
   
    map_EAS_properties_to_vCard_set2 : {
        NickName: {item: "nickname", type: "text", params: {} },
        // Misusing these EAS fields, so that "Custom2,3,4" is saved to the server.
        CustomerId: {item: "x-custom2", type: "text", params: {}},
        GovernmentId: {item: "x-custom3", type: "text", params: {}},
        AccountName:  {item: "x-custom4", type: "text", params: {}},

        IMAddress: {item: "impp", type: "text", params: {} },
        IMAddress2: {item: "impp", type: "text", params: {}, entry: 1 },
        IMAddress3: {item: "impp", type: "text", params: {}, entry: 2 },

        CompanyMainPhone: { item: "tel", type: "text", params: {type: "Company"}, prefix: true},
    },

    // There are currently no TB fields for these values, TbSync will store (and
    // resend) them, but will not allow to view/edit.
    unused_EAS_properties: [
        "Alias", //pseudo field
        "WeightedRank", //pseudo field
        "YomiCompanyName", //japanese phonetic equivalent
        "YomiFirstName", //japanese phonetic equivalent
        "YomiLastName", //japanese phonetic equivalent
        "CompressedRTF",
        "MMS",
        // Former custom EAS fields, no longer added to UI after 102.
        "ManagerName",
        "AssistantName",
        "Spouse",
    ],

    // Normalize a parameters entry, to be able to find matching existing
    // entries. If we want to be less restrictive, we need to check if all
    // the requested values exist. But we should be the only one who sets
    // the vCard props, so it should be safe. Except someone moves a contact.
    // Should we prevent that via a vendor id in the vcard?
    normalizeParameters: function (unordered) {
        return JSON.stringify(
            Object.keys(unordered).map(e => `${e}`.toLowerCase()).sort().reduce(
                (obj, key) => { 
                    obj[key] = `${unordered[key]}`.toLowerCase(); 
                return obj;
                }, 
                {}
            )
        );
    },

    getValue: function (vCardProperties, vCard_property) {
        let parameters = [vCard_property.params];
        if (vCard_property.fallbackParams) {
            parameters.push(...vCard_property.fallbackParams);
        }
        let entries;
        for (let normalizedParams of parameters.map(this.normalizeParameters)) {
            entries = vCardProperties.getAllEntries(vCard_property.item)
                .filter(e => normalizedParams == this.normalizeParameters(e.params));
            if (entries.length > 0) {
                break;
            }
        }

        // Which entry should we take?
        let entryNr = vCard_property.entry || 0;
        if (entries[entryNr]) {
            let value;
            // The org field sometimes comes back as a string, even though it
            // should be an array.
            if (this.vcard_array_fields[vCard_property.item]) {
                if (!Array.isArray(entries[entryNr].value)) {
                    // If the returned value is a single string, return it only
                    // when index 0 is requested, otherwise return nothing.
                    value =  vCard_property.index == 0 ? entries[entryNr].value : "";
                } else {
                    value = entries[entryNr].value[vCard_property.index];
                }
            } else {
                value = entries[entryNr].value;
            }

            if (value) {
                if (vCard_property.prefix && value.startsWith(`${vCard_property.params.type}: `)) {
                    return value.substring(`${vCard_property.params.type}: `.length);
                }
                return value;
            }
        }
        return "";
    },

    /**
     * Reads a DOM File and returns a Promise for its dataUrl.
     *
     * @param {File} file
     * @returns {string}
     */
    getDataUrl(file) {
        return new Promise((resolve, reject) => {
            var reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = function() {
                resolve(reader.result);
            };
            reader.onerror = function(error) {
                resolve("");
            };
        });
    },



    // --------------------------------------------------------------------------- //
    // Read WBXML and set Thunderbird item
    // --------------------------------------------------------------------------- //
    setThunderbirdItemFromWbxml: function (abItem, data, id, syncdata, mode = "standard") {
        let asversion = syncdata.accountData.getAccountProperty("asversion");
        if (TbSync.prefs.getIntPref("log.userdatalevel") > 2) TbSync.dump("Processing " + mode + " contact item", id);

        // Make sure we are dealing with a vCard, so we can update the card just
        // by updating its vCardProperties.
        if (!abItem._card.supportsVCard) {
            // This is an older card??
            throw new Error("It looks like you are trying to sync a TB91 sync state. Does not work.");
        }
        let vCardProperties = abItem._card.vCardProperties
        abItem.primaryKey = id;

        // Loop over all known EAS properties (two EAS sets Contacts and Contacts2).
        for (let set=0; set < 2; set++) {
            let properties = (set == 0) ? this.EAS_properties : this.EAS_properties2;

            for (let EAS_property of properties) {
                let vCard_property = (set == 0) ? this.map_EAS_properties_to_vCard[EAS_property] : this.map_EAS_properties_to_vCard_set2[EAS_property];
                let value;
                switch (EAS_property) {
                    case "Notes":
                        if (asversion == "2.5") {
                            value = eas.xmltools.checkString(data.Body);
                        } else if (data.Body && data.Body.Data) {
                            value = eas.xmltools.checkString(data.Body.Data);
                        }
                    break;

                    default:
                        value = eas.xmltools.checkString(data[EAS_property]);
                }
                
                let normalizedParams = this.normalizeParameters(vCard_property.params)
                let entries = vCardProperties.getAllEntries(vCard_property.item)
                    .filter(e => normalizedParams == this.normalizeParameters(e.params));
                // Which entry should we update? Add empty entries, if the requested entry number
                // does not yet exist.
                let entryNr = vCard_property.entry || 0;
                while (entries.length <= entryNr) {
                    let newEntry = new VCardPropertyEntry(
                        vCard_property.item, 
                        vCard_property.params, 
                        vCard_property.type, 
                        this.vcard_array_fields[vCard_property.item] 
                            ? new Array(this.vcard_array_fields[vCard_property.item]).fill("") 
                            : ""
                    );
                    vCardProperties.addEntry(newEntry);
                    entries = vCardProperties.getAllEntries(vCard_property.item);
                    entryNr = entries.length - 1;
                }

                // Is this property part of the send data?
                if (value) {
                    // Do we need to manipulate the value?
                    switch (EAS_property) {
                        case "Picture":
                            value = `data:image/jpeg;base64,${eas.xmltools.nodeAsArray(data.Picture)[0]}`; //Kerio sends Picture as container
                            break;
                        
                        case "Birthday":
                        case "Anniversary":
                            let dateObj = new Date(value);
                            value = dateObj.toISOString().substr(0, 10);
                            break;

                        case "Email1Address":
                        case "Email2Address":
                        case "Email3Address":
                            let parsedInput = MailServices.headerParser.makeFromDisplayAddress(value);
                            let fixedValue =  (parsedInput && parsedInput[0] && parsedInput[0].email) ? parsedInput[0].email : value;
                            if (fixedValue != value) {
                                if (TbSync.prefs.getIntPref("log.userdatalevel") > 2) TbSync.dump("Parsing email display string via RFC 2231 and RFC 2047 ("+EAS_property+")", value + " -> " + fixedValue);
                                value = fixedValue;
                            }
                            break;
                        
                        case "HomeAddressStreet":
                        case "BusinessAddressStreet":
                        case "OtherAddressStreet":
                            // Thunderbird accepts an array in the vCardProperty of the 2nd index of the adr field.
                            let seperator = String.fromCharCode(syncdata.accountData.getAccountProperty("seperator")); // options are 44 (,) or 10 (\n)
                            value = value.split(seperator);
                        break;
                    }

                    // Add a typePrefix for fields unknown to TB (better: TB should use the type itself).
                    if (vCard_property.prefix && !value.startsWith(`${vCard_property.params.type}: `)) {
                        value = `${vCard_property.params.type}: ${value}`;
                    }

                    // Is this an array value?
                    if (this.vcard_array_fields[vCard_property.item]) {
                        // Make sure this is an array.
                        if (!Array.isArray(entries[entryNr].value)) {
                            let arr = new Array(this.vcard_array_fields[vCard_property.item]).fill("");
                            arr[0] = entries[entryNr].value;
                            entries[entryNr].value = arr;
                        }
                        entries[entryNr].value[vCard_property.index] = value;
                    } else {
                        entries[entryNr].value = value;
                    }
                } else {
                    if (this.vcard_array_fields[vCard_property.item]) {
                        // Make sure this is an array.
                        if (!Array.isArray(entries[entryNr].value)) {
                            let arr = new Array(this.vcard_array_fields[vCard_property.item]).fill("");
                            arr[0] = entries[entryNr].value;
                            entries[entryNr].value = arr;
                        }
                        entries[entryNr].value[vCard_property.index] = "";
                    } else {
                        entries[entryNr].value = "";
                    }
                }
            }
        }

        // Take care of categories and children, stored in contacts property bag.
        let containers = [];
        containers.push(["Categories", "Category"]);
        containers.push(["Children", "Child"]);
        for (let c=0; c < containers.length; c++) {
            if (data[containers[c][0]] && data[containers[c][0]][containers[c][1]]) {
                let cats = [];
                if (Array.isArray(data[containers[c][0]][containers[c][1]])) cats = data[containers[c][0]][containers[c][1]];
                else cats.push(data[containers[c][0]][containers[c][1]]);
                abItem.setProperty(containers[c][0], this.categoriesToString(cats));
            }
        }

        // Take care of un-mappable EAS options, which are stored in the contacts
        // property bag.
        for (let i=0; i < this.unused_EAS_properties.length; i++) {
            if (data[this.unused_EAS_properties[i]]) abItem.setProperty("EAS-" + this.unused_EAS_properties[i], data[this.unused_EAS_properties[i]]);
        }

        // Remove all entries, which are marked for deletion.
        vCardProperties.entries = vCardProperties.entries.filter(e => Array.isArray(e.value) ? e.value.some(a => a != "") : e.value != "");

        // Further manipulations (a few getters are still usable \o/).
        if (syncdata.accountData.getAccountProperty("displayoverride")) {
            abItem._card.displayName = abItem._card.firstName + " " + abItem._card.lastName;
            if (abItem._card.displayName == " " ) {
                let company = (vCardProperties.getFirstValue("org") || [""])[0];
                abItem._card.displayName = company || abItem._card.primaryEmail
            }
        }
    },




    // --------------------------------------------------------------------------- //
    //read TB event and return its data as WBXML
    // --------------------------------------------------------------------------- //
    getWbxmlFromThunderbirdItem: async function (abItem, syncdata, isException = false) {
        let asversion = syncdata.accountData.getAccountProperty("asversion");
        let wbxml = eas.wbxmltools.createWBXML("", syncdata.type); //init wbxml with "" and not with precodes, and set initial codepage
        let nowDate = new Date();

        // Make sure we are dealing with a vCard, so we can access its vCardProperties.
        if (!abItem._card.supportsVCard) {
            throw new Error("It looks like you are trying to sync a TB91 sync state. Does not work.");
        }
        let vCardProperties = abItem._card.vCardProperties

        // Loop over all known EAS properties (send empty value if not set).
        for (let EAS_property of this.EAS_properties) {
            // Some props need special handling.
            let vCard_property = this.map_EAS_properties_to_vCard[EAS_property];
            let value;
            switch (EAS_property) {
                case "Notes":
                    // Needs to be done later, because we have to switch the code page.
                    continue;

                case "Picture": {
                    let photoUrl = abItem._card.photoURL;
                    if (!photoUrl) {
                        continue;
                    }
                    if (photoUrl.startsWith("file://")) {
                        let realPhotoFile = Services.io.newURI(photoUrl).QueryInterface(Ci.nsIFileURL).file;
                        let photoFile = await File.createFromNsIFile(realPhotoFile);
                        photoUrl = await this.getDataUrl(photoFile);
                    }
                    if (photoUrl.startsWith("data:image/")) {
                        let parts = photoUrl.split(",");
                        parts.shift();
                        value = parts.join(",");
                    }
                }
                break;

                case "Birthday":
                case "Anniversary": {
                    let raw = this.getValue(vCardProperties, vCard_property);
                    if (raw) {
                        let dateObj = new Date(raw);
                        value = dateObj.toISOString();
                        console.log(value);
                    }
                }
                break;
                    
                case "HomeAddressStreet":
                case "BusinessAddressStreet":
                case "OtherAddressStreet": {
                    let raw = this.getValue(vCardProperties, vCard_property);
                    try {
                        if (raw) {
                            // We either get a single string or an array for the
                            // street adr field from Thunderbird.
                            if (!Array.isArray(raw)) {
                                raw = [raw];
                            }
                            let seperator = String.fromCharCode(syncdata.accountData.getAccountProperty("seperator")); // options are 44 (,) or 10 (\n)
                            value = raw.join(seperator);
                        }
                    } catch (ex) {
                        throw new Error(`Failed to eval value: <${JSON.stringify(raw)}> @ ${JSON.stringify(vCard_property)}`);
                    }
                }
                break;

                default: {
                    value = this.getValue(vCardProperties, vCard_property);
                }
            }
            
            if (value) {
                wbxml.atag(EAS_property, value);
            }
        }

        // Take care of un-mappable EAS option.
        for (let i=0; i < this.unused_EAS_properties.length; i++) {
            let value = abItem.getProperty("EAS-" + this.unused_EAS_properties[i], "");
            if (value) wbxml.atag(this.unused_EAS_properties[i], value);
        }

        // Take care of categories and children
        let containers = [];
        containers.push(["Categories", "Category"]);
        containers.push(["Children", "Child"]);
        for (let c=0; c < containers.length; c++) {
            let cats = abItem.getProperty(containers[c][0], "");
            if (cats) {
                let catsArray = this.categoriesFromString(cats);
                wbxml.otag(containers[c][0]);
                for (let ca=0; ca < catsArray.length; ca++) wbxml.atag(containers[c][1], catsArray[ca]);
                wbxml.ctag();
            }
        }

        // Take care of notes - SWITCHING TO AirSyncBase (if 2.5, we still need Contact group here!)
        let description = this.getValue(vCardProperties, this.map_EAS_properties_to_vCard["Notes"]);
        if (asversion == "2.5") {
            wbxml.atag("Body", description);
        } else {
            wbxml.switchpage("AirSyncBase");
            wbxml.otag("Body");
                wbxml.atag("Type", "1");
                wbxml.atag("EstimatedDataSize", "" + description.length);
                wbxml.atag("Data", description);
            wbxml.ctag();
        }

        // Take care of Contacts2 group - SWITCHING TO CONTACTS2
        wbxml.switchpage("Contacts2");

        // Loop over all known TB properties of EAS group Contacts2 (send empty value if not set).
        for (let EAS_property of this.EAS_properties2) {
            let vCard_property = this.map_EAS_properties_to_vCard_set2[EAS_property];
            let value = this.getValue(vCardProperties, vCard_property);
            if (value) wbxml.atag(EAS_property, value);
        }

        return wbxml.getBytes();
    }
}

Contacts.EAS_properties = Object.keys(Contacts.map_EAS_properties_to_vCard);
Contacts.EAS_properties2 = Object.keys(Contacts.map_EAS_properties_to_vCard_set2);
