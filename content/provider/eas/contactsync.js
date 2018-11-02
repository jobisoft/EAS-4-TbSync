/*
 * This file is part of EAS-4-TbSync.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/. 
 */
 
 "use strict";

eas.sync.Contacts = {

    createItem : function (card = null) {
        let item = {
            get id() {return this.card.getProperty("TBSYNCID", "")},
            set id(newId) {this.card.setProperty("TBSYNCID", newId)},
            get icalString() {return this.card.displayName + " (" + this.card.firstName + ", " + this.card.lastName + ") <"+this.card.primaryEmail+">" },
            clone: function () { return this; } //no real clone
        };
        
        //actually add the card
        item.card = card ? card : Components.classes["@mozilla.org/addressbook/cardproperty;1"].createInstance(Components.interfaces.nsIAbCard);
                    
        return item;
    },
    
    //these functions handle categories compatible to the Category Manager Add-On, which is compatible to lots of other sync tools (sogo, carddav-sync, roundcube)
    categoriesFromString: function (catString) {
        let catsArray = [];
        if (catString.trim().length>0) catsArray = catString.trim().split("\u001A").filter(String);
        return catsArray;
    },

    categoriesToString: function (catsArray) {
        return catsArray.join("\u001A");
    },


    /* The following TB properties are not yet synced anywhere:
       - , FamilyName 
       - _AimScreenName
       - WebPage2 (home)
*/    

    //includes all properties, which can be mapped 1-to-1
    map_TB_properties_to_EAS_properties : {
        DisplayName: 'FileAs',
        FirstName: 'FirstName',
        LastName: 'LastName',
        PrimaryEmail: 'Email1Address',
        SecondEmail: 'Email2Address',
        Email3Address: 'Email3Address',
        WebPage1: 'WebPage',
        SpouseName: 'Spouse',
        CellularNumber: 'MobilePhoneNumber',
        PagerNumber: 'PagerNumber',

        HomeCity: 'HomeAddressCity',
        HomeCountry: 'HomeAddressCountry',
        HomeZipCode: 'HomeAddressPostalCode',
        HomeState: 'HomeAddressState',
        HomePhone: 'HomePhoneNumber',
        
        Company: 'CompanyName',
        Department: 'Department',
        JobTitle: 'JobTitle',
        
        WorkCity: 'BusinessAddressCity',
        WorkCountry: 'BusinessAddressCountry',
        WorkZipCode: 'BusinessAddressPostalCode',
        WorkState: 'BusinessAddressState',
        WorkPhone: 'BusinessPhoneNumber',
        
        //Missusing so that "Custom1" is saved to the server
        Custom1: 'OfficeLocation',

        //As in TZPUSH
        FaxNumber: 'HomeFaxNumber',
    
        //Custom fields added to UI
        AssistantName: 'AssistantName',
        AssistantPhoneNumber: 'AssistantPhoneNumber',
        BusinessFaxNumber: 'BusinessFaxNumber',
        Business2PhoneNumber: 'Business2PhoneNumber',
        Home2PhoneNumber: 'Home2PhoneNumber',
        CarPhoneNumber: 'CarPhoneNumber',
        MiddleName: 'MiddleName',
        RadioPhoneNumber: 'RadioPhoneNumber',
        OtherAddressCity: 'OtherAddressCity',
        OtherAddressCountry: 'OtherAddressCountry',
        OtherAddressPostalCode: 'OtherAddressPostalCode',
        OtherAddressState: 'OtherAddressState'
    },

    //there are currently no TB fields for these values, TbSync will store (and resend) them, but will not allow to view/edit
    unused_EAS_properties: [
        'Suffix',
        'Title',
        'Alias', //pseudo field
        'WeightedRank', //pseudo field
        'YomiCompanyName', //japanese phonetic equivalent
        'YomiFirstName', //japanese phonetic equivalent
        'YomiLastName', //japanese phonetic equivalent
        'CompressedRTF' 
    ],
    
    map_TB_properties_to_EAS_properties2 : {
        NickName: 'NickName',
        //Missusing so that "Custom2,3,4" is saved to the server
        Custom2: 'CustomerId',
        Custom3: 'GovernmentId',
        Custom4: 'AccountName',
        //custom fields added to UI
        IMAddress: 'IMAddress',
        IMAddress2: 'IMAddress2',
        IMAddress3: 'IMAddress3',
        ManagerName: 'ManagerName',
        CompanyMainPhone: 'CompanyMainPhone',
        MMS: 'MMS'
    },


    // --------------------------------------------------------------------------- //
    // Read WBXML and set Thunderbird item
    // --------------------------------------------------------------------------- //
    setThunderbirdItemFromWbxml: function (item, data, id, syncdata) {
        let asversion = tbSync.db.getAccountSetting(syncdata.account, "asversion");

        item.card.setProperty("TBSYNCID", id);

        //loop over all known TB properties which map 1-to-1 (two EAS sets Contacts and Contacts2)
        for (let set=0; set < 2; set++) {
            let properties = (set == 0) ? this.TB_properties : this.TB_properties2;

            for (let p=0; p < properties.length; p++) {            
                let TB_property = properties[p];
                let EAS_property = (set == 0) ? this.map_TB_properties_to_EAS_properties[TB_property] : this.map_TB_properties_to_EAS_properties2[TB_property];            
                let value = xmltools.checkString(data[EAS_property]);
                
                //is this property part of the send data?
                if (value) {
                    //do we need to manipulate the value?
                    switch (EAS_property) {
                        case "Email1Address":
                        case "Email2Address":
                        case "Email3Address":
                            let parsedInput = MailServices.headerParser.makeFromDisplayAddress(value);
                            let fixedValue =  (parsedInput && parsedInput[0] && parsedInput[0].email) ? parsedInput[0].email : value;
                            if (fixedValue != value) {
                                if (tbSync.prefSettings.getIntPref("log.userdatalevel")>2) tbSync.dump("Parsing email display string via RFC 2231 and RFC 2047 ("+EAS_property+")", value + " -> " + fixedValue);
                                value = fixedValue;
                            }
                            break;
                    }
                    
                    item.card.setProperty(TB_property, value);
                } else {
                    //clear
                    item.card.deleteProperty(TB_property);
                }
            }
        }

        //take care of birthday and anniversary
        let dates = [];
        dates.push(["Birthday", "BirthDay", "BirthMonth", "BirthYear"]); //EAS, TB1, TB2, TB3
        dates.push(["Anniversary", "AnniversaryDay", "AnniversaryMonth", "AnniversaryYear"]);        
        for (let p=0; p < dates.length; p++) {
            let value = xmltools.checkString(data[dates[p][0]]);
            if (value == "") {
                //clear
                item.card.deleteProperty(dates[p][1]);
                item.card.deleteProperty(dates[p][2]);
                item.card.deleteProperty(dates[p][3]);
            } else {
                //set
                let dateObj = new Date(value);
                item.card.setProperty(dates[p][3], dateObj.getFullYear().toString());
                item.card.setProperty(dates[p][2], (dateObj.getMonth()+1).toString());
                item.card.setProperty(dates[p][1], dateObj.getDate().toString());
            }
        }


        //take care of multiline address fields
        let streets = [];
        let seperator = String.fromCharCode(tbSync.db.getAccountSetting(syncdata.account,"seperator")); // options are 44 (,) or 10 (\n)
        streets.push(["HomeAddressStreet", "HomeAddress", "HomeAddress2"]); //EAS, TB1, TB2
        streets.push(["BusinessAddressStreet", "WorkAddress", "WorkAddress2"]);
        streets.push(["OtherAddressStreet", "OtherAddress", "OtherAddress2"]);
        for (let p=0; p < streets.length; p++) {
            let value = xmltools.checkString(data[streets[p][0]]);
            if (value == "") {
                //clear
                item.card.deleteProperty(streets[p][1]);
                item.card.deleteProperty(streets[p][2]);
            } else {
                //set
                let lines = value.split(seperator);
                item.card.setProperty(streets[p][1], lines.shift());
                item.card.setProperty(streets[p][2], lines.join(seperator));
            }
        }


        //take care of photo
        if (data.Picture) {
            tbSync.addphoto(id + '.jpg', syncdata.targetId, item.card, xmltools.nodeAsArray(data.Picture)[0]); //Kerio sends Picture as container
        }
        

        //take care of notes
        if (asversion == "2.5") {
            item.card.setProperty("Notes", xmltools.checkString(data.Body));
        } else {
            if (data.Body && data.Body.Data) item.card.setProperty("Notes", xmltools.checkString(data.Body.Data));
            else item.card.setProperty("Notes", "");
        }


        //take care of categories and children
        let containers = [];
        containers.push(["Categories", "Category"]);
        containers.push(["Children", "Child"]);
        for (let c=0; c < containers.length; c++) {
            if (data[containers[c][0]] && data[containers[c][0]][containers[c][1]]) {
                let cats = [];
                if (Array.isArray(data[containers[c][0]][containers[c][1]])) cats = data[containers[c][0]][containers[c][1]];
                else cats.push(data[containers[c][0]][containers[c][1]]);
                
                item.card.setProperty(containers[c][0], this.categoriesToString(cats));
            }
        }

        //take care of unmapable EAS option (Contact)
        for (let i=0; i < this.unused_EAS_properties.length; i++) {
            if (data[this.unused_EAS_properties[i]]) item.card.setProperty("EAS-" + this.unused_EAS_properties[i], data[this.unused_EAS_properties[i]]);
        }


        //further manipulations
        if (tbSync.db.getAccountSetting(syncdata.account, "displayoverride") == "1") {
           item.card.setProperty("DisplayName", item.card.getProperty("FirstName", "") + " " + item.card.getProperty("LastName", ""));

            if (item.card.getProperty("DisplayName", "" ) == " " )
                item.card.setProperty("DisplayName", item.card.getProperty("Company", item.card.getProperty("PrimaryEmail", "")));
        }
        
    },









    // --------------------------------------------------------------------------- //
    //read TB event and return its data as WBXML
    // --------------------------------------------------------------------------- //
    getWbxmlFromThunderbirdItem: function (item, syncdata, isException = false) {
        let asversion = tbSync.db.getAccountSetting(syncdata.account, "asversion");
        let wbxml = tbSync.wbxmltools.createWBXML("", syncdata.type); //init wbxml with "" and not with precodes, and set initial codepage
        let nowDate = new Date();


        //loop over all known TB properties which map 1-to-1 (send empty value if not set)
        for (let p=0; p < this.TB_properties.length; p++) {            
            let TB_property = this.TB_properties[p];
            let EAS_property = this.map_TB_properties_to_EAS_properties[TB_property];            
            let value = item.card.getProperty(TB_property,"");
            if (value) wbxml.atag(EAS_property, value);
        }


        //take care of birthday and anniversary
        let dates = [];
        dates.push(["Birthday", "BirthDay", "BirthMonth", "BirthYear"]);
        dates.push(["Anniversary", "AnniversaryDay", "AnniversaryMonth", "AnniversaryYear"]);        
        for (let p=0; p < dates.length; p++) {
            let year = item.card.getProperty(dates[p][3], "");
            let month = item.card.getProperty(dates[p][2], "");
            let day = item.card.getProperty(dates[p][1], "");
            if (year && month && day) {
                //set
                if (month.length<2) month="0"+month;
                if (day.length<2) day="0"+day;
                wbxml.atag(dates[p][0], year + "-" + month + "-" + day + "T00:00:00.000Z");
            }
        }


        //take care of multiline address fields
        let streets = [];
        let seperator = String.fromCharCode(tbSync.db.getAccountSetting(syncdata.account,"seperator")); // options are 44 (,) or 10 (\n)
        streets.push(["HomeAddressStreet", "HomeAddress", "HomeAddress2"]); //EAS, TB1, TB2
        streets.push(["BusinessAddressStreet", "WorkAddress", "WorkAddress2"]);
        streets.push(["OtherAddressStreet", "OtherAddress", "OtherAddress2"]);        
        for (let p=0; p < streets.length; p++) {
            let values = [];
            let s1 = item.card.getProperty(streets[p][1], "");
            let s2 = item.card.getProperty(streets[p][2], "");
            if (s1) values.push(s1);
            if (s2) values.push(s2);
            if (values.length>0) wbxml.atag(streets[p][0], values.join(seperator));            
        }


        //take care of photo
        if (item.card.getProperty("PhotoType", "") == "file") {
            wbxml.atag("Picture", tbSync.getphoto(item.card));                    
        }
        
        
        //take care of unmapable EAS option
        for (let i=0; i < this.unused_EAS_properties.length; i++) {
            let value = item.card.getProperty("EAS-" + this.unused_EAS_properties[i], "");
            if (value) wbxml.atag(this.unused_EAS_properties[i], value);
        }


        //take care of categories and children
        let containers = [];
        containers.push(["Categories", "Category"]);
        containers.push(["Children", "Child"]);
        for (let c=0; c < containers.length; c++) {
            let cats = item.card.getProperty(containers[c][0], "");
            if (cats) {
                let catsArray = this.categoriesFromString(cats);
                wbxml.otag(containers[c][0]);
                for (let ca=0; ca < catsArray.length; ca++) wbxml.atag(containers[c][1], catsArray[ca]);
                wbxml.ctag();            
            }
        }

        //take care of notes - SWITCHING TO AirSyncBase (if 2.5, we still need Contact group here!)
        let description = item.card.getProperty("Notes", "");
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


        //take care of Contacts2 group - SWITCHING TO CONTACTS2
        wbxml.switchpage("Contacts2");

        //loop over all known TB properties of EAS group Contacts2 (send empty value if not set)
        for (let p=0; p < this.TB_properties2.length; p++) {            
            let TB_property = this.TB_properties2[p];
            let EAS_property = this.map_TB_properties_to_EAS_properties2[TB_property];
            let value = item.card.getProperty(TB_property,"");
            if (value) wbxml.atag(EAS_property, value);
        }


        return wbxml.getBytes();
    }
    
}

eas.sync.Contacts.TB_properties = Object.keys(eas.sync.Contacts.map_TB_properties_to_EAS_properties);
eas.sync.Contacts.TB_properties2 = Object.keys(eas.sync.Contacts.map_TB_properties_to_EAS_properties2);
