using Microsoft.MetadirectoryServices;
using System.Collections.Generic;
using System.Xml.Linq;
using System;

namespace FimSync_Ezma
{
    class personPosition
    {
        // string definions
        string PersonRelationID, Arbeidsforhold_status,
            Arbeidsforhold_nummer, OrgUnitName, OrgUnitId_person, IsHovedarbeidsforhold, Startdato_stilling, Sluttdato_Stilling, AnsattforholdsKode,
            Stillingskode_Kodeverdi, Stillingskode_Beskrivelse;
        
        // Namespaces used in XML
        XNamespace v31 = "http://ansattlist.bluegarden.no/object/v3_1";
        XNamespace v1 = "http://kodeverk.bluegarden.no/object/v1";
        #region Person
        //public Person(XElement person, Dictionary<string, Tuple<List<string>, string>> groups)
        //{
        //    // PersonID must always be present
        //    PersonID = person.Element(v31 + "PersonID").Value;
            
        //    // All below is in try - catch as they're not always present
        //    // SSN
        //    try
        //    {
        //        Fodselsnummer = person.Element(v31 + "Fodselsnummer").Value;
        //    }
        //    catch
        //    { }

            
        //    // Signature
        //    try
        //    {
        //        Signatur = person.Element(v31 + "Signatur").Value;
        //    }
        //    catch
        //    { }
            

        //    // Givenname
        //    try
        //    {
        //        Fornavn = person.Element(v31 + "Fornavn").Value;
        //    } 
        //    catch
        //    { }

        //    // Surname
        //    try
        //    {
        //        Etternavn = person.Element(v31 + "Etternavn").Value;
        //    } 
        //    catch
        //    { }

        //    //try
        //    //{
        //    //    UserId = person.Element(v31 + "UserId").Value;
        //    //}
        //    //catch
        //    //{ }

        //    // Address line 1 - private
        //    try
        //    {
        //        AdresseLinje1_Bosted = person.Element(v31 + "AdresseLinje1_Bosted").Value;
        //    }
        //    catch
        //    { }

        //    // Address line 2 - private
        //    try
        //    {
        //        AdresseLinje2_Bosted = person.Element(v31 + "AdresseLinje2_Bosted").Value;
        //    }
        //    catch
        //    { }

        //    // Address line 1 - At work
        //    try
        //    {
        //        AdresseLinje1_Intern = person.Element(v31 + "AdresseLinje1_Intern").Value;
        //    }
        //    catch
        //    { }

        //    // Address line 2 - At work
        //    try
        //    {
        //        AdresseLinje2_Intern = person.Element(v31 + "AdresseLinje2_Intern").Value;
        //    }
        //    catch
        //    { }

        //    // ZIP number private
        //    try
        //    {
        //        Postnummer_Bosted = person.Element(v31 + "Postnummer_Bosted").Value;
        //    }
        //    catch
        //    { }

        //    // City address - private
        //    try
        //    {
        //        Poststed_Bosted = person.Element(v31 + "Poststed_Bosted").Value;
        //    }
        //    catch
        //    { }

        //    // ZIP address at Work
        //    try
        //    {
        //        Postnummer_Intern = person.Element(v31 + "Postnummer_Intern").Value;
        //    }
        //    catch
        //    { }

        //    // city address - at work
        //    try
        //    {
        //        Poststed_Intern = person.Element(v31 + "Poststed_Intern").Value;
        //    }
        //    catch
        //    { }

        //    // Employee relation status
        //    try
        //    {
        //        Arbeidsforhold_status = person.Attribute("status").Value;
        //    }
        //    catch
        //    { }

        //    // Employee realtion number
        //    try
        //    {
        //        Arbeidsforhold_nummer = person.Element(v31 + "Arbeidsforholdnummer").Value;
        //    }
        //    catch
        //    { }

        //    // OrgUnitName
        //    try
        //    {
        //        OrgUnitName = person.Element(v31 + "OrgUnitName").Value;
        //    }
        //    catch
        //    { }

        //    // OrgUnitID - pr person
        //    try
        //    {
        //        OrgUnitId_person = person.Element(v31 + "OrgUnitId").Value;
        //    }
        //    catch
        //    { }

        //    // Email address
        //    try
        //    {
        //        Email = person.Element(v31 + "Email").Value;
        //    }
        //    catch
        //    { }

        //    // Employer
        //    try
        //    {
        //        Arbeidsgiver = person.Element(v31 + "Arbeidsgiver").Value;
        //    }
        //    catch
        //    { }

        //    // Employer number
        //    try
        //    {
        //        ArbeidsgiverNummer = person.Element(v31 + "ArbeidsgiverNummer").Value;
        //    }
        //    catch
        //    { }

        //    // Is main employed relation. Boolean
        //    try
        //    {
        //        IsHovedarbeidsforhold = person.Element(v31 + "IsHovedarbeidsforhold").Value;
        //    }
        //    catch
        //    { }

        //    // StartDate position
        //    try
        //    {
        //        Startdato_stilling = person.Element(v31 + "Startdato").Value;
        //    }
        //    catch
        //    { }

        //    // EndDate position
        //    try
        //    {
        //        Sluttdato_Stilling = person.Element(v31 + "Sluttdato").Value;
        //    }
        //    catch
        //    { }

        //    // Employer realtion code
        //    try
        //    {
        //        AnsattforholdsKode = person.Element(v31 + "AnsattforholdsKode").Value;
        //    }
        //    catch
        //    { }

        //    // Position code value
        //    try
        //    {
        //        Stillingskode_Kodeverdi = person.Element(v31 + "Stillingskode").Element(v1 + "Kodeverdi").Value;
        //    }
        //    catch
        //    { }

        //    // Position code description
        //    try
        //    {
        //        Stillingskode_Beskrivelse = person.Element(v31 + "Stillingskode").Element(v1 + "Beskrivelse").Value;
        //    }
        //    catch
        //    { }

        //    // Mobile phone
        //    try
        //    {
        //        Mobiltelefon = person.Element(v31 + "Mobiltelefon").Value;
        //    }
        //    catch
        //    { }
            
        //    // Employee Number
        //    try
        //    {
        //        AnsattNummer = person.Element(v31 + "AnsattNummer").Value;
        //    }
        //    catch
        //    { }

        //    // StartDate - Person
        //    try
        //    {
        //        StartDato_person = person.Element(v31 + "StartDato").Value;
        //    }
        //    catch
        //    { }

        //    // EndDate - Person
        //    try
        //    {
        //        SluttDato_person = person.Element(v31 + "SluttDato").Value;
        //    }
        //    catch
        //    { }

        //    /*
             
        //    if (person.Element(v31 + "Kontaktinformasjon").Attribute("kontaktinformasjonType").Value == "Bostedsadresse")
        //    {
        //        XElement homeaddressinfo = person.Element(v31 + "Kontaktinformasjon");
        //        try
        //        {
        //            privatecountryCode = person.Element(v31 + "homeAddress").Element(v31 + "countryCode").Value;
        //        }
        //        catch
        //        { }

        //        try
        //        {
        //            privatecountryName = person.Element(v31 + "homeAddress").Element(v31 + "countryName").Value;
        //        }
        //        catch
        //        { }

        //        try
        //        {
        //            privateStreetAddress = person.Element(v31 + "privateStreetAddress").Value;
        //        }
        //        catch
        //        { }

        //    }

            
           

        //    try
        //    {
        //        postalAddress = person.Element(v31 + "homeAddress").Element(v31 + "postalAddress").Value;
        //    }
        //    catch
        //    { }

            

        //    try
        //    {
        //        postalCode = person.Element(v31 + "homeAddress").Element(v31 + "postalCode").Value;
        //    }
        //    catch
        //    { }

           

        //    */

        //    #region CreateGroups
        //    // Create unit groups and save to groups dictionary 
        //    try
        //    {
        //        if (!groups.ContainsKey(OrgUnitId_person))
        //        {
        //            groups.Add(OrgUnitId_person, Tuple.Create(new List<string>(), OrgUnitName));
        //            if (!groups[OrgUnitId_person].Item1.Contains(PersonID))
        //            {
        //                groups[OrgUnitId_person].Item1.Add(PersonID);
        //            }
        //        }
        //        else
        //        {
        //            if (!groups[OrgUnitId_person].Item1.Contains(PersonID))
        //            {
        //                groups[OrgUnitId_person].Item1.Add(PersonID);
        //            }
        //        }
        //    }
        //    catch
        //    { }

        //    /*
        //    // Create businessNumber groups and save to groups dictionary 
        //    try
        //    {
        //        XElement _positionStatistics = position.Element(v31 + "positionStatistics");
        //        XElement _businessNumber = _positionStatistics.Element(v31 + "businessNumber");
        //        grpName = _businessNumber.Attribute("name").Value;
        //        grpId = _businessNumber.Attribute("value").Value;


        //        if (!groups.ContainsKey(grpId))
        //        {
        //            groups.Add(grpId, Tuple.Create(new List<string>(), grpName, "orgUnit"));
        //            if (!groups[grpId].Item1.Contains(personIdHRM))
        //            {
        //                groups[grpId].Item1.Add(personIdHRM);
        //            }
        //        }
        //        else
        //        {
        //            if (!groups[grpId].Item1.Contains(personIdHRM))
        //            {
        //                groups[grpId].Item1.Add(personIdHRM);
        //            }
        //        }
        //    }
        //    catch
        //    { }

        //    // Create company groups and save to groups dictionary 
        //    try
        //    {
        //        XElement _positionStatistics = position.Element(v31 + "positionStatistics");
        //        XElement _businessNumber = _positionStatistics.Element(v31 + "companyNumber");
        //        grpName = _businessNumber.Attribute("name").Value;
        //        grpId = _businessNumber.Attribute("value").Value;


        //        if (!groups.ContainsKey(grpId))
        //        {
        //            groups.Add(grpId, Tuple.Create(new List<string>(), grpName, "company"));
        //            if (!groups[grpId].Item1.Contains(personIdHRM))
        //            {
        //                groups[grpId].Item1.Add(personIdHRM);
        //            }
        //        }
        //        else
        //        {
        //            if (!groups[grpId].Item1.Contains(personIdHRM))
        //            {
        //                groups[grpId].Item1.Add(personIdHRM);
        //            }
        //        }
        //    }
        //    catch
        //    { }
        //    */
        //    #endregion
        //}

        #endregion
        public personPosition(XElement position, string PersonID)
        {
            
            OrgUnitId_person = position.Element(v31 + "OrgUnitId").Value;

            PersonRelationID = PersonID + OrgUnitId_person;


            // Employee relation status
            try
            {
                Arbeidsforhold_status = position.Attribute("status").Value;
            }
            catch
            { }

            // Employee realtion number
            try
            {
                Arbeidsforhold_nummer = position.Element(v31 + "Arbeidsforholdnummer").Value;
            }
            catch
            { }

            // OrgUnitName
            try
            {
                OrgUnitName = position.Element(v31 + "OrgUnitName").Value;
            }
            catch
            { }

            // Is main employed relation. Boolean
            try
            {
                IsHovedarbeidsforhold = position.Element(v31 + "IsHovedarbeidsforhold").Value;
            }
            catch
            { }

            // StartDate position
            try
            {
                Startdato_stilling = position.Element(v31 + "Startdato").Value;
            }
            catch
            { }

            // EndDate position
            try
            {
                Sluttdato_Stilling = position.Element(v31 + "Sluttdato").Value;
            }
            catch
            { }

            // Employer realtion code
            try
            {
                AnsattforholdsKode = position.Element(v31 + "AnsattforholdsKode").Value;
            }
            catch
            { }

            // Position code value
            try
            {
                Stillingskode_Kodeverdi = position.Element(v31 + "Stillingskode").Element(v1 + "Kodeverdi").Value;
            }
            catch
            { }

            // Position code description
            try
            {
                Stillingskode_Beskrivelse = position.Element(v31 + "Stillingskode").Element(v1 + "Beskrivelse").Value;
            }
            catch
            { }
        }

        // Returns CSEntryChange for use by the FIM Sync Engine directly
        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "position";

            // Unique ID. Must always be present
            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("PersonRelationID", PersonRelationID));


            
            // Employee relation Status
            if (Arbeidsforhold_status != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Arbeidsforhold_status", Arbeidsforhold_status));

            // Employee relation Number
            if (Arbeidsforhold_nummer != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Arbeidsforhold_nummer", Arbeidsforhold_nummer));

            // Employee relation code
            if (AnsattforholdsKode != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("AnsattforholdsKode", AnsattforholdsKode));

            // OrgUnitName
            if (OrgUnitName != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("OrgUnitName", OrgUnitName));

            // OrgUnitId person
            if (OrgUnitId_person != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("OrgUnitId_person", OrgUnitId_person));

            // Is main employed relation
            if (IsHovedarbeidsforhold != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("IsHovedarbeidsforhold", IsHovedarbeidsforhold));

            // Position code value
            if (Stillingskode_Kodeverdi != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Stillingskode_Kodeverdi", Stillingskode_Kodeverdi));

            // Position code description
            if (Stillingskode_Beskrivelse != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Stillingskode_Beskrivelse", Stillingskode_Beskrivelse));

            if (Startdato_stilling != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Startdato_stilling", Startdato_stilling));

            if (Sluttdato_Stilling != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Sluttdato_Stilling", Sluttdato_Stilling));

            /*

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Stillinger", 
                Arbeidsforhold_status + "|" +                                            
                Arbeidsforhold_nummer + "|" +
                AnsattforholdsKode + "|" +
                OrgUnitName + "|" +
                OrgUnitId_person + "|" +
                IsHovedarbeidsforhold + "|" +
                Stillingskode_Kodeverdi + "|" +
                Stillingskode_Beskrivelse + "|" +
                Startdato_stilling + "|" +
                Sluttdato_Stilling));
            */


            return csentry;

        }
    }
}
