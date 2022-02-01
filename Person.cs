using Microsoft.MetadirectoryServices;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System;
using Newtonsoft.Json;

namespace FimSync_Ezma
{
    class Person
    {
        // string definions
        string PersonID, Fornavn, Etternavn, Fodselsnummer, Signatur, AnsattNummer, _OUstring, _UserRole,
            Mobiltelefon, Telefon_Intern, Telefon_Bosted, StartDato_person, SluttDato_person, Email, Arbeidsgiver, ArbeidsgiverNummer,
            AdresseLinje1_Bosted, AdresseLinje2_Bosted, AdresseLinje1_Intern, AdresseLinje2_Intern,
            Postnummer_Bosted, Poststed_Bosted, Postnummer_Intern, Poststed_Intern, Pri_Arbeidsforhold_status,
            Pri_Arbeidsforhold_nummer, Pri_OrgUnitName, OrgUnitName, Pri_OrgUnitIdRef, Pri_OrgUnitId ,Leder, OrgUnitId,
            Pri_Startdato_stilling, Pri_Stoppdato_Stilling, Pri_AnsattforholdsKode,
            Pri_Stillingskode_Kodeverdi, Pri_Stillingskode_Beskrivelse;


        List<string> Sec_Arbeidsforhold_status = new List<string>();
        List<string> Sec_Arbeidsforhold_nummer = new List<string>();
        List<string> Sec_OrgUnitName = new List<string>();
        List<string> Sec_Startdato_stilling = new List<string>();
        List<string> Sec_AnsattforholdsKode = new List<string>();
        List<string> Sec_Stillingskode_Kodeverdi = new List<string>();
        List<string> Sec_Stillingskode_Beskrivelse = new List<string>();
        List<string> Sec_OrgUnitId = new List<string>();



        // Namespaces used in XML
        XNamespace v31 = "http://ansattlist.bluegarden.no/object/v3_1";
        XNamespace v1 = "http://kodeverk.bluegarden.no/object/v1";


        public Person(XElement person, Dictionary<string, Tuple<List<string>, string>> groups, Dictionary<string, List<XElement>> positions, string OUstring, Dictionary<string, string> leadersID)
        {
            // PersonID must always be present
            PersonID = person.Element(v31 + "PersonID").Value.Replace("#", "");

            // Set member and role
            Roles.PersonRole role = new Roles.PersonRole();
            role.Source = "PAGA MA";
            role.Department = "HFK";
            role.RoleId = "employee";
            role.SubRoleId = "staff";

            _UserRole = JsonConvert.SerializeObject(role);

            // All below is in try - catch as they're not always present
            // Add OUString to person
            try
            {
                _OUstring = OUstring;
            }
            catch
            { }

            // SSN
            try
            {
                Fodselsnummer = person.Element(v31 + "Fodselsnummer").Value;
            }
            catch
            { }

            // Signature (UserID)
            try
            {
                Signatur = person.Element(v31 + "Signatur").Value;
            }
            catch
            { }


            // Givenname
            try
            {
                Fornavn = person.Element(v31 + "Fornavn").Value;
            }
            catch
            { }

            // Surname
            try
            {
                Etternavn = person.Element(v31 + "Etternavn").Value;
            }
            catch
            { }

            foreach (XElement contactInfo in person.Elements(v31 + "Kontaktinformasjon"))
            {
                switch (contactInfo.Attribute("kontaktinformasjonType").Value)
                {
                    // Private information
                    case "Bostedsadresse":
                        try
                        {
                            AdresseLinje1_Bosted = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Adresse").Element(v31 + "AdresseLinje1").Value;
                        }
                        catch { }

                        try
                        {
                            AdresseLinje2_Bosted = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Adresse").Element(v31 + "AdresseLinje2").Value;
                        }
                        catch { }

                        try
                        {
                            Postnummer_Bosted = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Adresse").Element(v31 + "Postnummer").Value;
                        }
                        catch { }

                        try
                        {
                            Poststed_Bosted = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Adresse").Element(v31 + "Poststed").Value;
                        }
                        catch { }

                        try
                        {
                            Telefon_Bosted = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Telefon").Value;
                        }
                        catch { }

                        break;

                    // Work information
                    case "Internadresse":
                        try
                        {
                            AdresseLinje1_Intern = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Adresse").Element(v31 + "AdresseLinje1").Value;
                        }
                        catch { }

                        try
                        {
                            AdresseLinje2_Intern = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Adresse").Element(v31 + "AdresseLinje2").Value;
                        }
                        catch { }

                        try
                        {
                            Postnummer_Intern = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Adresse").Element(v31 + "Postnummer").Value;
                        }
                        catch { }

                        try
                        {
                            Poststed_Intern = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Adresse").Element(v31 + "Poststed").Value;
                        }
                        catch { }

                        

                        try
                        {
                            Telefon_Intern = person.Element(v31 + "Kontaktinformasjon").Element(v31 + "Telefon").Value;
                        }
                        catch { }

                        break;
                }
            }

               
            foreach (XElement position in positions[PersonID])
            {
                if(position.Element(v31 + "IsHovedarbeidsforhold").Value == "J")
                {
                    // Employee relation status
                    try
                    {
                        Pri_Arbeidsforhold_status = position.Attribute("status").Value;
                    }
                    catch
                    { }

                    // OrgUnitID - pr person
                    try
                    {
                        Pri_OrgUnitIdRef = position.Element(v31 + "OrgUnitId").Value;
                        Pri_OrgUnitId = position.Element(v31 + "OrgUnitId").Value;
                    }
                    catch
                    { }

                    // Employee realtion number
                    try
                    {
                        Pri_Arbeidsforhold_nummer = position.Element(v31 + "Arbeidsforholdnummer").Value;
                    }
                    catch
                    { }

                    // Leader
                    try
                    {

                        Leder = position.Element(v31 + "Leder").Value;
                        Leder = leadersID[Leder];
                    }
                    catch
                    {

                    }

                    // OrgUnitName
                    try
                    {
                        Pri_OrgUnitName = position.Element(v31 + "OrgUnitName").Value;
                    }
                    catch
                    { }

                    // StartDate position
                    try
                    {
                        Pri_Startdato_stilling = position.Element(v31 + "Startdato").Value;
                    }
                    catch
                    { }

                    // EndDate position
                    try
                    {
                        Pri_Stoppdato_Stilling = position.Element(v31 + "Sluttdato").Value;
                    }
                    catch
                    { }

                    // Employer realtion code
                    try
                    {
                        Pri_AnsattforholdsKode = position.Element(v31 + "AnsattforholdsKode").Value.TrimStart().TrimEnd();
                    }
                    catch
                    { }

                    // Position code value
                    try
                    {
                        Pri_Stillingskode_Kodeverdi = position.Element(v31 + "Stillingskode").Element(v1 + "Kodeverdi").Value;
                    }
                    catch
                    { }

                    // Position code description
                    try
                    {
                        Pri_Stillingskode_Beskrivelse = position.Element(v31 + "Stillingskode").Element(v1 + "Beskrivelse").Value;
                    }
                    catch
                    { }
                }
                else
                {
                    // Employee relation status where is not primary

                    if (position.Attribute("status").Value == "Aktiv")
                    {
                        string id = null;
                        // Employee realtion number
                        try
                        {
                            Sec_Arbeidsforhold_nummer.Add(position.Element(v31 + "Arbeidsforholdnummer").Value);
                            // Set work position number # as prefix to the other strings. Easier to seprarate the positions later
                            id = position.Element(v31 + "Arbeidsforholdnummer").Value + "|";
                        }
                        catch
                        { }

                        // OrgUnitName
                        try
                        {
                            Sec_OrgUnitName.Add(id + position.Element(v31 + "OrgUnitName").Value);
                        }
                        catch
                        { }

                        
                        // OrgUnitId
                        try
                        {
                            Sec_OrgUnitId.Add(id + position.Element(v31 + "OrgUnitId").Value);
                        }
                        catch
                        { }

                        // Employer realtion code
                        try
                        {
                            Sec_AnsattforholdsKode.Add(id + position.Element(v31 + "AnsattforholdsKode").Value);
                        }
                        catch
                        { }

                        // Position code value
                        try
                        {
                            Sec_Stillingskode_Kodeverdi.Add(id +  position.Element(v31 + "Stillingskode").Element(v1 + "Kodeverdi").Value);
                        }
                        catch
                        { }

                        // Position code description
                        try
                        {
                            Sec_Stillingskode_Beskrivelse.Add(id +  position.Element(v31 + "Stillingskode").Element(v1 + "Beskrivelse").Value);
                        }
                        catch
                        { }
                    }
                }
            }
            
            // Email address
            try
            {
                Email = person.Element(v31 + "Email").Value;
            }
            catch
            { }

            // Employer
            try
            {
                Arbeidsgiver = person.Element(v31 + "Arbeidsgiver").Value;
            }
            catch
            { }

            // Employer number
            try
            {
                ArbeidsgiverNummer = person.Element(v31 + "ArbeidsgiverNummer").Value;
            }
            catch
            { }


            // Mobile phone
            try
            {
                Mobiltelefon = person.Element(v31 + "Mobiltelefon").Value;
            }
            catch
            { }
            
            // Employee Number
            try
            {
                AnsattNummer = person.Element(v31 + "AnsattNummer").Value;
            }
            catch
            { }

            // StartDate - Person
            try
            {
                StartDato_person = person.Element(v31 + "StartDato").Value;
            }
            catch
            { }

            // EndDate - Person
            try
            {
                SluttDato_person = person.Element(v31 + "SluttDato").Value;
            }
            catch
            { }

            
            #region CreateGroups
            // OrgUnitID - Use in groups add
            try
            {
                OrgUnitId = person.Element(v31 + "Arbeidsforhold").Element(v31 + "OrgUnitId").Value;
            }
            catch
            { }

            // OrgUnitName - Use in groups add
            try
            {
                OrgUnitName = person.Element(v31 + "Arbeidsforhold").Element(v31 + "OrgUnitName").Value;
            }
            catch
            { }

            // Create groups and save to groups dictionary
            try
            {
                if (!groups.ContainsKey(OrgUnitId))
                {
                    groups.Add(OrgUnitId, Tuple.Create(new List<string>(), OrgUnitName));
                    if (!groups[OrgUnitId].Item1.Contains(PersonID))
                    {
                        groups[OrgUnitId].Item1.Add(PersonID);
                    }
                }
                else
                {
                    if (!groups[OrgUnitId].Item1.Contains(PersonID))
                    {
                        groups[OrgUnitId].Item1.Add(PersonID);
                    }
                }
            }
            catch
            { }
            

            #endregion
        }

       

        // Returns CSEntryChange for use by the FIM Sync Engine directly
        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "person";

            // Unique ID. Must always be present
            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("PersonID", PersonID));

            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("UserRole", _UserRole));

            // OU-string
            if (_OUstring != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("PersonOU", _OUstring));
                        
            // Givenname
            if (Etternavn != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Etternavn", Etternavn));

            // Surname
            if (Fornavn != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Fornavn", Fornavn));

            // SSN
            if (Fodselsnummer != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Fodselsnummer", Fodselsnummer));
            
            // EmployeeNumber
            if (AnsattNummer != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("AnsattNummer", AnsattNummer));

            // UserId
            if (Signatur != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("UserId", Signatur));

            // Signature
            if (Signatur != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Signatur", Signatur));

            // Email
            if (Email != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Email", Email));

            // Startdate Person
            if (StartDato_person != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("StartDato_person", StartDato_person));

            // EndDate Person
            if (SluttDato_person != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("SluttDato_person", SluttDato_person));

            // Mobile phone
            if (Mobiltelefon != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Mobiltelefon", Mobiltelefon));

            // Employer
            if (Arbeidsgiver != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Arbeidsgiver", Arbeidsgiver));

            // EmployerNumber
            if (ArbeidsgiverNummer != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("ArbeidsgiverNummer", ArbeidsgiverNummer));

            // Street address Line 1 private
            if (AdresseLinje1_Bosted != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("AdresseLinje1_Bosted", AdresseLinje1_Bosted));

            // Street address Line 2 private
            if (AdresseLinje2_Bosted != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("AdresseLinje2_Bosted", AdresseLinje2_Bosted));

            // Phone - private
            if (Telefon_Intern != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Telefon_Intern", Telefon_Intern));

            // ZIP Address private
            if (Postnummer_Bosted != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Postnummer_Bosted", Postnummer_Bosted));

            // City Address private
            if (Poststed_Bosted != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Poststed_Bosted", Poststed_Bosted));

            // Street Address Line 1 at work
            if (AdresseLinje1_Intern != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("AdresseLinje1_Intern", AdresseLinje1_Intern));

            // Street Address Line 2 at work
            if (AdresseLinje2_Intern != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("AdresseLinje2_Intern", AdresseLinje2_Intern));

            // ZIP number at Work
            if (Postnummer_Intern != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Postnummer_Intern", Postnummer_Intern));

            // City Address at Work
            if (Poststed_Intern != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Poststed_Intern", Poststed_Intern));
            // Phone at Work
            if (Telefon_Bosted != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Telefon_Bosted", Telefon_Bosted));

            // Employee relation Status
            if (Pri_Arbeidsforhold_status != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_Arbeidsforhold_status", Pri_Arbeidsforhold_status));

            // Employee relation Number
            if (Pri_Arbeidsforhold_nummer != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_Arbeidsforhold_nummer", Pri_Arbeidsforhold_nummer));

            // Employee relation code
            if (Pri_AnsattforholdsKode != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_AnsattforholdsKode", Pri_AnsattforholdsKode));

            // OrgUnitName
            if (Pri_OrgUnitName != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_OrgUnitName", Pri_OrgUnitName));

            // OrgUnitId person as refrence
            if (Pri_OrgUnitIdRef != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_OrgUnitIdRef", Pri_OrgUnitIdRef));

            // OrgUnitId person as string
            if (Pri_OrgUnitIdRef != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_OrgUnitId", Pri_OrgUnitId));


            // Leader
            if (Leder != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Leder", Leder));

            // Position code value
            if (Pri_Stillingskode_Kodeverdi != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_Stillingskode_Kodeverdi", Pri_Stillingskode_Kodeverdi));

            // Position code description
            if (Pri_Stillingskode_Beskrivelse != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_Stillingskode_Beskrivelse", Pri_Stillingskode_Beskrivelse));

            if (Pri_Startdato_stilling != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_Startdato_stilling", Pri_Startdato_stilling));

            if (Pri_Stoppdato_Stilling != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Pri_Sluttdato_Stilling", Pri_Stoppdato_Stilling));


            // Other positions
           
            // Define List to be used in CSEntry add
            List<object> seccontent = new List<object>();
            if (Sec_Arbeidsforhold_nummer != null)
            {
                
                foreach (string val in Sec_Arbeidsforhold_nummer)
                {
                    seccontent.Add(val);
                }

                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Sec_Arbeidsforhold_nummer", seccontent));
            }
            
            // Employee relation code
            if (Sec_AnsattforholdsKode != null)
            {
                seccontent = new List<object>();
                foreach (string val in Sec_AnsattforholdsKode)
                {
                    seccontent.Add(val);
                }
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Sec_AnsattforholdsKode", seccontent));
            }
            
            // OrgUnitName
            if (Sec_OrgUnitName != null)
            {
                seccontent = new List<object>();
                foreach (string val in Sec_OrgUnitName)
                {
                    seccontent.Add(val);
                }
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Sec_OrgUnitName", seccontent));
                
            }

            // OrgUnitId person
            if (Sec_OrgUnitId != null)
            {
                seccontent = new List<object>();
                foreach (string val in Sec_OrgUnitId)
                {
                    seccontent.Add(val);
                }
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Sec_OrgUnitId", seccontent));
            }

            // OrgUnitId person
            if (Sec_Startdato_stilling != null)
            {
                seccontent = new List<object>();
                foreach (string val in Sec_OrgUnitId)
                {
                    seccontent.Add(val);
                }
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Sec_Startdato_stilling", seccontent));
            }
            
            // Position code value
            if (Sec_Stillingskode_Kodeverdi != null)
            {
                seccontent = new List<object>();
                foreach (string val in Sec_Stillingskode_Kodeverdi)
                {
                    seccontent.Add(val);
                }
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Sec_Stillingskode_Kodeverdi", seccontent));
            }

            // Position code description
            if (Sec_Stillingskode_Beskrivelse != null)
            {
                seccontent = new List<object>(); 
                foreach (string val in Sec_Stillingskode_Beskrivelse)
                {
                    seccontent.Add(val);
                }
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Sec_Stillingskode_Beskrivelse", seccontent));
            }
            
            return csentry;
        }
    }
}
