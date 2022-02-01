using Microsoft.MetadirectoryServices;
using System.Xml.Linq;
using System.Collections.Generic;
using System;

namespace FimSync_Ezma
{
    class UnitCSV
    {
        string unitid, Name, Owner_ID, Leader, Active, Costs, Persons, Orgnr_company, Orgjoin, Cost_carrier, parent_unit;
        

        public UnitCSV(XElement unit, Dictionary<string, Tuple<string, string, string>> hiracyParents)
        

        {
            unitid = unit.Element("ID").Value;

            try
            {
                Name = unit.Element("Navn").Value;
            }
            catch { }

            try
            {
                Owner_ID = unit.Element("Eier_ID").Value;
                
            }
            catch { }

            try { Leader = unit.Element("Leder").Value;
                hiracyParents.Add(unitid, Tuple.Create(Owner_ID,Name,Leader));
            }
            catch { }
            try { Active = unit.Element("Aktiv").Value; }
            catch { }
            try { Costs = unit.Element("Kostnader").Value; }
            catch { }
            try { Persons = unit.Element("Personer").Value; }
            catch { }
            try { Orgnr_company = unit.Element("Orgnr_bedrift").Value; }
            catch { }
            try { Orgjoin = unit.Element("Orgledd").Value; }
            catch { }
            try { Cost_carrier = unit.Element("Kostnadsbaerer").Value; }
            catch { }
            try { parent_unit = unit.Element("parent_unit").Value; }
            catch { }
            

        }

        // Returns CSEntryChange for use by the FIM Sync Engine directly
        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "unit";

            //Henter unik ID
            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("id", unitid));
            
            if (Name != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Navn", Name));
            if (Owner_ID != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Eier_ID", Owner_ID));
            if (Leader != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Leder", Leader));
            if (Active != null)
                if(Active == "J")
                { csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Aktiv", true)); }
                else { csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Aktiv", false)); }
            if (Costs != null)
                if(Costs == "J")
                { csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Kostnader", true)); }
                else { csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Kostnader", false)); }
            if (Persons != null)
                if(Persons == "J")
                { csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Personer", true)); }
                else { csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Personer", false)); }
            if (Orgnr_company != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Orgnr_bedrift", Orgnr_company));
            if (Orgjoin != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Orgledd", Orgjoin));
            if (Cost_carrier != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("Kostnadsbaerer", Cost_carrier));
            if (parent_unit != null)
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("parent_unit", parent_unit));

            return csentry;
        }
    }
}
