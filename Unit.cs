using Microsoft.MetadirectoryServices;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace FimSync_Ezma
{
    class Unit
    {
        string orgunitid, OrgNavn, GroupId, ParentGroupId, ErAktiv, _OUstring;

        public Unit(XElement unit, string OUstring)
        {
            XNamespace v311 = "http://bluegarden.no/organisation/structure/object/v31";

            orgunitid = unit.Element(v311 + "OrgUnitId").Value;
            OrgNavn = unit.Element(v311 + "OrgNavn").Value;
            GroupId = unit.Element(v311 + "GroupId").Value;
            ParentGroupId = unit.Element(v311 + "ParentGroupId").Value;
            ErAktiv = unit.Element(v311 + "ErAktiv").Value;
            try
            {
                _OUstring = OUstring;
            }
            catch
            { }
        }

        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "unit";
            
            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("orgunitid", orgunitid));
            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("OrgNavn", OrgNavn));
            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("UnitId", GroupId));
            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("ParentGroupId", ParentGroupId));
            csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("ErAktiv", ErAktiv));

            if (_OUstring != null)
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("unitOU", _OUstring));
            }
            return csentry;
        }
    }
}
