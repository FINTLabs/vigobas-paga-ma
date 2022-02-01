using Microsoft.MetadirectoryServices;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System;

namespace FimSync_Ezma
{
    class Group
    {
        string group_prefix, group_suffix, _oustring;
        private Dictionary<string, Tuple<List<string>, string>> _groups = new Dictionary<string, Tuple<List<string>, string>>();

        public Group(string groupKey, List<string> groupMembers, string groupName, KeyedCollection<string, ConfigParameter> configparameter, string oustring)
        {
            _groups.Add(groupKey, Tuple.Create(groupMembers, groupName));
            group_prefix = configparameter["Gruppeprefix"].Value;
            group_suffix = configparameter["Gruppesuffix"].Value;
            try
            {
                _oustring = oustring;
            }
            catch
            { }
        }
        
        // Returns CSEntryChange for use by the FIM Sync Engine directly
        internal Microsoft.MetadirectoryServices.CSEntryChange GetCSEntryChange()
        {
            CSEntryChange csentry = CSEntryChange.Create();
            csentry.ObjectModificationType = ObjectModificationType.Add;
            csentry.ObjectType = "group";

            foreach (var _group in _groups)
            {
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("groupId", group_prefix + _group.Key + group_suffix));
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("groupName", _groups[_group.Key].Item2.ToString()));
                //csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("groupDescription", _groups[_group.Key].Item3.ToString()));
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("groupSourceId", _group.Key));
                if (_oustring != null)
                {
                    csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("groupOU", _oustring));
                }
                
                IList<object> members = new List<object>();
                foreach (var member in _groups[_group.Key].Item1)
                {
                    members.Add(member.ToString());
                }
                csentry.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("members", members));
            }
            return csentry;
        }
    }
}
