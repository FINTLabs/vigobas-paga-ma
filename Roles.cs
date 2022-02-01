using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace FimSync_Ezma
{
    public class Roles
    {

        public class PersonRole
        {
            public string Source { get; set; }
            public string Department { get; set; }
            public string RoleId { get; set; }
            public string SubRoleId { get; set; }
        }
    }
}