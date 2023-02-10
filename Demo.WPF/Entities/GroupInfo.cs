
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupMigrationPnP.Entities
{
    public class GroupInfo
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Owner { get; set; }
        public string Permission { get; set; }
        //public RoleAssignment Role { get; set; }
        public List<GroupUser> Users { get; set; } = new List<GroupUser>();
    }
}
