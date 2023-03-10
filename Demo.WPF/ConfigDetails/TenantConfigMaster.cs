using PnP.Core.Model.Security;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupMigrationPnP.ConfigDetails
{
    public static class TenantConfigMaster
    {
        public const string sourceSiteURL = "q7q0.sharepoint.com";
        public const string TargetSiteURL = "innovaorgdomain.sharepoint.com";

        public const string sourceDomainName = "Q7Site";
        public const string targetDomainName = "OrgSite";

        public static IQueryable<ISharePointGroup> sourceGroups;
        public static PnPContext sourceContext;
        public static PnPContext destContext;
    }
}
