using GroupMigrationPnP.ConfigDetails;
using PnP.Core.Model.Security;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupMigrationPnP.HelperMethods
{
    public static class SiteConnection
    {
        public static Uri GenerateURIFromString(string inputValue)
        {
            if (Uri.IsWellFormedUriString(inputValue, UriKind.Absolute))
            {
                return new Uri(inputValue, UriKind.Absolute);
            }
            return null;
        }

        public static string GetSiteConfigurationDetails(Uri enteredURL)
        {
            if (enteredURL.Host.ToLower().Equals(TenantConfigMaster.TargetSiteURL))
                return TenantConfigMaster.targetDomainName;
            else if (enteredURL.Host.ToLower().Equals(TenantConfigMaster.sourceSiteURL))
                return TenantConfigMaster.sourceDomainName;
            return "";
        }

        public static IQueryable<ISharePointGroup> GetGroups(PnPContext ctx)
        {
            var lists = (from l in ctx.Web.SiteGroups.QueryProperties(l => l.OwnerTitle, l => l.Title, l => l.Description, l => l.IsHiddenInUI)
                         select l);
            return lists;
        }
    }
}
