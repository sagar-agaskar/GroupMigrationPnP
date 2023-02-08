using GroupMigrationPnP.ConfigDetails;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;
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
        public static IQueryable<IList> GetLists(PnPContext ctx)
        {
            ctx.Web.LoadAsync(p => p.Title, p => p.Lists);
            foreach (var list in ctx.Web.Lists.AsRequested())
            { }
                var lists = (from l in ctx.Web.Lists.QueryProperties(l => l.Title, l => l.Description)
                         select l);
            return lists;
        }

        #region basic details fetch 
        //await context.Web.LoadAsync(p => p.SiteGroups.QueryProperties(u => u.Users, u=>u.Title, u=>u.OwnerTitle));

        //            foreach (var group in context.Web.SiteGroups.AsRequested())
        //            {
        //                // do something with the group
        //            }

        //            await context.Web.LoadAsync(p => p.ContentTypes.QueryProperties(p=>p.Name, p=>p.Hidden));

        //            foreach (var group in context.Web.ContentTypes.AsRequested())
        //            { }
        #endregion
        #region source group details logic
        //// Use earlier generated context to find context for all sites in the source tenant
        //var clonedContext = context.Clone(sourceURI);

        //// find groups from source tenants
        //sourceGroups = SiteConnection.GetGroups(clonedContext);

        //StringBuilder groupInformationString = new StringBuilder();

        //// Need to use Async here to avoid getting deadlocked
        //foreach (var list in await sourceGroups.ToListAsync())
        //{
        //    // groupInformationString.AppendLine($"Group Title: {list.Title}, Description: {list.Description} - {list.IsHiddenInUI}");
        //    groupInformationString.AppendLine($"Group Title: {list.Title}");
        //    groupNames.Add(list.Title);
        //}

        ////display groups info in textbox
        //this.txtSourceGroups.Text = groupInformationString.ToString();
        //TenantConfigMaster.sourceGroups = sourceGroups;
        #endregion
        #region find difference in groups
        //var result = sourceGroups.Except(targetGroupsComparision);

        //StringBuilder groupInformationString = new StringBuilder();

        //// Need to use Async here to avoid getting deadlocked
        //foreach (var list in result.ToList())
        //{
        //    // groupInformationString.AppendLine($"Group Title: {list.Title}, Description: {list.Description} - {list.IsHiddenInUI}");
        //    groupInformationString.AppendLine($"{list}");
        //}
        //MessageBox.Show("Missing source groups at Target are : " + groupInformationString.ToString());

        //return null;
        #endregion
    }
}
