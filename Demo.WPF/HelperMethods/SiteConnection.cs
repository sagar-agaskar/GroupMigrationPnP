using GroupMigrationPnP;
using GroupMigrationPnP.ConfigDetails;
using GroupMigrationPnP.Entities;
using PnP.Core.Model;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

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
        public static void GetListDiscovery(string environment)
        {
            PnPContext context = null;

            if (environment.Equals("source"))
            {
                context = TenantConfigMaster.sourceContext;
            }
            else
            {
                context = TenantConfigMaster.destContext;
            }

            context.Web.Load(p=>p.Lists,p=>p.Lists.QueryProperties(
                p => p.Title,
                p => p.Description,
                p => p.DefaultDisplayFormUrl,
                p => p.ItemCount,
                p => p.TemplateType,
                p => p.Created,
                p => p.LastItemModifiedDate,
                p => p.Hidden));            
            
            string filename = String.Format("{0}_ListLibraryDiscovery_{1}.csv",
                                environment,
                                DateTime.UtcNow.ToString("HH-mm-ss"));
            String file = @"C:\MigrationRnD\" + filename;            

            String separator = ",";
            StringBuilder output = new StringBuilder();

            String[] headings = { "List/Library Name",
                "Description",
                "Url",
                "Items/Files Count", 
                "List Template",
                "Created Date",
                "Last Modified Date"};

            output.AppendLine(string.Join(separator, headings));

            foreach (var list in context.Web.Lists.AsRequested().Where(p=>p.Hidden==false))
            {
                string newLine = string.Format("{0}, {1}, {2}, {3}, {4},{5},{6}",
                    list.Title,
                    list.Description,  
                    list.DefaultDisplayFormUrl,
                    list.ItemCount.ToString(),
                    list.TemplateType,                    
                    list.Created.ToShortDateString(),
                    list.LastItemModifiedDate.ToShortDateString());

                output.AppendLine(string.Join(separator, newLine));
            }

            context.Web.Load(p => p.Webs);

            foreach (var subWeb in context.Web.Webs.AsRequested())
            {
                var clonedContext = context.Clone(subWeb.Url);
                output.AppendLine();

                clonedContext.Web.Load(p => p.Lists, p => p.Lists.QueryProperties(
                p => p.Title,
                p => p.Description,
                p => p.DefaultDisplayFormUrl,
                p => p.ItemCount,
                p => p.TemplateType,
                p => p.Created,
                p => p.LastItemModifiedDate,
                p => p.Hidden));

                foreach (var list in clonedContext.Web.Lists.AsRequested().Where(p => p.Hidden == false))
                {
                    string newLine = string.Format("{0}, {1}, {2}, {3}, {4},{5},{6}",
                        list.Title,
                        list.Description,
                        list.DefaultDisplayFormUrl,
                        list.ItemCount.ToString(),
                        list.TemplateType,
                        list.Created.ToShortDateString(),
                        list.LastItemModifiedDate.ToShortDateString());

                    output.AppendLine(string.Join(separator, newLine));
                }
            }

                try
            {
                File.AppendAllText(file, output.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("Data could not be written to the CSV file.");
                return;
            }
        }

        public static void GetSiteDiscovery(string environment)
        {
            PnPContext context = null;

            if (environment.Equals("source"))
            {
                context = TenantConfigMaster.sourceContext;
            }
            else
            {
                context = TenantConfigMaster.destContext;
            }

            var siteProperties = context.GetSiteCollectionManager().
                GetSiteCollectionProperties(new Uri(context.Web.Url.ToString()));
                      
            string filename = String.Format("{0}_SiteDiscovery_{1}.csv",
                environment,
                DateTime.UtcNow.ToString("HH-mm-ss"));

            String file = @"C:\MigrationRnD\" + filename;

            String separator = ",";
            StringBuilder output = new StringBuilder();
                        
                String[] headings = { "Site Name", "Sub Sites Count", "Last Modified Date", "Created Date","Owner", "Template Used", "Storage Use (MB)","URL" };
                output.AppendLine(string.Join(separator, headings));            

            string newLine = string.Format("{0}, {1}, {2}, {3}, {4},{5},{6},{7}",
                siteProperties.Title,
                Convert.ToString(siteProperties.WebsCount-1),
                siteProperties.LastContentModifiedDate.ToShortDateString(),
                "",//siteProperties.CreatedDate.ToShortDateString(),
                siteProperties.OwnerName,
                siteProperties.Template,
                Convert.ToString(siteProperties.StorageUsage),
                siteProperties.Url);
            output.AppendLine(string.Join(separator, newLine));

            context.Web.Load(p => p.Webs);

            foreach (var subWeb in context.Web.Webs.AsRequested())
            {
               var clonedContext = context.Clone(subWeb.Url);

               clonedContext.Web.Load(u =>u.Title,
                   u=>u.Url,                   
                   u => u.Created,
                   u => u.LastItemModifiedDate,
                   u =>u.WebTemplate);

                newLine = string.Format("{0}, {1}, {2}, {3}, {4},{5},{6},{7}",
                subWeb.Title,
                "",//subWeb.Created.ToShortDateString(),
                subWeb.LastItemModifiedDate.ToShortDateString(),
                subWeb.Created.ToShortDateString(),
                "",//subWeb.Author.Title,
                subWeb.WebTemplate,
                "",//storage used
                subWeb.Url.ToString()                
                );
                output.AppendLine(string.Join(separator, newLine));
            }

                try
            {
                File.AppendAllText(file, output.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("Data could not be written to the CSV file.");
                return;
            }
        }

        public static void AddGroups(string environment,ListBox listBox)
        {
            PnPContext context = null;
            if (environment.Equals("source"))
            {
                 context = TenantConfigMaster.sourceContext;
            }
            else
            {
                 context = TenantConfigMaster.destContext;
            }            

            context.Web.Load(p => p.SiteGroups.QueryProperties(u => u.Users, u => u.Title, u => u.OwnerTitle));
            if (context.Web.SiteGroups.Length > 0)
            {
                foreach (var group in context.Web.SiteGroups.AsRequested().ToList())
                {
                    listBox.Items.Add(group.Title);
                }
            }
        }

        public static GroupInfo GetSourceGroupDetails(string groupName,PnPContext ctx)
        {
           var siteGroup = ctx.Web.SiteGroups.QueryProperties(u => u.Users, u => u.Description, u => u.Title, u => u.OwnerTitle).FirstOrDefault(g => g.Title == groupName);
            //IRoleDefinitionCollection roles =siteGroup.GetRoleDefinitions();            
            
            GroupInfo srcGroupDetails = new GroupInfo(); //siteGroup.GetRoleDefinitions();
            srcGroupDetails.Name = groupName;
            srcGroupDetails.Owner = siteGroup.OwnerTitle;
            srcGroupDetails.Description = siteGroup.Description;
            //srcGroupDetails.Permission = siteGroup.GetRoleDefinitions();
            foreach (var user in siteGroup.Users)
            {
                GroupUser grpUser = new GroupUser();
                grpUser.UserName = user.Title;
                grpUser.UserEmail = user.UserPrincipalName;
                grpUser.LoginName = user.LoginName;

                srcGroupDetails.Users.Add(grpUser);
            }
            return srcGroupDetails;
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
