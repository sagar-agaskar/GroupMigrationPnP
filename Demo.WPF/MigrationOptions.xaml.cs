using GroupMigrationPnP.ConfigDetails;
using GroupMigrationPnP.Entities;
using GroupMigrationPnP.HelperMethods;
using Microsoft.ApplicationInsights.Extensibility.Implementation;
using PnP.Core;
using PnP.Core.Model;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace GroupMigrationPnP
{
    /// <summary>
    /// Interaction logic for MigrationOptions.xaml
    /// </summary>
    public partial class MigrationOptions : Window
    {
        public MigrationOptions()
        {
            InitializeComponent();
        }                

        private async void btnGetLists_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show(TenantConfigMaster.sourceContext.Web.ToString() + " -- " + TenantConfigMaster.destContext.Web.ToString());

            try
            {
                var srcContext = TenantConfigMaster.sourceContext;

                await srcContext.Web.Folders.LoadAsync(p=>p.Name);

                foreach (var srcList in srcContext.Web.Lists)
                {
                    lstSrcDetails.Items.Add(srcList.Title);
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnGetContentTypes_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnGetSiteColumns_Click(object sender, RoutedEventArgs e)
        {

        }
        
        private void btnGetGroups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SiteConnection.AddGroups("source",lstSrcDetails);
                SiteConnection.AddGroups("target",lstDestDetails);

                Button btnTransferGroups = new Button();
                btnTransferGroups.Content = "Add Missing Groups";
                btnTransferGroups.Click += new RoutedEventHandler(btnTransferGroups_Click);

                panelTarget.Children.Add(btnTransferGroups);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnTransferGroups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //MessageBox.Show("Hurray!!!");
                if (lstSrcDetails.SelectedItems.Count > 0)
                {
                    foreach (var item in lstSrcDetails.SelectedItems)
                    {
                        GroupInfo srcGroupInfo = SiteConnection.GetSourceGroupDetails(item.ToString(),TenantConfigMaster.sourceContext);
                        CreateGroupDestination(srcGroupInfo, TenantConfigMaster.destContext);
                    }
                }
                else
                {
                    MessageBox.Show("Please select atleast one source group");
                }
            }
            catch (Exception ex)
            {
                
            }
        }

        private void CreateGroupDestination(GroupInfo srcGroupInfo, PnPContext destContext)
        {
            try
            {
                destContext.Web.Load(p => p.SiteGroups.QueryProperties(u => u.Users, u => u.Title, u => u.OwnerTitle));

                var destGroup = destContext.Web.SiteGroups.FirstOrDefault(g => g.Title == srcGroupInfo.Name);

                if (destGroup == null)
                {

                    destGroup = destContext.Web.SiteGroups.Add(srcGroupInfo.Name);

                    if (srcGroupInfo.Description != null && srcGroupInfo.Description != "")
                        destGroup.Description = srcGroupInfo.Description;

                    destGroup.AddRoleDefinitions("Read");

                    var currentUser = destContext.Web.GetCurrentUser();

                    destGroup.SetUserAsOwner(currentUser.Id);

                    AddUsersInGroup(srcGroupInfo.Users, destContext, destGroup);
                }
                else
                {
                    MessageBox.Show("Group "+ srcGroupInfo.Name +"already exists at target site.");
                }
            }
            catch (SharePointRestServiceException ex)
            {
                MessageBox.Show(ex.Message + " - " + ex.InnerException + " - " + ex.HelpLink);
            }
        }

        private void AddUsersInGroup(List<GroupUser> users, PnPContext destContext, ISharePointGroup destGroup)
        {
            try
            {
                int i  = 0;
                string name = "";

                foreach (var user in users)
                {
                    ISharePointUser usertobeadded =null;                    
                    name = user.UserName;
                    i++;

                    try
                    {                        
                        var email = user.UserEmail.Replace("q7q0", "innovaorgdomain");
                        usertobeadded = destContext.Web.EnsureUser(email);
                        destGroup.Users.Add(usertobeadded.LoginName);
                        //i++;
                    }
                    catch (Exception ex)
                    {
                        GenerateCSVReport(i, name); //i++;
                    }                    
                }
                MessageBox.Show("Group" + destGroup.Title + " with users were added successfully.Please check Migration folder for missing users.");
            }
            catch (Exception ex)
            {                               
            }
        }

        public void GenerateCSVReport(int i, string name)
        {
            string path = @"C:\MigrationRnD\GroupReport.csv";

            // Set the variable "delimiter" to ", ".
            string delimiter = ", ";

            // This text is added only once to the file.
            if (!File.Exists(path))
            {
                // Create a file to write to.
                string createText = "Sr No." + delimiter + "User Name" + delimiter + "Remarks" + delimiter + Environment.NewLine;
                File.WriteAllText(path, createText);
            }

            // This text is always added, making the file longer over time
            // if it is not deleted.
            string appendText = i.ToString() + delimiter + name + delimiter + "Failed to add user since it does not exist at target" + delimiter + Environment.NewLine;
            File.AppendAllText(path, appendText);

            // Open the file to read from.
            string readText = File.ReadAllText(path);
            Console.WriteLine(readText);
        }

        private void btnGetPermissionLevels_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnGetMetadata_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnGetWorkflows_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
