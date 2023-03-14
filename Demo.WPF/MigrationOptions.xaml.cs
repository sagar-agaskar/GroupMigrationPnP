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
            try
            {
                lstSrcDetails.Items.Clear();
                lstDestDetails.Items.Clear();
                //panelTarget.Children.Clear();
                //panelSource.Children.Clear();
                btnTransferListData.Visibility = Visibility.Visible;
                btnTransferGroups.Visibility = Visibility.Hidden;

                TenantConfigMaster.sourceContext.Web.Load(p=>p.RoleDefinitions);
                TenantConfigMaster.sourceContext.Web.Load(p => p.Lists);
                TenantConfigMaster.destContext.Web.Load(p => p.Lists);

                foreach (var listName in TenantConfigMaster.sourceContext.Web.Lists.AsRequested().ToList())
                {
                    if (!listName.Hidden)
                    {
                        lstSrcDetails.Items.Add(listName.Title);
                    }
                }
                foreach (var listName in TenantConfigMaster.destContext.Web.Lists.AsRequested().ToList())
                {
                    if (!listName.Hidden)
                    {
                        lstDestDetails.Items.Add(listName.Title);
                    }                    
                }

                //Button btnTransferList = new Button();
                //btnTransferList.Name = "btnTransferList";
                //btnTransferList.Content = "Move list/library";
                //btnTransferList.Click += new RoutedEventHandler(btnTransferList_Click);
                
                //panelTarget.Children.Add(btnTransferList);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        
        private void btnSiteDetails_Click(object sender, RoutedEventArgs e)
        {
            lstSrcDetails.Items.Clear();
            lstDestDetails.Items.Clear();

            SiteConnection.GetSiteDiscovery("source");
            SiteConnection.GetSiteDiscovery("target");

            MessageBox.Show("Site details are captured and placed in Migration folder");
        }
        private void btnListsDetails_Click(object sender, RoutedEventArgs e)
        {
            lstSrcDetails.Items.Clear();
            lstDestDetails.Items.Clear();

            SiteConnection.GetListDiscovery("source");
            SiteConnection.GetListDiscovery("target");

            MessageBox.Show("List/Library details are captured and placed in Migration folder");
        }
            

        private void btnTransferList_Click(object sender, RoutedEventArgs e)
        {
            throw new NotImplementedException();
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
                lstSrcDetails.Items.Clear();
                lstDestDetails.Items.Clear();                

                SiteConnection.AddGroups("source",lstSrcDetails);
                SiteConnection.AddGroups("target",lstDestDetails);

                //Button btnTransferGroups = new Button();
                //btnTransferGroups.Name = "btnTransferGroups";
                //btnTransferGroups.Content = "Add Missing Groups";
                //btnTransferGroups.Click += new RoutedEventHandler(    );

                //panelTarget.Children.Add(btnTransferGroups);
                btnTransferGroups.Visibility = Visibility.Visible;
                btnTransferListData.Visibility = Visibility.Hidden;
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

        private void btnTransferListData_Click(object sender, RoutedEventArgs e)
        {
            if (lstSrcDetails.SelectedItems.Count>0)
            {
                foreach (var selectedList in lstSrcDetails.SelectedItems)
                {
                    //var ifListExis = TenantConfigMaster.sourceContext.Web.Lists.GetByTitle(selectedList.ToString());
                    var listName = selectedList.ToString();

                    TenantConfigMaster.sourceContext.Web.Load(p => p.Lists.QueryProperties(u => u.Title,u=>u.TemplateType));
                    var ifListExis = TenantConfigMaster.destContext.Web.Lists.First(g => g.Title == listName);
                    //ifListExis.Load(p=>p.TemplateType,p =>p.Title, p => p.Description, p => p.Fields);

                    if (ifListExis!=null)
                    {
                        MessageBox.Show("List"+ selectedList.ToString() +"already exists at target site");
                    }
                    else
                    {
                        var srcListDetails = TenantConfigMaster.sourceContext.Web.Lists.First(g => g.Title == listName);
                        srcListDetails.Load(p=>p.TemplateType,p =>p.Title, p => p.Description, p => p.Fields);

                        var myList = TenantConfigMaster.destContext.Web.Lists.Add(selectedList.ToString(), srcListDetails.TemplateType);
                        //ifListExis.Fields;
                        MessageBox.Show("List "+ myList .Title+ "created successfully at target site");
                    }
                }                
            }
            else
            {
                MessageBox.Show("Please select atleast one list to migrate.");
            }
        }
    }
}
