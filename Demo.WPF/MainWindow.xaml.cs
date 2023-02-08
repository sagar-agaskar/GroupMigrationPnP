using PnP.Core.Model;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using GroupMigrationPnP.ConfigDetails;
using GroupMigrationPnP.HelperMethods;
using System.Collections.Generic;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;

namespace GroupMigrationPnP
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly IPnPContextFactory pnpContextFactory;              

        public MainWindow(IPnPContextFactory pnpFactory)
        {
            this.pnpContextFactory = pnpFactory;

            InitializeComponent();
        }

        internal async Task FindSourceGroupsAsync()
        {
            //convert textbox entry to URI
            Uri sourceURI = SiteConnection.GenerateURIFromString(txtSiteCollectionURL.Text);

            //find configuration to be used from appsettings.json
            string findConfigSiteValue = SiteConnection.GetSiteConfigurationDetails(sourceURI);

            //check if input entry site exists in source tenant defined in appsettings.json
            if (sourceURI != null && findConfigSiteValue != string.Empty)
            {
                
                //create client context based on config value found in  appsettings.json
                using (var context = await pnpContextFactory.CreateAsync(findConfigSiteValue))
                {                                           
                    var clonedContext = context.Clone(sourceURI);                    

                    TenantConfigMaster.sourceContext = clonedContext;                                        

                    MessageBox.Show("Successfully authenticated..");

                    GroupMigrationPnP.TargetWindow targetWindow = new TargetWindow(pnpContextFactory);
                    targetWindow.Show();
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("Please enter valid source URL");                
            }
        }

        private async void btnAuthenticate_Click(object sender, RoutedEventArgs e)
        {
            if (txtSiteCollectionURL.Text != "")
            {
                btnAuthenticate.IsEnabled = false;
                await FindSourceGroupsAsync();
                btnAuthenticate.IsEnabled = true;
            }
            else
            {
                MessageBox.Show("Please enter valid site URL");                
            }
        }      

        #region Find Site Title description Master page details
        //<Button x:Name="btnSite" Visibility="Hidden" Content="Site Information" HorizontalAlignment="Left" Height="34" Margin="23,55,0,0" VerticalAlignment="Top" Width="115" Click="btnSite_Click" Grid.Column="1" Grid.Row="0"/>
        //private async void btnSite_Click(object sender, RoutedEventArgs e)
        //{
        //await SiteInfoAsync();
        //}
        //internal async Task SiteInfoAsync()
        //{
        //    using (var context = await pnpContextFactory.CreateAsync("Q7Site"))
        //    {
        //        // Retrieving web with lists and masterpageurl loaded ==> SharePoint REST query
        //        var web = await context.Web.GetAsync(w => w.Title, w => w.Description, w => w.MasterUrl);
        //        this.txtResults.Text = $"Title: {web.Title} Description: {web.Description} MasterPageUrl: {web.MasterUrl}";
        //    }
        //}
        #endregion

        #region Find Teams in the site
        //private async void btnTeam_Click(object sender, RoutedEventArgs e)
        //{
        //await TeamInfoAsync();
        //}

        //internal async Task TeamInfoAsync()
        //{
        //    using (var context = await pnpContextFactory.CreateAsync("Q7Site"))
        //    {
        //        // Retrieving lists of the target web
        //        var team = await context.Team.GetAsync(t => t.DisplayName, t => t.Description, t => t.Channels);

        //        StringBuilder sb = new StringBuilder();
        //        sb.AppendLine($"Name: {team.DisplayName} Description: {team.Description}");
        //        sb.AppendLine();

        //        foreach (var channel in team.Channels)
        //        {
        //            sb.AppendLine($"Id: {channel.Id} Name: {channel.DisplayName}");
        //        }
        //        this.txtResults.Text = sb.ToString();
        //    }
        //}
        #endregion
    }
}
