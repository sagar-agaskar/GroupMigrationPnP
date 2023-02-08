
using GroupMigrationPnP.ConfigDetails;
using GroupMigrationPnP.HelperMethods;
using PnP.Core.Model.Security;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System;
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
    /// Interaction logic for Target.xaml
    /// </summary>
    public partial class TargetWindow : Window
    {
        private readonly IPnPContextFactory pnpContextFactory;        

        public TargetWindow(IPnPContextFactory pnpFactory)
        {
            this.pnpContextFactory = pnpFactory;            
            InitializeComponent();
        }

        internal async Task FindTargetGroupsAsync()
        {
            //convert textbox entry to URI
            Uri sourceURI = SiteConnection.GenerateURIFromString(txtTargetSiteCollectionURL.Text);

            //find configuration to be used from appsettings.json
            string findConfigSiteValue = SiteConnection.GetSiteConfigurationDetails(sourceURI);

            //check if input entry site exists in source tenant defined in appsettings.json
            if (sourceURI != null && findConfigSiteValue != string.Empty)
            {
                //create client context based on config value found in  appsettings.json
                using (var context = await pnpContextFactory.CreateAsync(findConfigSiteValue))
                {                    
                    var clonedContext = context.Clone(sourceURI);                    

                    TenantConfigMaster.destContext = clonedContext;                    

                    MessageBox.Show("Successfully authenticated..");

                    MigrationOptions migrationOptions = new MigrationOptions();
                    migrationOptions.Show();
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("Please authenticate with valid source URL");               
            }
        }

        private async void btnAuthenticate_Click(object sender, RoutedEventArgs e)
        {
            if (txtTargetSiteCollectionURL.Text != "")
            {
                btnAuthenticate.IsEnabled = false;
                await FindTargetGroupsAsync();
                btnAuthenticate.IsEnabled = true;
            }
            else
            {
                MessageBox.Show("Please enter site URL");
            }
        }        
    }
}
