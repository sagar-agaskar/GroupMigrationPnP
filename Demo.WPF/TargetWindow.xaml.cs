
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
        private readonly List<string> sourceGroups;
        private IQueryable<ISharePointGroup> targetGroups;
        List<string> targetGroupsComparision = new List<string>();

        public TargetWindow(IPnPContextFactory pnpFactory, List<string> sourceGroups)
        {
            this.pnpContextFactory = pnpFactory;
            this.sourceGroups = sourceGroups;
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
                    // Use earlier generated context to find context for all sites in the source tenant
                    var clonedContext = context.Clone(sourceURI);

                    // find groups from source tenants
                    targetGroups = SiteConnection.GetGroups(clonedContext);

                    StringBuilder groupInformationString = new StringBuilder();

                    // Need to use Async here to avoid getting deadlocked
                    foreach (var list in await targetGroups.ToListAsync())
                    {
                        // groupInformationString.AppendLine($"Group Title: {list.Title}, Description: {list.Description} - {list.IsHiddenInUI}");
                        groupInformationString.AppendLine($"Group Title: {list.Title}");
                        targetGroupsComparision.Add(list.Title);
                    }

                    //display groups info in textbox
                    this.txtTargetGroups.Text = groupInformationString.ToString();                    
                }
            }
            else
            {
                MessageBox.Show("Please enter valid source URL");
            }            
        }

        private async void btnGetTargetGroups_Click(object sender, RoutedEventArgs e)
        {
            btnTargetGroups.IsEnabled = false;
            await FindTargetGroupsAsync();
            btnTargetGroups.IsEnabled = true;
        }
        private async void btnTransferGroups_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await MergeGroups();
            }
            catch (Exception ex)
            {

            }            
        }

        private Task MergeGroups()
        {
            //throw new NotImplementedException();            
            var result = sourceGroups.Except(targetGroupsComparision);

            StringBuilder groupInformationString = new StringBuilder();

            // Need to use Async here to avoid getting deadlocked
            foreach (var list in result.ToList())
            {
                // groupInformationString.AppendLine($"Group Title: {list.Title}, Description: {list.Description} - {list.IsHiddenInUI}");
                groupInformationString.AppendLine($"{list}");
            }
            MessageBox.Show("Missing source groups at Target are : " + groupInformationString.ToString());

            return null;
        }
    }
}
