using GroupMigrationPnP.ConfigDetails;
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
    /// Interaction logic for MigrationOptions.xaml
    /// </summary>
    public partial class MigrationOptions : Window
    {
        public MigrationOptions()
        {
            InitializeComponent();
        }                

        private void btnGetLists_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(TenantConfigMaster.sourceContext.Web.ToString() + " -- " + TenantConfigMaster.destContext.Web.ToString());
        }

        private void btnGetContentTypes_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnGetSiteColumns_Click(object sender, RoutedEventArgs e)
        {

        }

        private void btnGetGroups_Click(object sender, RoutedEventArgs e)
        {

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
