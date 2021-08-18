using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PDDL.Common;

namespace PDDL
{
    public partial class AboutUs : Form
    {
        public AboutUs()
        {
            InitializeComponent();
        }

        private void AboutUs_Load(object sender, EventArgs e)
        {
            lblNameHelp.Text = ConstantVariables.EnviNameHelpLabel;
            lblsoftVersion.Text = ConstantVariables.SoftwareVersionLabel;      
            lblReleaseDate.Text = ConstantVariables.ReleaseDateLabel;
            lblAddress.Text = ConstantVariables.AddressLabel;
            lblTeleSupport.Text = ConstantVariables.TelephoneSupportLabel;
            lblForanyQuery.Text = ConstantVariables.ForAnyQueryLabel;
            lblNewEnquiries.Text = ConstantVariables.NewEnquiriesLabel;
            lblWebsite.Text = ConstantVariables.WebsiteLabel;
        
            lblsoftVersionValue.Text = ConstantVariables.SoftwareVersionValueLabel;
            lblReleaseDateValue.Text = ConstantVariables.ReleaseDateValueLabel;
            lblAddressValue.Text = ConstantVariables.AddressValueLabel;
            lblTeleSupportValue.Text = ConstantVariables.TelephoneSupportValueLabel;
            lblForanyQueryValue.Text = ConstantVariables.ForAnyQueryValueLabel;
            lblNewEnquiriesValue.Text = ConstantVariables.NewEnquiriesValueLabel;
            lblWebsiteValue.Text = ConstantVariables.WebsiteValueLabel;
            btnHelpOk.Text = ConstantVariables.HelpOKButton;            
        }

        private void btnHelpOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }       
      
    }
}
