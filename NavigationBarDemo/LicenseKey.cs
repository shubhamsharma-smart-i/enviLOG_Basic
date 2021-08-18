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
using System.Security.Cryptography;
using System.Configuration;
using PDDL.Interface;

namespace PDDL
{
    public partial class LicenseKey : Form
    {
        Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
        
        public LicenseKey()
        {
            InitializeComponent();
        }

        private void LicenseKey_Load(object sender, EventArgs e)
        {
            lblMACAddress.Text = ConstantVariables.MACAddressLabel;
            lblLicenseKey.Text = ConstantVariables.LicenseKeyLabel;
            btnKeyOK.Text = ConstantVariables.OKKeyButton;
            btnKeyCancel.Text = ConstantVariables.CancelKeyButton;
            txtBoxMACAdd.Text = Global.sMacAddress;
        }

        private void btnKeyOK_Click(object sender, EventArgs e)
        {
            try
            {

                if (!String.IsNullOrEmpty(txtBoxLicenseKey.Text))
                {
                    try
                    {
                        Global.decyKeyMacAdd = Global.Decrypting(txtBoxLicenseKey.Text, "pwd@1234");
                    }
                    catch (Exception ex)
                    {
                        FileLog.ErrorLog(ex.Message + ex.StackTrace);
                    }
                    if (Global.sMacAddress == Global.decyKeyMacAdd)
                    {
                        config.AppSettings.Settings["MacAddress"].Value = txtBoxLicenseKey.Text;
                        config.Save(ConfigurationSaveMode.Modified, true);
                        ConfigurationManager.RefreshSection("appSettings");
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("License key is not correct", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Please enter License key", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace); 
            }
        }

        private void btnKeyCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

     }
}
