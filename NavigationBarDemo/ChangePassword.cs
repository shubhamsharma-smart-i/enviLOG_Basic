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
using PDDL.Interface;
using System.IO;
using System.Configuration;

namespace PDDL
{
    public partial class ChangePassword : Form
    {
        Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

        public ChangePassword()
        {
            InitializeComponent();
        }

        private void ChangePassword_Load(object sender, EventArgs e)
        {
            lblOldPassword.Text = ConstantVariables.OldPasswordLabel;
            lblNewPassword.Text = ConstantVariables.NewPasswordLabel;
            lblCnfrmPassword.Text = ConstantVariables.ConfirmPasswordLabel;
            chkBoxMasterPwd.Text = ConstantVariables.MasterPasswordChkBox;
            btnSave.Text = ConstantVariables.SaveButton;
            btnCancelCP.Text = ConstantVariables.CancelCPButton;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (chkBoxMasterPwd.Checked == false)
                {
                    if (!String.IsNullOrEmpty(txtBoxOldPwd.Text))     ////1
                    {
                        if (txtBoxOldPwd.Text == config.AppSettings.Settings["defaultPassword"].Value)    ////2
                        {
                            if (!String.IsNullOrEmpty(txtBoxNewPwd.Text))    ////3
                            {
                                if (txtBoxOldPwd.Text != txtBoxNewPwd.Text)    ////4
                                {
                                    if (!String.IsNullOrEmpty(txtBoxCnfrmPwd.Text))    ////5
                                    {
                                        if (txtBoxNewPwd.Text == txtBoxCnfrmPwd.Text)    ////6
                                        {
                                            config.AppSettings.Settings["defaultPassword"].Value = txtBoxNewPwd.Text;

                                            config.Save(ConfigurationSaveMode.Modified, true);
                                            ConfigurationManager.RefreshSection("appSettings");
                                            MessageBox.Show("Password Updated Successfully", "Save", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                           
                                            txtBoxOldPwd.Text = "";
                                            txtBoxNewPwd.Text = "";
                                            txtBoxCnfrmPwd.Text = "";
                                            this.Close();

                                        }    ////6
                                        else
                                        {
                                            MessageBox.Show("New Password and Confirm Password are not matching.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            txtBoxCnfrmPwd.Text = "";
                                        }

                                    }    ////5
                                    else
                                    {
                                        MessageBox.Show("Please enter Confirm Password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }

                                }    ////4
                                else
                                {
                                    MessageBox.Show("Please enter another password. Old password and new password are the same.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    txtBoxNewPwd.Text = "";
                                }

                            }    ////3
                            else
                            {
                                MessageBox.Show("Please enter New Password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }

                        }    ////2
                        else
                        {
                            MessageBox.Show("Old Password is not correct", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txtBoxOldPwd.Text = "";
                        }

                    }     ////1
                    else
                    {
                        MessageBox.Show("Please enter Old Password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                else      ////chkBoxMasterPwd.Checked == true
                {
                     if (!String.IsNullOrEmpty(txtBoxNewPwd.Text))    ////1
                     {
                        if (!String.IsNullOrEmpty(txtBoxCnfrmPwd.Text))    ////2
                        {
                            if (txtBoxNewPwd.Text == txtBoxCnfrmPwd.Text)    ////3
                            {
                                config.AppSettings.Settings["defaultPassword"].Value = txtBoxNewPwd.Text;
                               
                                config.Save(ConfigurationSaveMode.Modified, true);
                                ConfigurationManager.RefreshSection("appSettings");
                                MessageBox.Show("Password Updated Successfully", "Save", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                              
                                txtBoxOldPwd.Text = "";
                                txtBoxNewPwd.Text = "";
                                txtBoxCnfrmPwd.Text = "";
                                chkBoxMasterPwd.Checked = false;
                                this.Close();

                            }     ////3
                            else
                            {
                                MessageBox.Show("New Password and Confirm Password are not matching.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                txtBoxCnfrmPwd.Text = "";
                            }
                        }    ////2
                        else
                        {
                              MessageBox.Show("Please enter Confirm Password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                     }    ////1
                     else
                     {
                         MessageBox.Show("Please enter New Password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                     }
                 }
              
            }
            catch (Exception ex)
            {
                 FileLog.ErrorLog(ex.Message + ex.StackTrace); 
            }

        }

        private void btnCancelCP_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            txtBoxOldPwd.Text = "";
            txtBoxNewPwd.Text = "";
            txtBoxCnfrmPwd.Text = "";
            chkBoxMasterPwd.Checked = false;
            this.Close();
        }

        private void ChangePassword_FormClosing(object sender, FormClosingEventArgs e)
        {
            txtBoxOldPwd.Text = "";
            txtBoxNewPwd.Text = "";
            txtBoxCnfrmPwd.Text = "";
            chkBoxMasterPwd.Checked = false;
        }

        private void txtBoxCnfrmPwd_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
