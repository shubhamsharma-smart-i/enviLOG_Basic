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
    public partial class Admin : Form
    {
        public Admin()
        {
            InitializeComponent();
        }

        private void Admin_Load(object sender, EventArgs e)
        {
            lblUserName.Text = ConstantVariables.UserNameLabel;
            lblPassword.Text = ConstantVariables.PasswordLabel;
            btnAdminOk.Text = ConstantVariables.AdminOkbutton;
        }

        private void txtboxUsername_TextChanged(object sender, EventArgs e)
        {
            Global.userNameValue = txtboxUsername.Text;
        }

        private void txtBoxPassword_TextChanged(object sender, EventArgs e)
        {
            Global.passwordValue = txtBoxPassword.Text;
        }

        private void btnAdminOk_Click(object sender, EventArgs e)
        {
            if (System.Configuration.ConfigurationManager.AppSettings["defaultUserName"] == Global.userNameValue && System.Configuration.ConfigurationManager.AppSettings["defaultPassword"] == Global.passwordValue)
            {              
                Global.correctValue = true;
                this.DialogResult = DialogResult.OK;
                txtboxUsername.Text = "";
                txtBoxPassword.Text = "";
                this.Close();
            }
            else
            {
                MessageBox.Show("Invalid Username and Password", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Admin_FormClosing(object sender, FormClosingEventArgs e)
        {
            txtboxUsername.Text = "";
            txtBoxPassword.Text = "";
        }

        public bool GetcorrectValue()
        {
            return Global.correctValue;
        }      

    }
}
