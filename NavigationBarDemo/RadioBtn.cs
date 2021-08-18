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
    public partial class RadioBtn : Form
       
    {
        public RadioBtn()
        {
            InitializeComponent();
        }

        private void RadioBtn_Load(object sender, EventArgs e)
        {
            rdoBtnTab.Text = ConstantVariables.TabularRadiobutton;
            rdoBtnchart.Text = ConstantVariables.ChartRadiobutton;
            rdoBtnBoth.Text = ConstantVariables.BothRadiobutton;
            btnRdoBtnOk.Text = ConstantVariables.okRdobutton;
            btnRdoBtnCancel.Text = ConstantVariables.cancelRdobutton;
            rdoBtnBoth.Checked = true;
        }

        private void btnRdoBtnOk_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnRdoBtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void rdoBtnTab_CheckedChanged(object sender, EventArgs e)
        {
            Global.radioButtonValue = rdoBtnTab.Text;
        }

        private void rdoBtnchart_CheckedChanged(object sender, EventArgs e)
        {
            Global.radioButtonValue = rdoBtnchart.Text;
        }

        private void rdoBtnBoth_CheckedChanged(object sender, EventArgs e)
        {
            Global.radioButtonValue = rdoBtnBoth.Text;
        }

        public string GetrdbnValue()
        {
            return Global.radioButtonValue;
        }


    }
}
