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


namespace PDDL
{
    public partial class frmloading : Form
    {
        
        int count = 0;
     
        public frmloading()
        {
            InitializeComponent();
        }

        private void frmloading_Load(object sender, EventArgs e)
        {
            this.Height = 60;
            lbl_loading_msg.Visible = true;
            lbl_loading_msg.Text = ConstantVariables.LoadingLabel;
            this.Location = new Point(650,550);
            timer1.Enabled = true;
            timer1.Tick += timer1_Tick;
            timer1.Interval = Global.timerInterval;
            timer1.Enabled = true;
            lblPercentage.Visible = false;
         }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (count < 100)
            {               
                count++;
                lblPercentage.Visible = true;
                lblPercentage.Text = count.ToString() + "%";
            }
            else
            {
                timer1.Enabled = false;
            }  
        }

        private void frmloading_FormClosing(object sender, FormClosingEventArgs e)
        {
            count = 0;
        }

        private void picBoxLoading_Click(object sender, EventArgs e)
        {

        }

    }
}
