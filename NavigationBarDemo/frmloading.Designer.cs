namespace PDDL
{
    partial class frmloading
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.picBoxLoading = new System.Windows.Forms.PictureBox();
            this.lbl_loading_msg = new System.Windows.Forms.Label();
            this.lblPercentage = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.picBoxLoading)).BeginInit();
            this.SuspendLayout();
            // 
            // picBoxLoading
            // 
            this.picBoxLoading.BackColor = System.Drawing.Color.Transparent;
            this.picBoxLoading.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.picBoxLoading.Image = global::PDDL.Properties.Resources.ajax_loader;
            this.picBoxLoading.Location = new System.Drawing.Point(19, 32);
            this.picBoxLoading.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.picBoxLoading.Name = "picBoxLoading";
            this.picBoxLoading.Size = new System.Drawing.Size(427, 25);
            this.picBoxLoading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.picBoxLoading.TabIndex = 1;
            this.picBoxLoading.TabStop = false;
            this.picBoxLoading.Click += new System.EventHandler(this.picBoxLoading_Click);
            // 
            // lbl_loading_msg
            // 
            this.lbl_loading_msg.AutoSize = true;
            this.lbl_loading_msg.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_loading_msg.Location = new System.Drawing.Point(17, 9);
            this.lbl_loading_msg.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbl_loading_msg.Name = "lbl_loading_msg";
            this.lbl_loading_msg.Size = new System.Drawing.Size(0, 23);
            this.lbl_loading_msg.TabIndex = 2;
            this.lbl_loading_msg.Visible = false;
            // 
            // lblPercentage
            // 
            this.lblPercentage.AutoSize = true;
            this.lblPercentage.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPercentage.Location = new System.Drawing.Point(380, 9);
            this.lblPercentage.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblPercentage.Name = "lblPercentage";
            this.lblPercentage.Size = new System.Drawing.Size(0, 23);
            this.lblPercentage.TabIndex = 3;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // frmloading
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(459, 65);
            this.ControlBox = false;
            this.Controls.Add(this.lblPercentage);
            this.Controls.Add(this.lbl_loading_msg);
            this.Controls.Add(this.picBoxLoading);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmloading";
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmloading_FormClosing);
            this.Load += new System.EventHandler(this.frmloading_Load);
            ((System.ComponentModel.ISupportInitialize)(this.picBoxLoading)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox picBoxLoading;
        public System.Windows.Forms.Label lbl_loading_msg;
        private System.Windows.Forms.Label lblPercentage;
        private System.Windows.Forms.Timer timer1;
    }
}