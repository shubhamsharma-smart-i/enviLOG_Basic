namespace PDDL
{
    partial class LicenseKey
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LicenseKey));
            this.lblMACAddress = new System.Windows.Forms.Label();
            this.lblLicenseKey = new System.Windows.Forms.Label();
            this.txtBoxMACAdd = new System.Windows.Forms.TextBox();
            this.txtBoxLicenseKey = new System.Windows.Forms.TextBox();
            this.btnKeyOK = new System.Windows.Forms.Button();
            this.btnKeyCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblMACAddress
            // 
            this.lblMACAddress.AutoSize = true;
            this.lblMACAddress.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMACAddress.Location = new System.Drawing.Point(30, 30);
            this.lblMACAddress.Name = "lblMACAddress";
            this.lblMACAddress.Size = new System.Drawing.Size(0, 15);
            this.lblMACAddress.TabIndex = 0;
            // 
            // lblLicenseKey
            // 
            this.lblLicenseKey.AutoSize = true;
            this.lblLicenseKey.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLicenseKey.Location = new System.Drawing.Point(30, 70);
            this.lblLicenseKey.Name = "lblLicenseKey";
            this.lblLicenseKey.Size = new System.Drawing.Size(0, 15);
            this.lblLicenseKey.TabIndex = 1;
            // 
            // txtBoxMACAdd
            // 
            this.txtBoxMACAdd.Location = new System.Drawing.Point(150, 30);
            this.txtBoxMACAdd.Name = "txtBoxMACAdd";
            this.txtBoxMACAdd.ReadOnly = true;
            this.txtBoxMACAdd.Size = new System.Drawing.Size(220, 20);
            this.txtBoxMACAdd.TabIndex = 2;
            // 
            // txtBoxLicenseKey
            // 
            this.txtBoxLicenseKey.Location = new System.Drawing.Point(150, 70);
            this.txtBoxLicenseKey.Name = "txtBoxLicenseKey";
            this.txtBoxLicenseKey.Size = new System.Drawing.Size(220, 20);
            this.txtBoxLicenseKey.TabIndex = 3;
            // 
            // btnKeyOK
            // 
            this.btnKeyOK.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnKeyOK.Location = new System.Drawing.Point(150, 110);
            this.btnKeyOK.Name = "btnKeyOK";
            this.btnKeyOK.Size = new System.Drawing.Size(75, 30);
            this.btnKeyOK.TabIndex = 4;
            this.btnKeyOK.UseVisualStyleBackColor = true;
            this.btnKeyOK.Click += new System.EventHandler(this.btnKeyOK_Click);
            // 
            // btnKeyCancel
            // 
            this.btnKeyCancel.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnKeyCancel.Location = new System.Drawing.Point(270, 110);
            this.btnKeyCancel.Name = "btnKeyCancel";
            this.btnKeyCancel.Size = new System.Drawing.Size(75, 30);
            this.btnKeyCancel.TabIndex = 5;
            this.btnKeyCancel.UseVisualStyleBackColor = true;
            this.btnKeyCancel.Click += new System.EventHandler(this.btnKeyCancel_Click);
            // 
            // LicenseKey
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(402, 154);
            this.ControlBox = false;
            this.Controls.Add(this.btnKeyCancel);
            this.Controls.Add(this.btnKeyOK);
            this.Controls.Add(this.txtBoxLicenseKey);
            this.Controls.Add(this.txtBoxMACAdd);
            this.Controls.Add(this.lblLicenseKey);
            this.Controls.Add(this.lblMACAddress);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LicenseKey";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "LicenseKey";
            this.Load += new System.EventHandler(this.LicenseKey_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblMACAddress;
        private System.Windows.Forms.Label lblLicenseKey;
        private System.Windows.Forms.TextBox txtBoxMACAdd;
        private System.Windows.Forms.TextBox txtBoxLicenseKey;
        private System.Windows.Forms.Button btnKeyOK;
        private System.Windows.Forms.Button btnKeyCancel;
    }
}