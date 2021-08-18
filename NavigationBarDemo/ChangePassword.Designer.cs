namespace PDDL
{
    partial class ChangePassword
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
            this.lblOldPassword = new System.Windows.Forms.Label();
            this.lblNewPassword = new System.Windows.Forms.Label();
            this.lblCnfrmPassword = new System.Windows.Forms.Label();
            this.txtBoxOldPwd = new System.Windows.Forms.TextBox();
            this.txtBoxNewPwd = new System.Windows.Forms.TextBox();
            this.txtBoxCnfrmPwd = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancelCP = new System.Windows.Forms.Button();
            this.chkBoxMasterPwd = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // lblOldPassword
            // 
            this.lblOldPassword.AutoSize = true;
            this.lblOldPassword.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblOldPassword.Location = new System.Drawing.Point(20, 40);
            this.lblOldPassword.Name = "lblOldPassword";
            this.lblOldPassword.Size = new System.Drawing.Size(0, 15);
            this.lblOldPassword.TabIndex = 0;
            // 
            // lblNewPassword
            // 
            this.lblNewPassword.AutoSize = true;
            this.lblNewPassword.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNewPassword.Location = new System.Drawing.Point(20, 80);
            this.lblNewPassword.Name = "lblNewPassword";
            this.lblNewPassword.Size = new System.Drawing.Size(0, 15);
            this.lblNewPassword.TabIndex = 1;
            // 
            // lblCnfrmPassword
            // 
            this.lblCnfrmPassword.AutoSize = true;
            this.lblCnfrmPassword.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCnfrmPassword.Location = new System.Drawing.Point(20, 120);
            this.lblCnfrmPassword.Name = "lblCnfrmPassword";
            this.lblCnfrmPassword.Size = new System.Drawing.Size(0, 15);
            this.lblCnfrmPassword.TabIndex = 2;
            // 
            // txtBoxOldPwd
            // 
            this.txtBoxOldPwd.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxOldPwd.Location = new System.Drawing.Point(180, 35);
            this.txtBoxOldPwd.Name = "txtBoxOldPwd";
            this.txtBoxOldPwd.PasswordChar = '*';
            this.txtBoxOldPwd.Size = new System.Drawing.Size(200, 23);
            this.txtBoxOldPwd.TabIndex = 3;
            // 
            // txtBoxNewPwd
            // 
            this.txtBoxNewPwd.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxNewPwd.Location = new System.Drawing.Point(180, 75);
            this.txtBoxNewPwd.Name = "txtBoxNewPwd";
            this.txtBoxNewPwd.PasswordChar = '*';
            this.txtBoxNewPwd.Size = new System.Drawing.Size(200, 23);
            this.txtBoxNewPwd.TabIndex = 4;
            // 
            // txtBoxCnfrmPwd
            // 
            this.txtBoxCnfrmPwd.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxCnfrmPwd.Location = new System.Drawing.Point(180, 115);
            this.txtBoxCnfrmPwd.Name = "txtBoxCnfrmPwd";
            this.txtBoxCnfrmPwd.PasswordChar = '*';
            this.txtBoxCnfrmPwd.Size = new System.Drawing.Size(200, 23);
            this.txtBoxCnfrmPwd.TabIndex = 5;
            this.txtBoxCnfrmPwd.TextChanged += new System.EventHandler(this.txtBoxCnfrmPwd_TextChanged);
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.Location = new System.Drawing.Point(180, 190);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(80, 30);
            this.btnSave.TabIndex = 6;
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancelCP
            // 
            this.btnCancelCP.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancelCP.Location = new System.Drawing.Point(300, 190);
            this.btnCancelCP.Name = "btnCancelCP";
            this.btnCancelCP.Size = new System.Drawing.Size(80, 30);
            this.btnCancelCP.TabIndex = 7;
            this.btnCancelCP.UseVisualStyleBackColor = true;
            this.btnCancelCP.Click += new System.EventHandler(this.btnCancelCP_Click);
            // 
            // chkBoxMasterPwd
            // 
            this.chkBoxMasterPwd.AutoSize = true;
            this.chkBoxMasterPwd.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkBoxMasterPwd.Location = new System.Drawing.Point(180, 160);
            this.chkBoxMasterPwd.Name = "chkBoxMasterPwd";
            this.chkBoxMasterPwd.Size = new System.Drawing.Size(15, 14);
            this.chkBoxMasterPwd.TabIndex = 8;
            this.chkBoxMasterPwd.UseVisualStyleBackColor = true;
            // 
            // ChangePassword
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(414, 232);
            this.Controls.Add(this.chkBoxMasterPwd);
            this.Controls.Add(this.btnCancelCP);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.txtBoxCnfrmPwd);
            this.Controls.Add(this.txtBoxNewPwd);
            this.Controls.Add(this.txtBoxOldPwd);
            this.Controls.Add(this.lblCnfrmPassword);
            this.Controls.Add(this.lblNewPassword);
            this.Controls.Add(this.lblOldPassword);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ChangePassword";
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Change  Password";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ChangePassword_FormClosing);
            this.Load += new System.EventHandler(this.ChangePassword_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblOldPassword;
        private System.Windows.Forms.Label lblNewPassword;
        private System.Windows.Forms.Label lblCnfrmPassword;
        private System.Windows.Forms.TextBox txtBoxOldPwd;
        private System.Windows.Forms.TextBox txtBoxNewPwd;
        private System.Windows.Forms.TextBox txtBoxCnfrmPwd;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancelCP;
        private System.Windows.Forms.CheckBox chkBoxMasterPwd;
    }
}