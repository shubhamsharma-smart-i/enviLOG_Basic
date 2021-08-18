namespace PDDL
{
    partial class RadioBtn
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
            this.rdoBtnTab = new System.Windows.Forms.RadioButton();
            this.rdoBtnchart = new System.Windows.Forms.RadioButton();
            this.rdoBtnBoth = new System.Windows.Forms.RadioButton();
            this.btnRdoBtnOk = new System.Windows.Forms.Button();
            this.btnRdoBtnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // rdoBtnTab
            // 
            this.rdoBtnTab.AutoSize = true;
            this.rdoBtnTab.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoBtnTab.Location = new System.Drawing.Point(50, 25);
            this.rdoBtnTab.Name = "rdoBtnTab";
            this.rdoBtnTab.Size = new System.Drawing.Size(14, 13);
            this.rdoBtnTab.TabIndex = 0;
            this.rdoBtnTab.TabStop = true;
            this.rdoBtnTab.UseVisualStyleBackColor = true;
            this.rdoBtnTab.CheckedChanged += new System.EventHandler(this.rdoBtnTab_CheckedChanged);
            // 
            // rdoBtnchart
            // 
            this.rdoBtnchart.AutoSize = true;
            this.rdoBtnchart.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoBtnchart.Location = new System.Drawing.Point(50, 55);
            this.rdoBtnchart.Name = "rdoBtnchart";
            this.rdoBtnchart.Size = new System.Drawing.Size(14, 13);
            this.rdoBtnchart.TabIndex = 1;
            this.rdoBtnchart.TabStop = true;
            this.rdoBtnchart.UseVisualStyleBackColor = true;
            this.rdoBtnchart.CheckedChanged += new System.EventHandler(this.rdoBtnchart_CheckedChanged);
            // 
            // rdoBtnBoth
            // 
            this.rdoBtnBoth.AutoSize = true;
            this.rdoBtnBoth.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoBtnBoth.Location = new System.Drawing.Point(50, 85);
            this.rdoBtnBoth.Name = "rdoBtnBoth";
            this.rdoBtnBoth.Size = new System.Drawing.Size(14, 13);
            this.rdoBtnBoth.TabIndex = 2;
            this.rdoBtnBoth.TabStop = true;
            this.rdoBtnBoth.UseMnemonic = false;
            this.rdoBtnBoth.UseVisualStyleBackColor = true;
            this.rdoBtnBoth.CheckedChanged += new System.EventHandler(this.rdoBtnBoth_CheckedChanged);
            // 
            // btnRdoBtnOk
            // 
            this.btnRdoBtnOk.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRdoBtnOk.Location = new System.Drawing.Point(50, 125);
            this.btnRdoBtnOk.Name = "btnRdoBtnOk";
            this.btnRdoBtnOk.Size = new System.Drawing.Size(80, 30);
            this.btnRdoBtnOk.TabIndex = 3;
            this.btnRdoBtnOk.UseVisualStyleBackColor = true;
            this.btnRdoBtnOk.Click += new System.EventHandler(this.btnRdoBtnOk_Click);
            // 
            // btnRdoBtnCancel
            // 
            this.btnRdoBtnCancel.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRdoBtnCancel.Location = new System.Drawing.Point(170, 125);
            this.btnRdoBtnCancel.Name = "btnRdoBtnCancel";
            this.btnRdoBtnCancel.Size = new System.Drawing.Size(80, 30);
            this.btnRdoBtnCancel.TabIndex = 4;
            this.btnRdoBtnCancel.UseVisualStyleBackColor = true;
            this.btnRdoBtnCancel.Click += new System.EventHandler(this.btnRdoBtnCancel_Click);
            // 
            // RadioBtn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(284, 165);
            this.Controls.Add(this.btnRdoBtnCancel);
            this.Controls.Add(this.btnRdoBtnOk);
            this.Controls.Add(this.rdoBtnBoth);
            this.Controls.Add(this.rdoBtnchart);
            this.Controls.Add(this.rdoBtnTab);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "RadioBtn";
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Report";
            this.Load += new System.EventHandler(this.RadioBtn_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Button btnRdoBtnOk;
        public System.Windows.Forms.Button btnRdoBtnCancel;
        public System.Windows.Forms.RadioButton rdoBtnTab;
        public System.Windows.Forms.RadioButton rdoBtnchart;
        public System.Windows.Forms.RadioButton rdoBtnBoth;
    }
}