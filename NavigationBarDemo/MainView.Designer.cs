namespace PDDL
{
    partial class MainView
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainView));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.menuItemSave = new System.Windows.Forms.ToolStripMenuItem();
            this.createPDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuItemPrint = new System.Windows.Forms.ToolStripMenuItem();
            this.menuItemTools = new System.Windows.Forms.ToolStripMenuItem();
            this.importHexFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuItemSetting = new System.Windows.Forms.ToolStripMenuItem();
            this.menuItemHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.panelNaviBar = new System.Windows.Forms.Panel();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.SerialPort = new System.IO.Ports.SerialPort(this.components);
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPageSelect = new System.Windows.Forms.TabPage();
            this.lblModNoSerNo = new System.Windows.Forms.Label();
            this.btnSearch = new System.Windows.Forms.Button();
            this.lblLoggerName = new System.Windows.Forms.Label();
            this.tabPageProgram = new System.Windows.Forms.TabPage();
            this.btnProgramLogger = new System.Windows.Forms.Button();
            this.grpBoxLoggerInfo = new System.Windows.Forms.GroupBox();
            this.pictureBox5 = new System.Windows.Forms.PictureBox();
            this.lblSerialNoValue = new System.Windows.Forms.Label();
            this.lblFirmwareValue = new System.Windows.Forms.Label();
            this.lblLoggerDateTimeValue = new System.Windows.Forms.Label();
            this.lblSerialNo = new System.Windows.Forms.Label();
            this.lblFirmware = new System.Windows.Forms.Label();
            this.lblLoggerDateTime = new System.Windows.Forms.Label();
            this.grpBoxAlarms = new System.Windows.Forms.GroupBox();
            this.lblLEDMsg = new System.Windows.Forms.Label();
            this.chkBoxDispOnTime = new System.Windows.Forms.CheckBox();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.chkBoxLED = new System.Windows.Forms.CheckBox();
            this.lblMinute3 = new System.Windows.Forms.Label();
            this.numUpDwnDispOnTime = new System.Windows.Forms.NumericUpDown();
            this.grpBoxAlarmSettings = new System.Windows.Forms.GroupBox();
            this.lblUpperAlram = new System.Windows.Forms.Label();
            this.lblLowerAlarm = new System.Windows.Forms.Label();
            this.lblRH2 = new System.Windows.Forms.Label();
            this.lblRH1 = new System.Windows.Forms.Label();
            this.numUpDwnHumiMax = new System.Windows.Forms.NumericUpDown();
            this.numUpDwnHumiMin = new System.Windows.Forms.NumericUpDown();
            this.lblHumidity = new System.Windows.Forms.Label();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.numUpDwnTempMax = new System.Windows.Forms.NumericUpDown();
            this.lblCelcius2 = new System.Windows.Forms.Label();
            this.lblCelcius1 = new System.Windows.Forms.Label();
            this.numUpDwnTempMin = new System.Windows.Forms.NumericUpDown();
            this.lblTemperature = new System.Windows.Forms.Label();
            this.grpBoxMeasurement = new System.Windows.Forms.GroupBox();
            this.lblRemark = new System.Windows.Forms.Label();
            this.comboBoxType = new System.Windows.Forms.ComboBox();
            this.lblMinute1 = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.lblMinute2 = new System.Windows.Forms.Label();
            this.numUpDwnstartdelay = new System.Windows.Forms.NumericUpDown();
            this.dateTimePickstop2 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickstop1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickstart2 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePickstart1 = new System.Windows.Forms.DateTimePicker();
            this.numUpDwninterval = new System.Windows.Forms.NumericUpDown();
            this.txtboxremark = new System.Windows.Forms.TextBox();
            this.lblStartDelay = new System.Windows.Forms.Label();
            this.lblStopTime = new System.Windows.Forms.Label();
            this.lblStartTime = new System.Windows.Forms.Label();
            this.lblInterval = new System.Windows.Forms.Label();
            this.lblType = new System.Windows.Forms.Label();
            this.lblDeviceName = new System.Windows.Forms.Label();
            this.tabPageReadData = new System.Windows.Forms.TabPage();
            this.btnReadData = new System.Windows.Forms.Button();
            this.tabPageShowDataChart = new System.Windows.Forms.TabPage();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.panelGraph = new System.Windows.Forms.Panel();
            this.zedGraphControl = new ZedGraph.ZedGraphControl();
            this.tabPageShowDataTab = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.tabPageAdminSett = new System.Windows.Forms.TabPage();
            this.grpBoxCompSett = new System.Windows.Forms.GroupBox();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtBoxCompLogo = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.txtBoxCompLoc = new System.Windows.Forms.TextBox();
            this.txtBoxCompName = new System.Windows.Forms.TextBox();
            this.lblCompLogo = new System.Windows.Forms.Label();
            this.lblCompLoc = new System.Windows.Forms.Label();
            this.lblCompName = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.naviBar1 = new Guifreaks.NavigationBar.NaviBar(this.components);
            this.naviBandLogger = new Guifreaks.NavigationBar.NaviBand(this.components);
            this.naviGroupLog = new Guifreaks.NavigationBar.NaviGroup(this.components);
            this.lblShowDataTab = new System.Windows.Forms.Label();
            this.lblShowDataChart = new System.Windows.Forms.Label();
            this.lblReadData = new System.Windows.Forms.Label();
            this.lblProgram = new System.Windows.Forms.Label();
            this.lblSelect = new System.Windows.Forms.Label();
            this.naviBandInfo = new Guifreaks.NavigationBar.NaviBand(this.components);
            this.naviGroupInfo = new Guifreaks.NavigationBar.NaviGroup(this.components);
            this.lblMaxHumiInfoValue = new System.Windows.Forms.Label();
            this.lblMinHumiInfoValue = new System.Windows.Forms.Label();
            this.lblInfoMaxHumi = new System.Windows.Forms.Label();
            this.lblInfoMinHumi = new System.Windows.Forms.Label();
            this.lblMaxTempInfoValue = new System.Windows.Forms.Label();
            this.lblMinTempInfoValue = new System.Windows.Forms.Label();
            this.lblToInfoValue = new System.Windows.Forms.Label();
            this.lblFromInfoValue = new System.Windows.Forms.Label();
            this.lblIntervalInfoValue = new System.Windows.Forms.Label();
            this.lblMeasurementsInfoValue = new System.Windows.Forms.Label();
            this.lblSerialNoInfoValue = new System.Windows.Forms.Label();
            this.lblInfoMaxTemp = new System.Windows.Forms.Label();
            this.lblInfoMinTemp = new System.Windows.Forms.Label();
            this.lblInfoTo = new System.Windows.Forms.Label();
            this.lblInfoFrom = new System.Windows.Forms.Label();
            this.lblInfoInterval = new System.Windows.Forms.Label();
            this.lblInfoMeasurements = new System.Windows.Forms.Label();
            this.lblInfoSerialNo = new System.Windows.Forms.Label();
            this.userManualToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.installationGuideToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutUsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.panelNaviBar.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.tabPageSelect.SuspendLayout();
            this.tabPageProgram.SuspendLayout();
            this.grpBoxLoggerInfo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).BeginInit();
            this.grpBoxAlarms.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnDispOnTime)).BeginInit();
            this.grpBoxAlarmSettings.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnHumiMax)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnHumiMin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnTempMax)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnTempMin)).BeginInit();
            this.grpBoxMeasurement.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnstartdelay)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwninterval)).BeginInit();
            this.tabPageReadData.SuspendLayout();
            this.tabPageShowDataChart.SuspendLayout();
            this.panelGraph.SuspendLayout();
            this.tabPageShowDataTab.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.tabPageAdminSett.SuspendLayout();
            this.grpBoxCompSett.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.naviBar1)).BeginInit();
            this.naviBar1.SuspendLayout();
            this.naviBandLogger.ClientArea.SuspendLayout();
            this.naviBandLogger.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.naviGroupLog)).BeginInit();
            this.naviGroupLog.SuspendLayout();
            this.naviBandInfo.ClientArea.SuspendLayout();
            this.naviBandInfo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.naviGroupInfo)).BeginInit();
            this.naviGroupInfo.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuItemSave,
            this.menuItemPrint,
            this.menuItemTools,
            this.menuItemSetting,
            this.menuItemHelp});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1199, 24);
            this.menuStrip1.TabIndex = 31;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // menuItemSave
            // 
            this.menuItemSave.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.createPDFToolStripMenuItem,
            this.exportExcelToolStripMenuItem});
            this.menuItemSave.Font = new System.Drawing.Font("Cambria", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuItemSave.Image = global::PDDL.Properties.Resources.save;
            this.menuItemSave.Name = "menuItemSave";
            this.menuItemSave.Size = new System.Drawing.Size(28, 20);
            this.menuItemSave.Click += new System.EventHandler(this.menuItemSave_Click);
            // 
            // createPDFToolStripMenuItem
            // 
            this.createPDFToolStripMenuItem.Image = global::PDDL.Properties.Resources.PDF_icon1;
            this.createPDFToolStripMenuItem.Name = "createPDFToolStripMenuItem";
            this.createPDFToolStripMenuItem.Size = new System.Drawing.Size(68, 22);
            this.createPDFToolStripMenuItem.Click += new System.EventHandler(this.createPDFToolStripMenuItem_Click);
            // 
            // exportExcelToolStripMenuItem
            // 
            this.exportExcelToolStripMenuItem.Image = global::PDDL.Properties.Resources.Excel_icon11;
            this.exportExcelToolStripMenuItem.Name = "exportExcelToolStripMenuItem";
            this.exportExcelToolStripMenuItem.Size = new System.Drawing.Size(68, 22);
            this.exportExcelToolStripMenuItem.Visible = false;
            this.exportExcelToolStripMenuItem.Click += new System.EventHandler(this.exportExcelToolStripMenuItem_Click);
            // 
            // menuItemPrint
            // 
            this.menuItemPrint.Font = new System.Drawing.Font("Cambria", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuItemPrint.Image = global::PDDL.Properties.Resources.print;
            this.menuItemPrint.Name = "menuItemPrint";
            this.menuItemPrint.Size = new System.Drawing.Size(28, 20);
            this.menuItemPrint.Click += new System.EventHandler(this.menuItemPrint_Click);
            // 
            // menuItemTools
            // 
            this.menuItemTools.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importHexFileToolStripMenuItem});
            this.menuItemTools.Font = new System.Drawing.Font("Cambria", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuItemTools.ForeColor = System.Drawing.SystemColors.ControlText;
            this.menuItemTools.Image = global::PDDL.Properties.Resources.Folder_Settings_Tools_icon;
            this.menuItemTools.Name = "menuItemTools";
            this.menuItemTools.Size = new System.Drawing.Size(28, 20);
            this.menuItemTools.Visible = false;
            // 
            // importHexFileToolStripMenuItem
            // 
            this.importHexFileToolStripMenuItem.Image = global::PDDL.Properties.Resources.Folder_Open;
            this.importHexFileToolStripMenuItem.Name = "importHexFileToolStripMenuItem";
            this.importHexFileToolStripMenuItem.Size = new System.Drawing.Size(68, 22);
            this.importHexFileToolStripMenuItem.Click += new System.EventHandler(this.importHexFileToolStripMenuItem_Click);
            // 
            // menuItemSetting
            // 
            this.menuItemSetting.Font = new System.Drawing.Font("Cambria", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuItemSetting.Image = global::PDDL.Properties.Resources.setting;
            this.menuItemSetting.Name = "menuItemSetting";
            this.menuItemSetting.Size = new System.Drawing.Size(28, 20);
            this.menuItemSetting.Click += new System.EventHandler(this.menuItemSetting_Click);
            // 
            // menuItemHelp
            // 
            this.menuItemHelp.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.userManualToolStripMenuItem,
            this.installationGuideToolStripMenuItem,
            this.aboutUsToolStripMenuItem});
            this.menuItemHelp.Font = new System.Drawing.Font("Cambria", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuItemHelp.Image = global::PDDL.Properties.Resources.help;
            this.menuItemHelp.Name = "menuItemHelp";
            this.menuItemHelp.Size = new System.Drawing.Size(28, 20);
            this.menuItemHelp.Click += new System.EventHandler(this.menuItemHelp_Click);
            // 
            // panelNaviBar
            // 
            this.panelNaviBar.Controls.Add(this.naviBar1);
            this.panelNaviBar.Location = new System.Drawing.Point(0, 30);
            this.panelNaviBar.Name = "panelNaviBar";
            this.panelNaviBar.Size = new System.Drawing.Size(353, 657);
            this.panelNaviBar.TabIndex = 33;
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.SelectedPath = "f";
            // 
            // tabControl
            // 
            this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl.Controls.Add(this.tabPageSelect);
            this.tabControl.Controls.Add(this.tabPageProgram);
            this.tabControl.Controls.Add(this.tabPageReadData);
            this.tabControl.Controls.Add(this.tabPageShowDataChart);
            this.tabControl.Controls.Add(this.tabPageShowDataTab);
            this.tabControl.Controls.Add(this.tabPageAdminSett);
            this.tabControl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl.Location = new System.Drawing.Point(355, 30);
            this.tabControl.Name = "tabControl";
            this.tabControl.Padding = new System.Drawing.Point(0, 0);
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(841, 678);
            this.tabControl.TabIndex = 34;
            // 
            // tabPageSelect
            // 
            this.tabPageSelect.BackgroundImage = global::PDDL.Properties.Resources.SoftLay3;
            this.tabPageSelect.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.tabPageSelect.Controls.Add(this.lblModNoSerNo);
            this.tabPageSelect.Controls.Add(this.btnSearch);
            this.tabPageSelect.Controls.Add(this.lblLoggerName);
            this.tabPageSelect.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageSelect.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tabPageSelect.Location = new System.Drawing.Point(4, 25);
            this.tabPageSelect.Margin = new System.Windows.Forms.Padding(0);
            this.tabPageSelect.Name = "tabPageSelect";
            this.tabPageSelect.Size = new System.Drawing.Size(833, 649);
            this.tabPageSelect.TabIndex = 0;
            this.tabPageSelect.UseVisualStyleBackColor = true;
            // 
            // lblModNoSerNo
            // 
            this.lblModNoSerNo.AutoSize = true;
            this.lblModNoSerNo.Font = new System.Drawing.Font("Cambria", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblModNoSerNo.ForeColor = System.Drawing.Color.Gold;
            this.lblModNoSerNo.Location = new System.Drawing.Point(350, 468);
            this.lblModNoSerNo.Name = "lblModNoSerNo";
            this.lblModNoSerNo.Size = new System.Drawing.Size(0, 23);
            this.lblModNoSerNo.TabIndex = 4;
            // 
            // btnSearch
            // 
            this.btnSearch.BackColor = System.Drawing.Color.Transparent;
            this.btnSearch.BackgroundImage = global::PDDL.Properties.Resources.buttonN;
            this.btnSearch.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSearch.FlatAppearance.BorderSize = 0;
            this.btnSearch.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSearch.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSearch.Location = new System.Drawing.Point(390, 530);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(140, 40);
            this.btnSearch.TabIndex = 3;
            this.btnSearch.UseVisualStyleBackColor = false;
            this.btnSearch.Click += new System.EventHandler(this.btnsearchlog_Click);
            // 
            // lblLoggerName
            // 
            this.lblLoggerName.AutoSize = true;
            this.lblLoggerName.Font = new System.Drawing.Font("Cambria", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLoggerName.ForeColor = System.Drawing.Color.Gold;
            this.lblLoggerName.Location = new System.Drawing.Point(300, 500);
            this.lblLoggerName.Name = "lblLoggerName";
            this.lblLoggerName.Size = new System.Drawing.Size(0, 23);
            this.lblLoggerName.TabIndex = 2;
            this.lblLoggerName.UseMnemonic = false;
            // 
            // tabPageProgram
            // 
            this.tabPageProgram.Controls.Add(this.btnProgramLogger);
            this.tabPageProgram.Controls.Add(this.grpBoxLoggerInfo);
            this.tabPageProgram.Controls.Add(this.grpBoxAlarms);
            this.tabPageProgram.Controls.Add(this.grpBoxAlarmSettings);
            this.tabPageProgram.Controls.Add(this.grpBoxMeasurement);
            this.tabPageProgram.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageProgram.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tabPageProgram.Location = new System.Drawing.Point(4, 25);
            this.tabPageProgram.Margin = new System.Windows.Forms.Padding(0);
            this.tabPageProgram.Name = "tabPageProgram";
            this.tabPageProgram.Size = new System.Drawing.Size(833, 649);
            this.tabPageProgram.TabIndex = 1;
            this.tabPageProgram.UseVisualStyleBackColor = true;
            // 
            // btnProgramLogger
            // 
            this.btnProgramLogger.BackColor = System.Drawing.Color.Transparent;
            this.btnProgramLogger.BackgroundImage = global::PDDL.Properties.Resources.Program_loggerN;
            this.btnProgramLogger.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnProgramLogger.FlatAppearance.BorderSize = 0;
            this.btnProgramLogger.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnProgramLogger.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProgramLogger.Location = new System.Drawing.Point(15, 590);
            this.btnProgramLogger.Name = "btnProgramLogger";
            this.btnProgramLogger.Size = new System.Drawing.Size(140, 40);
            this.btnProgramLogger.TabIndex = 5;
            this.btnProgramLogger.UseVisualStyleBackColor = false;
            this.btnProgramLogger.Click += new System.EventHandler(this.btnProgramLogger_Click);
            // 
            // grpBoxLoggerInfo
            // 
            this.grpBoxLoggerInfo.AutoSize = true;
            this.grpBoxLoggerInfo.Controls.Add(this.pictureBox5);
            this.grpBoxLoggerInfo.Controls.Add(this.lblSerialNoValue);
            this.grpBoxLoggerInfo.Controls.Add(this.lblFirmwareValue);
            this.grpBoxLoggerInfo.Controls.Add(this.lblLoggerDateTimeValue);
            this.grpBoxLoggerInfo.Controls.Add(this.lblSerialNo);
            this.grpBoxLoggerInfo.Controls.Add(this.lblFirmware);
            this.grpBoxLoggerInfo.Controls.Add(this.lblLoggerDateTime);
            this.grpBoxLoggerInfo.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpBoxLoggerInfo.Location = new System.Drawing.Point(15, 457);
            this.grpBoxLoggerInfo.Name = "grpBoxLoggerInfo";
            this.grpBoxLoggerInfo.Size = new System.Drawing.Size(700, 126);
            this.grpBoxLoggerInfo.TabIndex = 4;
            this.grpBoxLoggerInfo.TabStop = false;
            // 
            // pictureBox5
            // 
            this.pictureBox5.Image = global::PDDL.Properties.Resources.LoggIcon;
            this.pictureBox5.Location = new System.Drawing.Point(14, 33);
            this.pictureBox5.Name = "pictureBox5";
            this.pictureBox5.Size = new System.Drawing.Size(70, 60);
            this.pictureBox5.TabIndex = 9;
            this.pictureBox5.TabStop = false;
            // 
            // lblSerialNoValue
            // 
            this.lblSerialNoValue.AutoSize = true;
            this.lblSerialNoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSerialNoValue.Location = new System.Drawing.Point(260, 85);
            this.lblSerialNoValue.Name = "lblSerialNoValue";
            this.lblSerialNoValue.Size = new System.Drawing.Size(0, 19);
            this.lblSerialNoValue.TabIndex = 6;
            // 
            // lblFirmwareValue
            // 
            this.lblFirmwareValue.AutoSize = true;
            this.lblFirmwareValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFirmwareValue.Location = new System.Drawing.Point(260, 60);
            this.lblFirmwareValue.Name = "lblFirmwareValue";
            this.lblFirmwareValue.Size = new System.Drawing.Size(0, 19);
            this.lblFirmwareValue.TabIndex = 5;
            // 
            // lblLoggerDateTimeValue
            // 
            this.lblLoggerDateTimeValue.AutoSize = true;
            this.lblLoggerDateTimeValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLoggerDateTimeValue.Location = new System.Drawing.Point(260, 33);
            this.lblLoggerDateTimeValue.Name = "lblLoggerDateTimeValue";
            this.lblLoggerDateTimeValue.Size = new System.Drawing.Size(0, 19);
            this.lblLoggerDateTimeValue.TabIndex = 4;
            // 
            // lblSerialNo
            // 
            this.lblSerialNo.AutoSize = true;
            this.lblSerialNo.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSerialNo.Location = new System.Drawing.Point(100, 85);
            this.lblSerialNo.Name = "lblSerialNo";
            this.lblSerialNo.Size = new System.Drawing.Size(0, 19);
            this.lblSerialNo.TabIndex = 2;
            // 
            // lblFirmware
            // 
            this.lblFirmware.AutoSize = true;
            this.lblFirmware.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFirmware.Location = new System.Drawing.Point(100, 60);
            this.lblFirmware.Name = "lblFirmware";
            this.lblFirmware.Size = new System.Drawing.Size(0, 19);
            this.lblFirmware.TabIndex = 1;
            // 
            // lblLoggerDateTime
            // 
            this.lblLoggerDateTime.AutoSize = true;
            this.lblLoggerDateTime.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLoggerDateTime.Location = new System.Drawing.Point(100, 33);
            this.lblLoggerDateTime.Name = "lblLoggerDateTime";
            this.lblLoggerDateTime.Size = new System.Drawing.Size(0, 19);
            this.lblLoggerDateTime.TabIndex = 0;
            // 
            // grpBoxAlarms
            // 
            this.grpBoxAlarms.Controls.Add(this.lblLEDMsg);
            this.grpBoxAlarms.Controls.Add(this.chkBoxDispOnTime);
            this.grpBoxAlarms.Controls.Add(this.pictureBox4);
            this.grpBoxAlarms.Controls.Add(this.chkBoxLED);
            this.grpBoxAlarms.Controls.Add(this.lblMinute3);
            this.grpBoxAlarms.Controls.Add(this.numUpDwnDispOnTime);
            this.grpBoxAlarms.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpBoxAlarms.Location = new System.Drawing.Point(15, 355);
            this.grpBoxAlarms.Name = "grpBoxAlarms";
            this.grpBoxAlarms.Size = new System.Drawing.Size(700, 90);
            this.grpBoxAlarms.TabIndex = 3;
            this.grpBoxAlarms.TabStop = false;
            // 
            // lblLEDMsg
            // 
            this.lblLEDMsg.AutoSize = true;
            this.lblLEDMsg.ForeColor = System.Drawing.Color.Red;
            this.lblLEDMsg.Location = new System.Drawing.Point(260, 21);
            this.lblLEDMsg.Name = "lblLEDMsg";
            this.lblLEDMsg.Size = new System.Drawing.Size(0, 19);
            this.lblLEDMsg.TabIndex = 7;
            // 
            // chkBoxDispOnTime
            // 
            this.chkBoxDispOnTime.AutoSize = true;
            this.chkBoxDispOnTime.CausesValidation = false;
            this.chkBoxDispOnTime.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkBoxDispOnTime.Location = new System.Drawing.Point(100, 53);
            this.chkBoxDispOnTime.Name = "chkBoxDispOnTime";
            this.chkBoxDispOnTime.Size = new System.Drawing.Size(15, 14);
            this.chkBoxDispOnTime.TabIndex = 6;
            this.chkBoxDispOnTime.UseVisualStyleBackColor = true;
            this.chkBoxDispOnTime.CheckedChanged += new System.EventHandler(this.chkBoxDespTime_CheckedChanged);
            // 
            // pictureBox4
            // 
            this.pictureBox4.Image = global::PDDL.Properties.Resources.batSaveIcon;
            this.pictureBox4.Location = new System.Drawing.Point(14, 21);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(62, 50);
            this.pictureBox4.TabIndex = 5;
            this.pictureBox4.TabStop = false;
            // 
            // chkBoxLED
            // 
            this.chkBoxLED.AutoSize = true;
            this.chkBoxLED.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkBoxLED.Location = new System.Drawing.Point(100, 21);
            this.chkBoxLED.Name = "chkBoxLED";
            this.chkBoxLED.Size = new System.Drawing.Size(15, 14);
            this.chkBoxLED.TabIndex = 4;
            this.chkBoxLED.UseVisualStyleBackColor = true;
            this.chkBoxLED.CheckedChanged += new System.EventHandler(this.chkbxenableLED_CheckedChanged);
            // 
            // lblMinute3
            // 
            this.lblMinute3.AutoSize = true;
            this.lblMinute3.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMinute3.Location = new System.Drawing.Point(360, 57);
            this.lblMinute3.Name = "lblMinute3";
            this.lblMinute3.Size = new System.Drawing.Size(0, 19);
            this.lblMinute3.TabIndex = 3;
            // 
            // numUpDwnDispOnTime
            // 
            this.numUpDwnDispOnTime.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numUpDwnDispOnTime.Location = new System.Drawing.Point(260, 53);
            this.numUpDwnDispOnTime.Maximum = new decimal(new int[] {
            60,
            0,
            0,
            0});
            this.numUpDwnDispOnTime.Name = "numUpDwnDispOnTime";
            this.numUpDwnDispOnTime.Size = new System.Drawing.Size(99, 26);
            this.numUpDwnDispOnTime.TabIndex = 2;
            this.numUpDwnDispOnTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numUpDwnDispOnTime.ValueChanged += new System.EventHandler(this.numUpDwnDispOnTime_ValueChanged);
            this.numUpDwnDispOnTime.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.numUpDwnDispOnTime_KeyPress);
            // 
            // grpBoxAlarmSettings
            // 
            this.grpBoxAlarmSettings.Controls.Add(this.lblUpperAlram);
            this.grpBoxAlarmSettings.Controls.Add(this.lblLowerAlarm);
            this.grpBoxAlarmSettings.Controls.Add(this.lblRH2);
            this.grpBoxAlarmSettings.Controls.Add(this.lblRH1);
            this.grpBoxAlarmSettings.Controls.Add(this.numUpDwnHumiMax);
            this.grpBoxAlarmSettings.Controls.Add(this.numUpDwnHumiMin);
            this.grpBoxAlarmSettings.Controls.Add(this.lblHumidity);
            this.grpBoxAlarmSettings.Controls.Add(this.pictureBox3);
            this.grpBoxAlarmSettings.Controls.Add(this.numUpDwnTempMax);
            this.grpBoxAlarmSettings.Controls.Add(this.lblCelcius2);
            this.grpBoxAlarmSettings.Controls.Add(this.lblCelcius1);
            this.grpBoxAlarmSettings.Controls.Add(this.numUpDwnTempMin);
            this.grpBoxAlarmSettings.Controls.Add(this.lblTemperature);
            this.grpBoxAlarmSettings.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpBoxAlarmSettings.Location = new System.Drawing.Point(15, 245);
            this.grpBoxAlarmSettings.Name = "grpBoxAlarmSettings";
            this.grpBoxAlarmSettings.Size = new System.Drawing.Size(700, 100);
            this.grpBoxAlarmSettings.TabIndex = 2;
            this.grpBoxAlarmSettings.TabStop = false;
            // 
            // lblUpperAlram
            // 
            this.lblUpperAlram.AutoSize = true;
            this.lblUpperAlram.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblUpperAlram.Location = new System.Drawing.Point(475, 12);
            this.lblUpperAlram.Name = "lblUpperAlram";
            this.lblUpperAlram.Size = new System.Drawing.Size(0, 15);
            this.lblUpperAlram.TabIndex = 13;
            // 
            // lblLowerAlarm
            // 
            this.lblLowerAlarm.AutoSize = true;
            this.lblLowerAlarm.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLowerAlarm.Location = new System.Drawing.Point(270, 12);
            this.lblLowerAlarm.Name = "lblLowerAlarm";
            this.lblLowerAlarm.Size = new System.Drawing.Size(0, 15);
            this.lblLowerAlarm.TabIndex = 12;
            // 
            // lblRH2
            // 
            this.lblRH2.AutoSize = true;
            this.lblRH2.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRH2.Location = new System.Drawing.Point(570, 67);
            this.lblRH2.Name = "lblRH2";
            this.lblRH2.Size = new System.Drawing.Size(0, 19);
            this.lblRH2.TabIndex = 11;
            // 
            // lblRH1
            // 
            this.lblRH1.AutoSize = true;
            this.lblRH1.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRH1.Location = new System.Drawing.Point(365, 67);
            this.lblRH1.Name = "lblRH1";
            this.lblRH1.Size = new System.Drawing.Size(0, 19);
            this.lblRH1.TabIndex = 10;
            // 
            // numUpDwnHumiMax
            // 
            this.numUpDwnHumiMax.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numUpDwnHumiMax.Location = new System.Drawing.Point(465, 65);
            this.numUpDwnHumiMax.Maximum = new decimal(new int[] {
            99,
            0,
            0,
            0});
            this.numUpDwnHumiMax.Name = "numUpDwnHumiMax";
            this.numUpDwnHumiMax.Size = new System.Drawing.Size(100, 26);
            this.numUpDwnHumiMax.TabIndex = 9;
            this.numUpDwnHumiMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numUpDwnHumiMax.ValueChanged += new System.EventHandler(this.numUpDwnHumiMax_ValueChanged);
            this.numUpDwnHumiMax.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.numUpDwnHumiMax_KeyPress);
            // 
            // numUpDwnHumiMin
            // 
            this.numUpDwnHumiMin.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numUpDwnHumiMin.Location = new System.Drawing.Point(260, 65);
            this.numUpDwnHumiMin.Maximum = new decimal(new int[] {
            99,
            0,
            0,
            0});
            this.numUpDwnHumiMin.Name = "numUpDwnHumiMin";
            this.numUpDwnHumiMin.Size = new System.Drawing.Size(100, 26);
            this.numUpDwnHumiMin.TabIndex = 8;
            this.numUpDwnHumiMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numUpDwnHumiMin.ValueChanged += new System.EventHandler(this.numUpDwnHumiMin_ValueChanged);
            this.numUpDwnHumiMin.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.numUpDwnHumiMin_KeyPress);
            // 
            // lblHumidity
            // 
            this.lblHumidity.AutoSize = true;
            this.lblHumidity.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHumidity.Location = new System.Drawing.Point(100, 70);
            this.lblHumidity.Name = "lblHumidity";
            this.lblHumidity.Size = new System.Drawing.Size(0, 19);
            this.lblHumidity.TabIndex = 7;
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = global::PDDL.Properties.Resources.bellicon;
            this.pictureBox3.Location = new System.Drawing.Point(14, 21);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(62, 55);
            this.pictureBox3.TabIndex = 6;
            this.pictureBox3.TabStop = false;
            // 
            // numUpDwnTempMax
            // 
            this.numUpDwnTempMax.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numUpDwnTempMax.Location = new System.Drawing.Point(465, 33);
            this.numUpDwnTempMax.Maximum = new decimal(new int[] {
            70,
            0,
            0,
            0});
            this.numUpDwnTempMax.Minimum = new decimal(new int[] {
            30,
            0,
            0,
            -2147483648});
            this.numUpDwnTempMax.Name = "numUpDwnTempMax";
            this.numUpDwnTempMax.Size = new System.Drawing.Size(100, 26);
            this.numUpDwnTempMax.TabIndex = 5;
            this.numUpDwnTempMax.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numUpDwnTempMax.ValueChanged += new System.EventHandler(this.numUpDwnTempMax_ValueChanged);
            this.numUpDwnTempMax.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.numUpDwnTempMax_KeyPress);
            // 
            // lblCelcius2
            // 
            this.lblCelcius2.AutoSize = true;
            this.lblCelcius2.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCelcius2.Location = new System.Drawing.Point(570, 35);
            this.lblCelcius2.Name = "lblCelcius2";
            this.lblCelcius2.Size = new System.Drawing.Size(0, 15);
            this.lblCelcius2.TabIndex = 3;
            // 
            // lblCelcius1
            // 
            this.lblCelcius1.AutoSize = true;
            this.lblCelcius1.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCelcius1.Location = new System.Drawing.Point(365, 35);
            this.lblCelcius1.Name = "lblCelcius1";
            this.lblCelcius1.Size = new System.Drawing.Size(0, 15);
            this.lblCelcius1.TabIndex = 2;
            // 
            // numUpDwnTempMin
            // 
            this.numUpDwnTempMin.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numUpDwnTempMin.Location = new System.Drawing.Point(260, 33);
            this.numUpDwnTempMin.Maximum = new decimal(new int[] {
            70,
            0,
            0,
            0});
            this.numUpDwnTempMin.Minimum = new decimal(new int[] {
            30,
            0,
            0,
            -2147483648});
            this.numUpDwnTempMin.Name = "numUpDwnTempMin";
            this.numUpDwnTempMin.Size = new System.Drawing.Size(100, 26);
            this.numUpDwnTempMin.TabIndex = 1;
            this.numUpDwnTempMin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numUpDwnTempMin.ValueChanged += new System.EventHandler(this.numUpDwnTempMin_ValueChanged);
            this.numUpDwnTempMin.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.numUpDwnTempMin_KeyPress);
            // 
            // lblTemperature
            // 
            this.lblTemperature.AutoSize = true;
            this.lblTemperature.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTemperature.Location = new System.Drawing.Point(100, 35);
            this.lblTemperature.Name = "lblTemperature";
            this.lblTemperature.Size = new System.Drawing.Size(0, 19);
            this.lblTemperature.TabIndex = 0;
            // 
            // grpBoxMeasurement
            // 
            this.grpBoxMeasurement.Controls.Add(this.lblRemark);
            this.grpBoxMeasurement.Controls.Add(this.comboBoxType);
            this.grpBoxMeasurement.Controls.Add(this.lblMinute1);
            this.grpBoxMeasurement.Controls.Add(this.pictureBox2);
            this.grpBoxMeasurement.Controls.Add(this.lblMinute2);
            this.grpBoxMeasurement.Controls.Add(this.numUpDwnstartdelay);
            this.grpBoxMeasurement.Controls.Add(this.dateTimePickstop2);
            this.grpBoxMeasurement.Controls.Add(this.dateTimePickstop1);
            this.grpBoxMeasurement.Controls.Add(this.dateTimePickstart2);
            this.grpBoxMeasurement.Controls.Add(this.dateTimePickstart1);
            this.grpBoxMeasurement.Controls.Add(this.numUpDwninterval);
            this.grpBoxMeasurement.Controls.Add(this.txtboxremark);
            this.grpBoxMeasurement.Controls.Add(this.lblStartDelay);
            this.grpBoxMeasurement.Controls.Add(this.lblStopTime);
            this.grpBoxMeasurement.Controls.Add(this.lblStartTime);
            this.grpBoxMeasurement.Controls.Add(this.lblInterval);
            this.grpBoxMeasurement.Controls.Add(this.lblType);
            this.grpBoxMeasurement.Controls.Add(this.lblDeviceName);
            this.grpBoxMeasurement.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpBoxMeasurement.Location = new System.Drawing.Point(15, 15);
            this.grpBoxMeasurement.Name = "grpBoxMeasurement";
            this.grpBoxMeasurement.Size = new System.Drawing.Size(700, 220);
            this.grpBoxMeasurement.TabIndex = 1;
            this.grpBoxMeasurement.TabStop = false;
            // 
            // lblRemark
            // 
            this.lblRemark.AutoSize = true;
            this.lblRemark.ForeColor = System.Drawing.Color.Red;
            this.lblRemark.Location = new System.Drawing.Point(490, 34);
            this.lblRemark.Name = "lblRemark";
            this.lblRemark.Size = new System.Drawing.Size(184, 19);
            this.lblRemark.TabIndex = 19;
            this.lblRemark.Text = "Maximum 8 Characters ";
            // 
            // comboBoxType
            // 
            this.comboBoxType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxType.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxType.FormattingEnabled = true;
            this.comboBoxType.Items.AddRange(new object[] {
            "MEASURE UPON START TIME ",
            "START IMMEDIATELY UNTILL END OF MEMORY ",
            "START / STOP MEASUREMENT WITH TIME ",
            "START UPON KEY PRESS"});
            this.comboBoxType.Location = new System.Drawing.Point(260, 63);
            this.comboBoxType.Name = "comboBoxType";
            this.comboBoxType.Size = new System.Drawing.Size(320, 23);
            this.comboBoxType.TabIndex = 18;
            this.comboBoxType.SelectedIndexChanged += new System.EventHandler(this.comboBoxType_SelectedIndexChanged);
            // 
            // lblMinute1
            // 
            this.lblMinute1.AutoSize = true;
            this.lblMinute1.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMinute1.Location = new System.Drawing.Point(360, 96);
            this.lblMinute1.Name = "lblMinute1";
            this.lblMinute1.Size = new System.Drawing.Size(0, 19);
            this.lblMinute1.TabIndex = 17;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::PDDL.Properties.Resources.measureicon;
            this.pictureBox2.Location = new System.Drawing.Point(14, 34);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(52, 50);
            this.pictureBox2.TabIndex = 16;
            this.pictureBox2.TabStop = false;
            // 
            // lblMinute2
            // 
            this.lblMinute2.AutoSize = true;
            this.lblMinute2.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMinute2.Location = new System.Drawing.Point(360, 190);
            this.lblMinute2.Name = "lblMinute2";
            this.lblMinute2.Size = new System.Drawing.Size(0, 19);
            this.lblMinute2.TabIndex = 15;
            // 
            // numUpDwnstartdelay
            // 
            this.numUpDwnstartdelay.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numUpDwnstartdelay.Location = new System.Drawing.Point(260, 186);
            this.numUpDwnstartdelay.Maximum = new decimal(new int[] {
            65535,
            0,
            0,
            0});
            this.numUpDwnstartdelay.Name = "numUpDwnstartdelay";
            this.numUpDwnstartdelay.Size = new System.Drawing.Size(100, 26);
            this.numUpDwnstartdelay.TabIndex = 14;
            this.numUpDwnstartdelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numUpDwnstartdelay.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.numUpDwnstartdelay_KeyPress);
            // 
            // dateTimePickstop2
            // 
            this.dateTimePickstop2.CustomFormat = "HH:mm";
            this.dateTimePickstop2.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePickstop2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickstop2.Location = new System.Drawing.Point(450, 155);
            this.dateTimePickstop2.Name = "dateTimePickstop2";
            this.dateTimePickstop2.ShowUpDown = true;
            this.dateTimePickstop2.Size = new System.Drawing.Size(100, 23);
            this.dateTimePickstop2.TabIndex = 13;
            this.dateTimePickstop2.Value = new System.DateTime(2018, 9, 29, 11, 47, 0, 0);
            // 
            // dateTimePickstop1
            // 
            this.dateTimePickstop1.CustomFormat = "dd-MM-yyyy";
            this.dateTimePickstop1.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePickstop1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickstop1.Location = new System.Drawing.Point(260, 155);
            this.dateTimePickstop1.Name = "dateTimePickstop1";
            this.dateTimePickstop1.Size = new System.Drawing.Size(120, 23);
            this.dateTimePickstop1.TabIndex = 12;
            this.dateTimePickstop1.ValueChanged += new System.EventHandler(this.dateTimePickstop1_ValueChanged);
            // 
            // dateTimePickstart2
            // 
            this.dateTimePickstart2.CustomFormat = "HH:mm";
            this.dateTimePickstart2.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePickstart2.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickstart2.Location = new System.Drawing.Point(450, 125);
            this.dateTimePickstart2.Name = "dateTimePickstart2";
            this.dateTimePickstart2.ShowUpDown = true;
            this.dateTimePickstart2.Size = new System.Drawing.Size(100, 23);
            this.dateTimePickstart2.TabIndex = 11;
            this.dateTimePickstart2.Value = new System.DateTime(2018, 9, 29, 11, 47, 19, 0);
            // 
            // dateTimePickstart1
            // 
            this.dateTimePickstart1.CustomFormat = "dd-MM-yyyy";
            this.dateTimePickstart1.Font = new System.Drawing.Font("Cambria", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePickstart1.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dateTimePickstart1.Location = new System.Drawing.Point(260, 125);
            this.dateTimePickstart1.Name = "dateTimePickstart1";
            this.dateTimePickstart1.Size = new System.Drawing.Size(120, 23);
            this.dateTimePickstart1.TabIndex = 10;
            this.dateTimePickstart1.ValueChanged += new System.EventHandler(this.dateTimePickstart1_ValueChanged);
            // 
            // numUpDwninterval
            // 
            this.numUpDwninterval.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.numUpDwninterval.Location = new System.Drawing.Point(260, 94);
            this.numUpDwninterval.Maximum = new decimal(new int[] {
            255,
            0,
            0,
            0});
            this.numUpDwninterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numUpDwninterval.Name = "numUpDwninterval";
            this.numUpDwninterval.Size = new System.Drawing.Size(100, 26);
            this.numUpDwninterval.TabIndex = 8;
            this.numUpDwninterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numUpDwninterval.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numUpDwninterval.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.numUpDwninterval_KeyPress);
            // 
            // txtboxremark
            // 
            this.txtboxremark.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtboxremark.Location = new System.Drawing.Point(260, 31);
            this.txtboxremark.MaxLength = 8;
            this.txtboxremark.Name = "txtboxremark";
            this.txtboxremark.Size = new System.Drawing.Size(230, 26);
            this.txtboxremark.TabIndex = 6;
            this.txtboxremark.TextChanged += new System.EventHandler(this.txtboxremark_TextChanged);
            this.txtboxremark.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtboxremark_KeyPress);
            this.txtboxremark.Leave += new System.EventHandler(this.txtboxremark_Leave);
            // 
            // lblStartDelay
            // 
            this.lblStartDelay.AutoSize = true;
            this.lblStartDelay.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStartDelay.Location = new System.Drawing.Point(100, 188);
            this.lblStartDelay.Name = "lblStartDelay";
            this.lblStartDelay.Size = new System.Drawing.Size(0, 19);
            this.lblStartDelay.TabIndex = 5;
            // 
            // lblStopTime
            // 
            this.lblStopTime.AutoSize = true;
            this.lblStopTime.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStopTime.Location = new System.Drawing.Point(100, 158);
            this.lblStopTime.Name = "lblStopTime";
            this.lblStopTime.Size = new System.Drawing.Size(0, 19);
            this.lblStopTime.TabIndex = 4;
            // 
            // lblStartTime
            // 
            this.lblStartTime.AutoSize = true;
            this.lblStartTime.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStartTime.Location = new System.Drawing.Point(100, 130);
            this.lblStartTime.Name = "lblStartTime";
            this.lblStartTime.Size = new System.Drawing.Size(0, 19);
            this.lblStartTime.TabIndex = 3;
            // 
            // lblInterval
            // 
            this.lblInterval.AutoSize = true;
            this.lblInterval.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInterval.Location = new System.Drawing.Point(100, 96);
            this.lblInterval.Name = "lblInterval";
            this.lblInterval.Size = new System.Drawing.Size(0, 19);
            this.lblInterval.TabIndex = 2;
            // 
            // lblType
            // 
            this.lblType.AutoSize = true;
            this.lblType.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblType.Location = new System.Drawing.Point(100, 63);
            this.lblType.Name = "lblType";
            this.lblType.Size = new System.Drawing.Size(0, 19);
            this.lblType.TabIndex = 1;
            // 
            // lblDeviceName
            // 
            this.lblDeviceName.AutoSize = true;
            this.lblDeviceName.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDeviceName.Location = new System.Drawing.Point(100, 34);
            this.lblDeviceName.Name = "lblDeviceName";
            this.lblDeviceName.Size = new System.Drawing.Size(0, 19);
            this.lblDeviceName.TabIndex = 0;
            // 
            // tabPageReadData
            // 
            this.tabPageReadData.BackgroundImage = global::PDDL.Properties.Resources.SoftLay3;
            this.tabPageReadData.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.tabPageReadData.Controls.Add(this.btnReadData);
            this.tabPageReadData.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageReadData.Location = new System.Drawing.Point(4, 25);
            this.tabPageReadData.Margin = new System.Windows.Forms.Padding(0);
            this.tabPageReadData.Name = "tabPageReadData";
            this.tabPageReadData.Size = new System.Drawing.Size(833, 649);
            this.tabPageReadData.TabIndex = 2;
            this.tabPageReadData.UseVisualStyleBackColor = true;
            // 
            // btnReadData
            // 
            this.btnReadData.BackColor = System.Drawing.Color.Transparent;
            this.btnReadData.BackgroundImage = global::PDDL.Properties.Resources.Read_DataN;
            this.btnReadData.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnReadData.FlatAppearance.BorderSize = 0;
            this.btnReadData.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnReadData.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnReadData.Location = new System.Drawing.Point(390, 530);
            this.btnReadData.Name = "btnReadData";
            this.btnReadData.Size = new System.Drawing.Size(140, 40);
            this.btnReadData.TabIndex = 11;
            this.btnReadData.UseVisualStyleBackColor = false;
            this.btnReadData.Click += new System.EventHandler(this.btnReadData_Click);
            // 
            // tabPageShowDataChart
            // 
            this.tabPageShowDataChart.Controls.Add(this.btnRefresh);
            this.tabPageShowDataChart.Controls.Add(this.panelGraph);
            this.tabPageShowDataChart.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageShowDataChart.Location = new System.Drawing.Point(4, 25);
            this.tabPageShowDataChart.Name = "tabPageShowDataChart";
            this.tabPageShowDataChart.Size = new System.Drawing.Size(833, 649);
            this.tabPageShowDataChart.TabIndex = 4;
            this.tabPageShowDataChart.UseVisualStyleBackColor = true;
            // 
            // btnRefresh
            // 
            this.btnRefresh.BackColor = System.Drawing.Color.Transparent;
            this.btnRefresh.BackgroundImage = global::PDDL.Properties.Resources.Refresh_buttonN;
            this.btnRefresh.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnRefresh.FlatAppearance.BorderSize = 0;
            this.btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRefresh.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRefresh.Location = new System.Drawing.Point(310, 32);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(140, 40);
            this.btnRefresh.TabIndex = 4;
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // panelGraph
            // 
            this.panelGraph.Controls.Add(this.zedGraphControl);
            this.panelGraph.Location = new System.Drawing.Point(0, 100);
            this.panelGraph.Name = "panelGraph";
            this.panelGraph.Size = new System.Drawing.Size(984, 490);
            this.panelGraph.TabIndex = 0;
            // 
            // zedGraphControl
            // 
            this.zedGraphControl.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.zedGraphControl.Location = new System.Drawing.Point(5, 25);
            this.zedGraphControl.Margin = new System.Windows.Forms.Padding(4);
            this.zedGraphControl.Name = "zedGraphControl";
            this.zedGraphControl.ScrollGrace = 0D;
            this.zedGraphControl.ScrollMaxX = 0D;
            this.zedGraphControl.ScrollMaxY = 0D;
            this.zedGraphControl.ScrollMaxY2 = 0D;
            this.zedGraphControl.ScrollMinX = 0D;
            this.zedGraphControl.ScrollMinY = 0D;
            this.zedGraphControl.ScrollMinY2 = 0D;
            this.zedGraphControl.Size = new System.Drawing.Size(716, 440);
            this.zedGraphControl.TabIndex = 0;
            // 
            // tabPageShowDataTab
            // 
            this.tabPageShowDataTab.Controls.Add(this.panel2);
            this.tabPageShowDataTab.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageShowDataTab.Location = new System.Drawing.Point(4, 25);
            this.tabPageShowDataTab.Margin = new System.Windows.Forms.Padding(0);
            this.tabPageShowDataTab.Name = "tabPageShowDataTab";
            this.tabPageShowDataTab.Size = new System.Drawing.Size(833, 649);
            this.tabPageShowDataTab.TabIndex = 3;
            this.tabPageShowDataTab.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Margin = new System.Windows.Forms.Padding(0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(833, 649);
            this.panel2.TabIndex = 0;
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.BackgroundColor = System.Drawing.Color.White;
            this.dataGridView.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.dataGridView.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.LightBlue;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle1.NullValue = null;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.NullValue = null;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView.GridColor = System.Drawing.Color.Black;
            this.dataGridView.Location = new System.Drawing.Point(0, 0);
            this.dataGridView.Margin = new System.Windows.Forms.Padding(0);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.ReadOnly = true;
            this.dataGridView.RowHeadersVisible = false;
            this.dataGridView.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridViewCellStyle3.NullValue = null;
            this.dataGridView.RowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView.Size = new System.Drawing.Size(833, 649);
            this.dataGridView.TabIndex = 0;
            // 
            // tabPageAdminSett
            // 
            this.tabPageAdminSett.Controls.Add(this.grpBoxCompSett);
            this.tabPageAdminSett.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabPageAdminSett.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tabPageAdminSett.Location = new System.Drawing.Point(4, 25);
            this.tabPageAdminSett.Margin = new System.Windows.Forms.Padding(0);
            this.tabPageAdminSett.Name = "tabPageAdminSett";
            this.tabPageAdminSett.Size = new System.Drawing.Size(833, 649);
            this.tabPageAdminSett.TabIndex = 5;
            this.tabPageAdminSett.UseVisualStyleBackColor = true;
            // 
            // grpBoxCompSett
            // 
            this.grpBoxCompSett.Controls.Add(this.btnSubmit);
            this.grpBoxCompSett.Controls.Add(this.btnBrowse);
            this.grpBoxCompSett.Controls.Add(this.txtBoxCompLogo);
            this.grpBoxCompSett.Controls.Add(this.pictureBox1);
            this.grpBoxCompSett.Controls.Add(this.txtBoxCompLoc);
            this.grpBoxCompSett.Controls.Add(this.txtBoxCompName);
            this.grpBoxCompSett.Controls.Add(this.lblCompLogo);
            this.grpBoxCompSett.Controls.Add(this.lblCompLoc);
            this.grpBoxCompSett.Controls.Add(this.lblCompName);
            this.grpBoxCompSett.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpBoxCompSett.Location = new System.Drawing.Point(15, 15);
            this.grpBoxCompSett.Name = "grpBoxCompSett";
            this.grpBoxCompSett.Size = new System.Drawing.Size(814, 220);
            this.grpBoxCompSett.TabIndex = 0;
            this.grpBoxCompSett.TabStop = false;
            // 
            // btnSubmit
            // 
            this.btnSubmit.BackgroundImage = global::PDDL.Properties.Resources.SubmitN;
            this.btnSubmit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnSubmit.FlatAppearance.BorderSize = 0;
            this.btnSubmit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSubmit.Location = new System.Drawing.Point(260, 163);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(140, 40);
            this.btnSubmit.TabIndex = 8;
            this.btnSubmit.UseVisualStyleBackColor = true;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // btnBrowse
            // 
            this.btnBrowse.BackgroundImage = global::PDDL.Properties.Resources.BrowseN;
            this.btnBrowse.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnBrowse.FlatAppearance.BorderSize = 0;
            this.btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnBrowse.Location = new System.Drawing.Point(630, 107);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(140, 40);
            this.btnBrowse.TabIndex = 7;
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtBoxCompLogo
            // 
            this.txtBoxCompLogo.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxCompLogo.Location = new System.Drawing.Point(260, 116);
            this.txtBoxCompLogo.Name = "txtBoxCompLogo";
            this.txtBoxCompLogo.Size = new System.Drawing.Size(350, 26);
            this.txtBoxCompLogo.TabIndex = 6;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::PDDL.Properties.Resources.CompIcon;
            this.pictureBox1.Location = new System.Drawing.Point(14, 34);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(50, 50);
            this.pictureBox1.TabIndex = 5;
            this.pictureBox1.TabStop = false;
            // 
            // txtBoxCompLoc
            // 
            this.txtBoxCompLoc.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxCompLoc.Location = new System.Drawing.Point(260, 76);
            this.txtBoxCompLoc.MaxLength = 30;
            this.txtBoxCompLoc.Name = "txtBoxCompLoc";
            this.txtBoxCompLoc.Size = new System.Drawing.Size(250, 26);
            this.txtBoxCompLoc.TabIndex = 4;
            // 
            // txtBoxCompName
            // 
            this.txtBoxCompName.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBoxCompName.Location = new System.Drawing.Point(260, 36);
            this.txtBoxCompName.MaxLength = 30;
            this.txtBoxCompName.Name = "txtBoxCompName";
            this.txtBoxCompName.Size = new System.Drawing.Size(250, 26);
            this.txtBoxCompName.TabIndex = 3;
            // 
            // lblCompLogo
            // 
            this.lblCompLogo.AutoSize = true;
            this.lblCompLogo.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCompLogo.Location = new System.Drawing.Point(90, 120);
            this.lblCompLogo.Name = "lblCompLogo";
            this.lblCompLogo.Size = new System.Drawing.Size(0, 19);
            this.lblCompLogo.TabIndex = 2;
            // 
            // lblCompLoc
            // 
            this.lblCompLoc.AutoSize = true;
            this.lblCompLoc.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCompLoc.Location = new System.Drawing.Point(90, 80);
            this.lblCompLoc.Name = "lblCompLoc";
            this.lblCompLoc.Size = new System.Drawing.Size(0, 19);
            this.lblCompLoc.TabIndex = 1;
            // 
            // lblCompName
            // 
            this.lblCompName.AutoSize = true;
            this.lblCompName.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCompName.Location = new System.Drawing.Point(90, 40);
            this.lblCompName.Name = "lblCompName";
            this.lblCompName.Size = new System.Drawing.Size(0, 19);
            this.lblCompName.TabIndex = 0;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // naviBar1
            // 
            this.naviBar1.ActiveBand = this.naviBandLogger;
            this.naviBar1.Controls.Add(this.naviBandLogger);
            this.naviBar1.Controls.Add(this.naviBandInfo);
            this.naviBar1.Dock = System.Windows.Forms.DockStyle.Left;
            this.naviBar1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.naviBar1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.naviBar1.LayoutStyle = Guifreaks.NavigationBar.NaviLayoutStyle.Office2007Silver;
            this.naviBar1.Location = new System.Drawing.Point(0, 0);
            this.naviBar1.Name = "naviBar1";
            this.naviBar1.Size = new System.Drawing.Size(353, 657);
            this.naviBar1.TabIndex = 0;
            this.naviBar1.Text = "naviBar1";
            this.naviBar1.VisibleLargeButtons = 2;
            // 
            // naviBandLogger
            // 
            // 
            // naviBandLogger.ClientArea
            // 
            this.naviBandLogger.ClientArea.Controls.Add(this.naviGroupLog);
            this.naviBandLogger.ClientArea.Location = new System.Drawing.Point(0, 0);
            this.naviBandLogger.ClientArea.Name = "ClientArea";
            this.naviBandLogger.ClientArea.Size = new System.Drawing.Size(351, 526);
            this.naviBandLogger.ClientArea.TabIndex = 0;
            this.naviBandLogger.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.naviBandLogger.ForeColor = System.Drawing.SystemColors.ControlText;
            this.naviBandLogger.LargeImage = global::PDDL.Properties.Resources.LogIcon1;
            this.naviBandLogger.LayoutStyle = Guifreaks.NavigationBar.NaviLayoutStyle.Office2007Silver;
            this.naviBandLogger.Location = new System.Drawing.Point(1, 27);
            this.naviBandLogger.Name = "naviBandLogger";
            this.naviBandLogger.Size = new System.Drawing.Size(351, 526);
            this.naviBandLogger.SmallImage = global::PDDL.Properties.Resources.LogIcon1;
            this.naviBandLogger.TabIndex = 0;
            // 
            // naviGroupLog
            // 
            this.naviGroupLog.Caption = null;
            this.naviGroupLog.Controls.Add(this.lblShowDataTab);
            this.naviGroupLog.Controls.Add(this.lblShowDataChart);
            this.naviGroupLog.Controls.Add(this.lblReadData);
            this.naviGroupLog.Controls.Add(this.lblProgram);
            this.naviGroupLog.Controls.Add(this.lblSelect);
            this.naviGroupLog.Dock = System.Windows.Forms.DockStyle.Top;
            this.naviGroupLog.Expanded = false;
            this.naviGroupLog.ExpandedHeight = 540;
            this.naviGroupLog.HeaderContextMenuStrip = null;
            this.naviGroupLog.LayoutStyle = Guifreaks.NavigationBar.NaviLayoutStyle.Office2007Silver;
            this.naviGroupLog.Location = new System.Drawing.Point(0, 0);
            this.naviGroupLog.Name = "naviGroupLog";
            this.naviGroupLog.Padding = new System.Windows.Forms.Padding(1, 22, 1, 1);
            this.naviGroupLog.Size = new System.Drawing.Size(351, 540);
            this.naviGroupLog.TabIndex = 0;
            this.naviGroupLog.Text = "naviGroup1";
            // 
            // lblShowDataTab
            // 
            this.lblShowDataTab.BackColor = System.Drawing.SystemColors.Window;
            this.lblShowDataTab.Font = new System.Drawing.Font("Cambria", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblShowDataTab.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblShowDataTab.Image = global::PDDL.Properties.Resources.showtabdata;
            this.lblShowDataTab.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblShowDataTab.Location = new System.Drawing.Point(22, 145);
            this.lblShowDataTab.Name = "lblShowDataTab";
            this.lblShowDataTab.Size = new System.Drawing.Size(210, 30);
            this.lblShowDataTab.TabIndex = 9;
            this.lblShowDataTab.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblShowDataTab.Click += new System.EventHandler(this.lblshowdatatab_Click);
            // 
            // lblShowDataChart
            // 
            this.lblShowDataChart.BackColor = System.Drawing.SystemColors.Window;
            this.lblShowDataChart.Font = new System.Drawing.Font("Cambria", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblShowDataChart.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblShowDataChart.Image = global::PDDL.Properties.Resources.showchart;
            this.lblShowDataChart.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblShowDataChart.Location = new System.Drawing.Point(22, 185);
            this.lblShowDataChart.Name = "lblShowDataChart";
            this.lblShowDataChart.Size = new System.Drawing.Size(192, 30);
            this.lblShowDataChart.TabIndex = 8;
            this.lblShowDataChart.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblShowDataChart.Click += new System.EventHandler(this.lblshowdatachart_Click);
            // 
            // lblReadData
            // 
            this.lblReadData.BackColor = System.Drawing.SystemColors.Window;
            this.lblReadData.Font = new System.Drawing.Font("Cambria", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblReadData.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblReadData.Image = global::PDDL.Properties.Resources.readdata1;
            this.lblReadData.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblReadData.Location = new System.Drawing.Point(22, 110);
            this.lblReadData.Name = "lblReadData";
            this.lblReadData.Size = new System.Drawing.Size(127, 25);
            this.lblReadData.TabIndex = 7;
            this.lblReadData.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblReadData.Click += new System.EventHandler(this.lblreaddata_Click);
            // 
            // lblProgram
            // 
            this.lblProgram.BackColor = System.Drawing.SystemColors.Window;
            this.lblProgram.Font = new System.Drawing.Font("Cambria", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProgram.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblProgram.Image = global::PDDL.Properties.Resources.program;
            this.lblProgram.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblProgram.Location = new System.Drawing.Point(22, 70);
            this.lblProgram.Name = "lblProgram";
            this.lblProgram.Size = new System.Drawing.Size(119, 30);
            this.lblProgram.TabIndex = 6;
            this.lblProgram.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblProgram.Click += new System.EventHandler(this.lblprogram_Click);
            // 
            // lblSelect
            // 
            this.lblSelect.BackColor = System.Drawing.SystemColors.Window;
            this.lblSelect.Font = new System.Drawing.Font("Cambria", 12.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSelect.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblSelect.Image = global::PDDL.Properties.Resources.select;
            this.lblSelect.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblSelect.Location = new System.Drawing.Point(22, 38);
            this.lblSelect.Name = "lblSelect";
            this.lblSelect.Size = new System.Drawing.Size(95, 30);
            this.lblSelect.TabIndex = 5;
            this.lblSelect.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lblSelect.Click += new System.EventHandler(this.lblselect_Click);
            // 
            // naviBandInfo
            // 
            // 
            // naviBandInfo.ClientArea
            // 
            this.naviBandInfo.ClientArea.Controls.Add(this.naviGroupInfo);
            this.naviBandInfo.ClientArea.Location = new System.Drawing.Point(0, 0);
            this.naviBandInfo.ClientArea.Name = "ClientArea";
            this.naviBandInfo.ClientArea.Size = new System.Drawing.Size(351, 526);
            this.naviBandInfo.ClientArea.TabIndex = 0;
            this.naviBandInfo.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.naviBandInfo.LargeImage = global::PDDL.Properties.Resources.info;
            this.naviBandInfo.LayoutStyle = Guifreaks.NavigationBar.NaviLayoutStyle.Office2007Silver;
            this.naviBandInfo.Location = new System.Drawing.Point(1, 27);
            this.naviBandInfo.Name = "naviBandInfo";
            this.naviBandInfo.Size = new System.Drawing.Size(351, 526);
            this.naviBandInfo.SmallImage = global::PDDL.Properties.Resources.info;
            this.naviBandInfo.TabIndex = 6;
            // 
            // naviGroupInfo
            // 
            this.naviGroupInfo.Caption = null;
            this.naviGroupInfo.Controls.Add(this.lblMaxHumiInfoValue);
            this.naviGroupInfo.Controls.Add(this.lblMinHumiInfoValue);
            this.naviGroupInfo.Controls.Add(this.lblInfoMaxHumi);
            this.naviGroupInfo.Controls.Add(this.lblInfoMinHumi);
            this.naviGroupInfo.Controls.Add(this.lblMaxTempInfoValue);
            this.naviGroupInfo.Controls.Add(this.lblMinTempInfoValue);
            this.naviGroupInfo.Controls.Add(this.lblToInfoValue);
            this.naviGroupInfo.Controls.Add(this.lblFromInfoValue);
            this.naviGroupInfo.Controls.Add(this.lblIntervalInfoValue);
            this.naviGroupInfo.Controls.Add(this.lblMeasurementsInfoValue);
            this.naviGroupInfo.Controls.Add(this.lblSerialNoInfoValue);
            this.naviGroupInfo.Controls.Add(this.lblInfoMaxTemp);
            this.naviGroupInfo.Controls.Add(this.lblInfoMinTemp);
            this.naviGroupInfo.Controls.Add(this.lblInfoTo);
            this.naviGroupInfo.Controls.Add(this.lblInfoFrom);
            this.naviGroupInfo.Controls.Add(this.lblInfoInterval);
            this.naviGroupInfo.Controls.Add(this.lblInfoMeasurements);
            this.naviGroupInfo.Controls.Add(this.lblInfoSerialNo);
            this.naviGroupInfo.Dock = System.Windows.Forms.DockStyle.Top;
            this.naviGroupInfo.ExpandedHeight = 540;
            this.naviGroupInfo.HeaderContextMenuStrip = null;
            this.naviGroupInfo.LayoutStyle = Guifreaks.NavigationBar.NaviLayoutStyle.Office2007Silver;
            this.naviGroupInfo.Location = new System.Drawing.Point(0, 0);
            this.naviGroupInfo.Name = "naviGroupInfo";
            this.naviGroupInfo.Padding = new System.Windows.Forms.Padding(1, 22, 1, 1);
            this.naviGroupInfo.Size = new System.Drawing.Size(351, 540);
            this.naviGroupInfo.TabIndex = 0;
            this.naviGroupInfo.Text = "naviGroup1";
            // 
            // lblMaxHumiInfoValue
            // 
            this.lblMaxHumiInfoValue.AutoSize = true;
            this.lblMaxHumiInfoValue.BackColor = System.Drawing.SystemColors.Window;
            this.lblMaxHumiInfoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMaxHumiInfoValue.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblMaxHumiInfoValue.Location = new System.Drawing.Point(207, 242);
            this.lblMaxHumiInfoValue.Name = "lblMaxHumiInfoValue";
            this.lblMaxHumiInfoValue.Size = new System.Drawing.Size(0, 19);
            this.lblMaxHumiInfoValue.TabIndex = 61;
            // 
            // lblMinHumiInfoValue
            // 
            this.lblMinHumiInfoValue.AutoSize = true;
            this.lblMinHumiInfoValue.BackColor = System.Drawing.SystemColors.Window;
            this.lblMinHumiInfoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMinHumiInfoValue.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblMinHumiInfoValue.Location = new System.Drawing.Point(207, 215);
            this.lblMinHumiInfoValue.Name = "lblMinHumiInfoValue";
            this.lblMinHumiInfoValue.Size = new System.Drawing.Size(0, 19);
            this.lblMinHumiInfoValue.TabIndex = 60;
            // 
            // lblInfoMaxHumi
            // 
            this.lblInfoMaxHumi.AutoSize = true;
            this.lblInfoMaxHumi.BackColor = System.Drawing.SystemColors.Window;
            this.lblInfoMaxHumi.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfoMaxHumi.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblInfoMaxHumi.Location = new System.Drawing.Point(22, 243);
            this.lblInfoMaxHumi.Name = "lblInfoMaxHumi";
            this.lblInfoMaxHumi.Size = new System.Drawing.Size(0, 19);
            this.lblInfoMaxHumi.TabIndex = 59;
            // 
            // lblInfoMinHumi
            // 
            this.lblInfoMinHumi.AutoSize = true;
            this.lblInfoMinHumi.BackColor = System.Drawing.SystemColors.Window;
            this.lblInfoMinHumi.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfoMinHumi.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblInfoMinHumi.Location = new System.Drawing.Point(22, 215);
            this.lblInfoMinHumi.Name = "lblInfoMinHumi";
            this.lblInfoMinHumi.Size = new System.Drawing.Size(0, 19);
            this.lblInfoMinHumi.TabIndex = 58;
            // 
            // lblMaxTempInfoValue
            // 
            this.lblMaxTempInfoValue.AutoSize = true;
            this.lblMaxTempInfoValue.BackColor = System.Drawing.SystemColors.Window;
            this.lblMaxTempInfoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMaxTempInfoValue.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblMaxTempInfoValue.Location = new System.Drawing.Point(207, 190);
            this.lblMaxTempInfoValue.Name = "lblMaxTempInfoValue";
            this.lblMaxTempInfoValue.Size = new System.Drawing.Size(0, 19);
            this.lblMaxTempInfoValue.TabIndex = 56;
            // 
            // lblMinTempInfoValue
            // 
            this.lblMinTempInfoValue.AutoSize = true;
            this.lblMinTempInfoValue.BackColor = System.Drawing.SystemColors.Window;
            this.lblMinTempInfoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMinTempInfoValue.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblMinTempInfoValue.Location = new System.Drawing.Point(207, 166);
            this.lblMinTempInfoValue.Name = "lblMinTempInfoValue";
            this.lblMinTempInfoValue.Size = new System.Drawing.Size(0, 19);
            this.lblMinTempInfoValue.TabIndex = 55;
            // 
            // lblToInfoValue
            // 
            this.lblToInfoValue.AutoSize = true;
            this.lblToInfoValue.BackColor = System.Drawing.SystemColors.Window;
            this.lblToInfoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblToInfoValue.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblToInfoValue.Location = new System.Drawing.Point(207, 143);
            this.lblToInfoValue.Name = "lblToInfoValue";
            this.lblToInfoValue.Size = new System.Drawing.Size(0, 19);
            this.lblToInfoValue.TabIndex = 54;
            // 
            // lblFromInfoValue
            // 
            this.lblFromInfoValue.AutoSize = true;
            this.lblFromInfoValue.BackColor = System.Drawing.SystemColors.Window;
            this.lblFromInfoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFromInfoValue.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblFromInfoValue.Location = new System.Drawing.Point(207, 120);
            this.lblFromInfoValue.Name = "lblFromInfoValue";
            this.lblFromInfoValue.Size = new System.Drawing.Size(0, 19);
            this.lblFromInfoValue.TabIndex = 53;
            // 
            // lblIntervalInfoValue
            // 
            this.lblIntervalInfoValue.AutoSize = true;
            this.lblIntervalInfoValue.BackColor = System.Drawing.SystemColors.Window;
            this.lblIntervalInfoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblIntervalInfoValue.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblIntervalInfoValue.Location = new System.Drawing.Point(207, 98);
            this.lblIntervalInfoValue.Name = "lblIntervalInfoValue";
            this.lblIntervalInfoValue.Size = new System.Drawing.Size(0, 19);
            this.lblIntervalInfoValue.TabIndex = 52;
            // 
            // lblMeasurementsInfoValue
            // 
            this.lblMeasurementsInfoValue.AutoSize = true;
            this.lblMeasurementsInfoValue.BackColor = System.Drawing.SystemColors.Window;
            this.lblMeasurementsInfoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMeasurementsInfoValue.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblMeasurementsInfoValue.Location = new System.Drawing.Point(207, 74);
            this.lblMeasurementsInfoValue.Name = "lblMeasurementsInfoValue";
            this.lblMeasurementsInfoValue.Size = new System.Drawing.Size(0, 19);
            this.lblMeasurementsInfoValue.TabIndex = 51;
            // 
            // lblSerialNoInfoValue
            // 
            this.lblSerialNoInfoValue.AutoSize = true;
            this.lblSerialNoInfoValue.BackColor = System.Drawing.SystemColors.Window;
            this.lblSerialNoInfoValue.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSerialNoInfoValue.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblSerialNoInfoValue.Location = new System.Drawing.Point(207, 51);
            this.lblSerialNoInfoValue.Name = "lblSerialNoInfoValue";
            this.lblSerialNoInfoValue.Size = new System.Drawing.Size(0, 19);
            this.lblSerialNoInfoValue.TabIndex = 50;
            // 
            // lblInfoMaxTemp
            // 
            this.lblInfoMaxTemp.AutoSize = true;
            this.lblInfoMaxTemp.BackColor = System.Drawing.SystemColors.Window;
            this.lblInfoMaxTemp.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfoMaxTemp.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblInfoMaxTemp.Location = new System.Drawing.Point(22, 190);
            this.lblInfoMaxTemp.Name = "lblInfoMaxTemp";
            this.lblInfoMaxTemp.Size = new System.Drawing.Size(0, 19);
            this.lblInfoMaxTemp.TabIndex = 49;
            // 
            // lblInfoMinTemp
            // 
            this.lblInfoMinTemp.AutoSize = true;
            this.lblInfoMinTemp.BackColor = System.Drawing.SystemColors.Window;
            this.lblInfoMinTemp.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfoMinTemp.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblInfoMinTemp.Location = new System.Drawing.Point(22, 165);
            this.lblInfoMinTemp.Name = "lblInfoMinTemp";
            this.lblInfoMinTemp.Size = new System.Drawing.Size(0, 19);
            this.lblInfoMinTemp.TabIndex = 57;
            // 
            // lblInfoTo
            // 
            this.lblInfoTo.AutoSize = true;
            this.lblInfoTo.BackColor = System.Drawing.SystemColors.Window;
            this.lblInfoTo.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfoTo.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblInfoTo.Location = new System.Drawing.Point(22, 144);
            this.lblInfoTo.Name = "lblInfoTo";
            this.lblInfoTo.Size = new System.Drawing.Size(0, 19);
            this.lblInfoTo.TabIndex = 47;
            // 
            // lblInfoFrom
            // 
            this.lblInfoFrom.AutoSize = true;
            this.lblInfoFrom.BackColor = System.Drawing.SystemColors.Window;
            this.lblInfoFrom.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfoFrom.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblInfoFrom.Location = new System.Drawing.Point(22, 121);
            this.lblInfoFrom.Name = "lblInfoFrom";
            this.lblInfoFrom.Size = new System.Drawing.Size(0, 19);
            this.lblInfoFrom.TabIndex = 46;
            // 
            // lblInfoInterval
            // 
            this.lblInfoInterval.AutoSize = true;
            this.lblInfoInterval.BackColor = System.Drawing.SystemColors.Window;
            this.lblInfoInterval.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfoInterval.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblInfoInterval.Location = new System.Drawing.Point(22, 98);
            this.lblInfoInterval.Name = "lblInfoInterval";
            this.lblInfoInterval.Size = new System.Drawing.Size(0, 19);
            this.lblInfoInterval.TabIndex = 45;
            // 
            // lblInfoMeasurements
            // 
            this.lblInfoMeasurements.AutoSize = true;
            this.lblInfoMeasurements.BackColor = System.Drawing.SystemColors.Window;
            this.lblInfoMeasurements.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfoMeasurements.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblInfoMeasurements.Location = new System.Drawing.Point(22, 75);
            this.lblInfoMeasurements.Name = "lblInfoMeasurements";
            this.lblInfoMeasurements.Size = new System.Drawing.Size(0, 19);
            this.lblInfoMeasurements.TabIndex = 44;
            // 
            // lblInfoSerialNo
            // 
            this.lblInfoSerialNo.AutoSize = true;
            this.lblInfoSerialNo.BackColor = System.Drawing.SystemColors.Window;
            this.lblInfoSerialNo.Font = new System.Drawing.Font("Cambria", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInfoSerialNo.ForeColor = System.Drawing.SystemColors.Highlight;
            this.lblInfoSerialNo.Location = new System.Drawing.Point(22, 52);
            this.lblInfoSerialNo.Name = "lblInfoSerialNo";
            this.lblInfoSerialNo.Size = new System.Drawing.Size(0, 19);
            this.lblInfoSerialNo.TabIndex = 43;
            // 
            // userManualToolStripMenuItem
            // 
            this.userManualToolStripMenuItem.Name = "userManualToolStripMenuItem";
            this.userManualToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.userManualToolStripMenuItem.Text = "User Manual";
            this.userManualToolStripMenuItem.Click += new System.EventHandler(this.userManualToolStripMenuItem_Click);
            // 
            // installationGuideToolStripMenuItem
            // 
            this.installationGuideToolStripMenuItem.Name = "installationGuideToolStripMenuItem";
            this.installationGuideToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.installationGuideToolStripMenuItem.Text = "Installation Guide";
            this.installationGuideToolStripMenuItem.Click += new System.EventHandler(this.installationGuideToolStripMenuItem_Click);
            // 
            // aboutUsToolStripMenuItem
            // 
            this.aboutUsToolStripMenuItem.Name = "aboutUsToolStripMenuItem";
            this.aboutUsToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.aboutUsToolStripMenuItem.Text = "About Us";
            this.aboutUsToolStripMenuItem.Click += new System.EventHandler(this.aboutUsToolStripMenuItem_Click);
            // 
            // MainView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1199, 709);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.panelNaviBar);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainView";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "enviLOG Basic";
            this.Load += new System.EventHandler(this.MainView_Load);
            this.Shown += new System.EventHandler(this.MainView_Shown);
            this.SizeChanged += new System.EventHandler(this.MainView_SizeChanged);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.panelNaviBar.ResumeLayout(false);
            this.tabControl.ResumeLayout(false);
            this.tabPageSelect.ResumeLayout(false);
            this.tabPageSelect.PerformLayout();
            this.tabPageProgram.ResumeLayout(false);
            this.tabPageProgram.PerformLayout();
            this.grpBoxLoggerInfo.ResumeLayout(false);
            this.grpBoxLoggerInfo.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).EndInit();
            this.grpBoxAlarms.ResumeLayout(false);
            this.grpBoxAlarms.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnDispOnTime)).EndInit();
            this.grpBoxAlarmSettings.ResumeLayout(false);
            this.grpBoxAlarmSettings.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnHumiMax)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnHumiMin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnTempMax)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnTempMin)).EndInit();
            this.grpBoxMeasurement.ResumeLayout(false);
            this.grpBoxMeasurement.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwnstartdelay)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDwninterval)).EndInit();
            this.tabPageReadData.ResumeLayout(false);
            this.tabPageShowDataChart.ResumeLayout(false);
            this.panelGraph.ResumeLayout(false);
            this.tabPageShowDataTab.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.tabPageAdminSett.ResumeLayout(false);
            this.grpBoxCompSett.ResumeLayout(false);
            this.grpBoxCompSett.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.naviBar1)).EndInit();
            this.naviBar1.ResumeLayout(false);
            this.naviBandLogger.ClientArea.ResumeLayout(false);
            this.naviBandLogger.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.naviGroupLog)).EndInit();
            this.naviGroupLog.ResumeLayout(false);
            this.naviBandInfo.ClientArea.ResumeLayout(false);
            this.naviBandInfo.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.naviGroupInfo)).EndInit();
            this.naviGroupInfo.ResumeLayout(false);
            this.naviGroupInfo.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem menuItemSave;
        private System.Windows.Forms.ToolStripMenuItem menuItemPrint;
        private System.Windows.Forms.Label lblLoggerName;
        private System.Windows.Forms.Button btnSearch;
        private System.Windows.Forms.Panel panelNaviBar;
        private System.Windows.Forms.TabPage tabPageSelect;
        private System.Windows.Forms.TabPage tabPageProgram;
        private System.Windows.Forms.TabPage tabPageReadData;
        private Guifreaks.NavigationBar.NaviBar naviBar1;
        private Guifreaks.NavigationBar.NaviBand naviBandLogger;
        private Guifreaks.NavigationBar.NaviBand naviBandInfo;
        private Guifreaks.NavigationBar.NaviGroup naviGroupLog;
        private System.Windows.Forms.Label lblReadData;
        private System.Windows.Forms.Label lblProgram;
        private System.Windows.Forms.Label lblSelect;
        private Guifreaks.NavigationBar.NaviGroup naviGroupInfo;
        private System.Windows.Forms.Label lblMaxTempInfoValue;
        private System.Windows.Forms.Label lblMinTempInfoValue;
        private System.Windows.Forms.Label lblToInfoValue;
        private System.Windows.Forms.Label lblFromInfoValue;
        private System.Windows.Forms.Label lblIntervalInfoValue;
        private System.Windows.Forms.Label lblMeasurementsInfoValue;
        private System.Windows.Forms.Label lblSerialNoInfoValue;
        private System.Windows.Forms.Label lblInfoMaxTemp;
        private System.Windows.Forms.Label lblInfoMinTemp;
        private System.Windows.Forms.Label lblInfoTo;
        private System.Windows.Forms.Label lblInfoFrom;
        private System.Windows.Forms.Label lblInfoInterval;
        private System.Windows.Forms.Label lblInfoMeasurements;
        private System.Windows.Forms.Label lblInfoSerialNo;
        private System.Windows.Forms.GroupBox grpBoxMeasurement;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Label lblMinute2;
        private System.Windows.Forms.NumericUpDown numUpDwnstartdelay;
        private System.Windows.Forms.DateTimePicker dateTimePickstop2;
        private System.Windows.Forms.DateTimePicker dateTimePickstop1;
        private System.Windows.Forms.DateTimePicker dateTimePickstart2;
        private System.Windows.Forms.DateTimePicker dateTimePickstart1;
        private System.Windows.Forms.NumericUpDown numUpDwninterval;
        private System.Windows.Forms.TextBox txtboxremark;
        private System.Windows.Forms.Label lblStartDelay;
        private System.Windows.Forms.Label lblStopTime;
        private System.Windows.Forms.Label lblStartTime;
        private System.Windows.Forms.Label lblInterval;
        private System.Windows.Forms.Label lblType;
        private System.Windows.Forms.Label lblDeviceName;
        private System.Windows.Forms.GroupBox grpBoxAlarmSettings;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.NumericUpDown numUpDwnTempMax;
        private System.Windows.Forms.Label lblCelcius2;
        private System.Windows.Forms.Label lblCelcius1;
        private System.Windows.Forms.NumericUpDown numUpDwnTempMin;
        private System.Windows.Forms.Label lblTemperature;
        private System.Windows.Forms.GroupBox grpBoxAlarms;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.CheckBox chkBoxLED;
        private System.Windows.Forms.Label lblMinute3;
        private System.Windows.Forms.NumericUpDown numUpDwnDispOnTime;
        private System.Windows.Forms.GroupBox grpBoxLoggerInfo;
        private System.Windows.Forms.PictureBox pictureBox5;
        private System.Windows.Forms.Label lblSerialNoValue;
        private System.Windows.Forms.Label lblFirmwareValue;
        private System.Windows.Forms.Label lblLoggerDateTimeValue;
        private System.Windows.Forms.Label lblSerialNo;
        private System.Windows.Forms.Label lblFirmware;
        private System.Windows.Forms.Label lblLoggerDateTime;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.IO.Ports.SerialPort SerialPort;
        private System.Windows.Forms.Label lblMinute1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnReadData;
        private System.Windows.Forms.Label lblHumidity;
        private System.Windows.Forms.ComboBox comboBoxType;
        private System.Windows.Forms.Label lblLowerAlarm;
        private System.Windows.Forms.Label lblRH2;
        private System.Windows.Forms.Label lblRH1;
        private System.Windows.Forms.NumericUpDown numUpDwnHumiMax;
        private System.Windows.Forms.NumericUpDown numUpDwnHumiMin;
        private System.Windows.Forms.Label lblUpperAlram;
        private System.Windows.Forms.CheckBox chkBoxDispOnTime;
        private System.Windows.Forms.Label lblLEDMsg;
        private System.Windows.Forms.Label lblRemark;
        private System.Windows.Forms.Panel panelGraph;
        private ZedGraph.ZedGraphControl zedGraphControl;
        private System.Windows.Forms.Label lblMaxHumiInfoValue;
        private System.Windows.Forms.Label lblMinHumiInfoValue;
        private System.Windows.Forms.Label lblInfoMaxHumi;
        private System.Windows.Forms.Label lblInfoMinHumi;
        private System.Windows.Forms.ToolStripMenuItem menuItemHelp;
        private System.Windows.Forms.ToolStripMenuItem menuItemTools;
        private System.Windows.Forms.ToolStripMenuItem importHexFileToolStripMenuItem;
        private System.Windows.Forms.Button btnRefresh;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.TabPage tabPageAdminSett;
        private System.Windows.Forms.GroupBox grpBoxCompSett;
        private System.Windows.Forms.Label lblCompLogo;
        private System.Windows.Forms.Label lblCompLoc;
        private System.Windows.Forms.Label lblCompName;
        private System.Windows.Forms.TextBox txtBoxCompLoc;
        private System.Windows.Forms.TextBox txtBoxCompName;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox txtBoxCompLogo;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.Button btnProgramLogger;
        private System.Windows.Forms.Label lblModNoSerNo;
        private System.Windows.Forms.ToolStripMenuItem createPDFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exportExcelToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPageShowDataTab;
        private System.Windows.Forms.TabPage tabPageShowDataChart;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Label lblShowDataTab;
        private System.Windows.Forms.Label lblShowDataChart;
        private System.Windows.Forms.ToolStripMenuItem menuItemSetting;
        private System.Windows.Forms.ToolStripMenuItem userManualToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem installationGuideToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutUsToolStripMenuItem;
    }
}