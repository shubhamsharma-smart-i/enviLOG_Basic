using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Configuration;
using System.Diagnostics;
using System.Threading;

using ZedGraph;
using PDDL.Common;
using PDDL.Interface;
using iTextSharp.text;
using iTextSharp.text.pdf;

using System.IO;
using System.Web;

using System.Web.UI;
using System.Drawing;
using System.Management;
using System.Globalization;

using System.Drawing.Printing;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Net.NetworkInformation;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace PDDL
{
    public partial class MainView : Form
    {
        DeviceCommunication devComm = new DeviceCommunication();
        ChangePassword chgpass = new ChangePassword();
        frmloading load = new frmloading();
        LicenseKey lk = new LicenseKey();
        RadioBtn rdbn = new RadioBtn();
        Admin adm = new Admin();

        DataTable dt = new DataTable();
        Thread sf;
        Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
        List<System.Drawing.Color> graphTempColorList = new List<System.Drawing.Color>();
        List<System.Drawing.Color> graphHumiColorList = new List<System.Drawing.Color>();

        public MainView()
        {
            InitializeComponent();
            ConstantControls();
        }

        private void MainView_Load(object sender, EventArgs e)
        {

            Process proc = null;
            try
            {
                //string batDir = string.Format(@"C:\Users\shubham.s\Desktop\PDD_15-07-2021");
                string str = Environment.CurrentDirectory + "\\Sample.bat";
                System.Diagnostics.Process.Start(str);
                proc = new Process();
                proc.StartInfo.WorkingDirectory = str;
                proc.StartInfo.FileName = "Sample.bat";
                proc.StartInfo.CreateNoWindow = false;
                proc.Start();
                proc.WaitForExit();
                // MessageBox.Show("Bat file executed !!");
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.StackTrace.ToString());
            }
            try
            {
                File.Copy(Application.StartupPath + @"\enviro_logo.jpg", System.IO.Path.GetTempPath() + @"\enviLOG Basic\enviro_logo.jpg");
            }
            catch (Exception ex)
            {
            }
            try
            {
                if (IsProcessOpen("PDDL") == true)
                {
                    if (Global.CountValue > 1)
                    {
                        MessageBox.Show("Application is already running");
                        Application.Exit();
                    }
                    else
                    {
                        Global.CountValue--;
                    }
                }

                //try
                //{
                //    GetMACAddress();
                //    if (config.AppSettings.Settings["MacAddress"].Value != "")
                //    {
                //        Global.decyMacAdd = Global.Decrypting(config.AppSettings.Settings["MacAddress"].Value, "pwd@1234");
                //        if (Global.sMacAddress != Global.decyMacAdd)
                //        {
                //            DialogResult result = MessageBox.Show("Licensing Error.Please contact Enviro Technologies", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                //            if (result == DialogResult.OK)
                //            {
                //                DialogResult drk = lk.ShowDialog();
                //            }
                //            else
                //            {
                //                Application.Exit();
                //            }
                //        }
                //    }

                this.Location = new Point(0, 0);
                this.Size = Screen.PrimaryScreen.WorkingArea.Size;

                lblRemark.Visible = false;
                numUpDwnDispOnTime.Enabled = false;
                naviBandInfo.Visible = false;

                lblSelect.Enabled = false;

                if (Global.model == "PDL-K03")
                {
                    lblLoggerName.Location = new Point(300, 500);
                }
                if (Global.model == "PDL-K01")
                {
                    lblLoggerName.Location = new Point(330, 500);
                }

                tabControl.TabPages.Remove(tabPageProgram);
                tabControl.TabPages.Remove(tabPageAdminSett);
                tabControl.TabPages.Remove(tabPageReadData);
                tabControl.TabPages.Remove(tabPageShowDataChart);
                tabControl.TabPages.Remove(tabPageShowDataTab);
                Global.CurrentDateTime();


                if (Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["PreparedAndPrinted"]) == 0)
                {
                    PDFWithLogo.PreparedAndPrintedBy = "Prepared By ";
                }
                else if (Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["PreparedAndPrinted"]) == 1)
                {
                    PDFWithLogo.PreparedAndPrintedBy = "Printed By ";
                }


                if (Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["defaultCreateExcel"]) == 0)
                {
                    exportExcelToolStripMenuItem.Visible = true;
                }
                else if (Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["defaultCreateExcel"]) == 1)
                {
                    exportExcelToolStripMenuItem.Visible = false;
                }

                try
                {
                    if (Global.responseStringW == "" || Global.responseStringW == null)
                    {
                        lblLoggerName.Text = "";
                    }
                    if (Global.responseDevName == "PDL-K03")
                    {
                        lblLoggerName.Text = ConstantVariables.LoggerNameValueTR + "    /    " + Global.chkSerialNoStr;
                    }
                    else if (Global.responseDevName == "PDL-K01")
                    {
                        lblLoggerName.Text = ConstantVariables.LoggerNameValueT + "    /    " + Global.chkSerialNoStr;
                    }
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);//(Shubham sharma)
            }
        }

        /// <summary>
        /// Method for checking application is running in task manager 
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public bool IsProcessOpen(string name)
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.Contains("NavigationBarDemo"))
                {
                    Global.CountValue++;
                }
            }
            return true;
        }

        /// <summary>
        /// Getting MAC address of the system
        /// </summary>
        /// <returns></returns>
        //public string GetMACAddress()
        //{
        //    NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
        //    foreach (NetworkInterface adapter in nics)
        //    {
        //        if (Global.sMacAddress == String.Empty)// only return MAC Address from first card  
        //        {
        //            IPInterfaceProperties properties = adapter.GetIPProperties();
        //            Global.sMacAddress = adapter.GetPhysicalAddress().ToString();
        //        }
        //    }
        //    return Global.sMacAddress;
        //}

        private void MainView_Shown(object sender, EventArgs e)
        {
            if (Global.responseStringW == "")
            {
                MessageBox.Show("No any device is connected", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void MainView_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                foreach (Form frm in Application.OpenForms)
                {
                    frm.WindowState = FormWindowState.Minimized;
                }
            }
            else if (this.WindowState == FormWindowState.Normal)
            {
                foreach (Form frm in Application.OpenForms)
                {
                    frm.WindowState = FormWindowState.Normal;
                }
            }
        }

        /// <summary>
        /// All controls text
        /// </summary>
        public void ConstantControls()
        {
            lblSelect.Text = ConstantVariables.SelectLabel;
            lblProgram.Text = ConstantVariables.ProgramLabel;
            lblReadData.Text = ConstantVariables.ReadDataLabel;
            lblShowDataChart.Text = ConstantVariables.ShowDataChartLabel;
            lblShowDataTab.Text = ConstantVariables.ShowDataTabLabel;

            menuItemSave.Text = ConstantVariables.SaveMenuItem;
            createPDFToolStripMenuItem.Text = ConstantVariables.CreatePDFMenuItem;
            exportExcelToolStripMenuItem.Text = ConstantVariables.ExportExcelMenuItem;
            menuItemPrint.Text = ConstantVariables.PrintMenuItem;
            menuItemTools.Text = ConstantVariables.ToolsMenuItem;
            importHexFileToolStripMenuItem.Text = ConstantVariables.ImportHexFileMenuItem;
            menuItemSetting.Text = ConstantVariables.SettingMenuItem;
            menuItemHelp.Text = ConstantVariables.HelpMenuItem;

            tabPageSelect.Text = ConstantVariables.SelectTabPage;
            tabPageProgram.Text = ConstantVariables.ProgramTabPage;
            tabPageAdminSett.Text = ConstantVariables.AdminSettingTabPage;
            tabPageReadData.Text = ConstantVariables.ReadDataTabPage;
            tabPageShowDataChart.Text = ConstantVariables.ShowDataChartTabPage;
            tabPageShowDataTab.Text = ConstantVariables.ShowDataTabTabPage;

            naviBandLogger.Text = ConstantVariables.LoggerNaviBand;
            naviBandInfo.Text = ConstantVariables.InfoNaviBand;

            lblLoggerName.Text = ConstantVariables.LoggerNameLabel;
            lblModNoSerNo.Text = ConstantVariables.ModNoSerNoLabel;

            grpBoxMeasurement.Text = ConstantVariables.MeasureGroupBox;
            lblDeviceName.Text = ConstantVariables.DeviceNameLabel;
            lblType.Text = ConstantVariables.TypeLabel;
            lblInterval.Text = ConstantVariables.IntervalLabel;
            lblStartTime.Text = ConstantVariables.StartTimeLabel;
            lblStopTime.Text = ConstantVariables.StopTimeLabel;
            lblStartDelay.Text = ConstantVariables.StartDelayLabel;

            lblMinute1.Text = ConstantVariables.MinuteLabel1;
            lblMinute2.Text = ConstantVariables.MinuteLabel2;
            lblMinute3.Text = ConstantVariables.MinuteLabel3;

            grpBoxAlarmSettings.Text = ConstantVariables.AlarmSettingsGroupBox;
            lblTemperature.Text = ConstantVariables.TemperatureLabel;
            lblHumidity.Text = ConstantVariables.HumidityLabel;
            lblLowerAlarm.Text = ConstantVariables.LowerAlarmLabel;
            lblUpperAlram.Text = ConstantVariables.UpperAlarmLabel;
            lblCelcius1.Text = ConstantVariables.CelciusLabel1;
            lblCelcius2.Text = ConstantVariables.CelciusLabel2;
            lblRH1.Text = ConstantVariables.RHLabel1;
            lblRH2.Text = ConstantVariables.RHLabel2;

            grpBoxAlarms.Text = ConstantVariables.BatteryConsumeGroupBox;
            chkBoxLED.Text = ConstantVariables.LEDCheckBox;
            chkBoxDispOnTime.Text = ConstantVariables.DisplayOnTimeChkBox;

            grpBoxLoggerInfo.Text = ConstantVariables.LoggerInfoGroupBox;
            lblLoggerDateTime.Text = ConstantVariables.LoggerDateTimeLabel;
            lblFirmware.Text = ConstantVariables.FirmwareLabel;
            lblSerialNo.Text = ConstantVariables.SerialNoLabel;

            grpBoxCompSett.Text = ConstantVariables.CompanySettGroupBox;
            lblCompName.Text = ConstantVariables.CompanyNameLabel;
            lblCompLoc.Text = ConstantVariables.CompanyLocLabel;
            lblCompLogo.Text = ConstantVariables.CompanyLogoLabel;


            lblInfoSerialNo.Text = ConstantVariables.InfoSerialNoLabel;
            lblInfoMeasurements.Text = ConstantVariables.InfoMeasurementsLabel;
            lblInfoInterval.Text = ConstantVariables.IntervalLabel;
            lblInfoFrom.Text = ConstantVariables.InfoFromLabel;
            lblInfoTo.Text = ConstantVariables.InfoToLabel;
            lblInfoMinTemp.Text = ConstantVariables.InfoMinTempLabel;
            lblInfoMaxTemp.Text = ConstantVariables.InfoMaxTempLabel;
            lblInfoMinHumi.Text = ConstantVariables.InfoMinHumiLabel;
            lblInfoMaxHumi.Text = ConstantVariables.InfoMaxHumiLabel;
        }

        private void lblselect_Click(object sender, EventArgs e)
        {
            lblSelect.Enabled = false;
            lblProgram.Enabled = true;
            lblReadData.Enabled = true;
            lblShowDataTab.Enabled = true;
            lblShowDataChart.Enabled = true;

            naviBandInfo.Visible = false;

            lblLoggerName.Text = "";

            tabControl.TabPages.Add(tabPageSelect);
            tabControl.TabPages.Remove(tabPageProgram);
            tabControl.TabPages.Remove(tabPageAdminSett);
            tabControl.TabPages.Remove(tabPageReadData);
            tabControl.TabPages.Remove(tabPageShowDataChart);
            tabControl.TabPages.Remove(tabPageShowDataTab);

        }

        private void lblprogram_Click(object sender, EventArgs e)
        {
            lblProgram.Enabled = false;
            lblSelect.Enabled = true;
            lblReadData.Enabled = true;
            lblShowDataTab.Enabled = true;
            lblShowDataChart.Enabled = true;

            tabPageProgram.AutoScroll = true;

            naviBandInfo.Visible = false;

            Global.lblProgramWasClicked = true;

            tabControl.TabPages.Add(tabPageProgram);
            tabControl.TabPages.Add(tabPageAdminSett);
            tabControl.TabPages.Remove(tabPageSelect);
            tabControl.TabPages.Remove(tabPageReadData);
            tabControl.TabPages.Remove(tabPageShowDataChart);
            tabControl.TabPages.Remove(tabPageShowDataTab);
            try
            {
                btnsearchlog_Click(sender, e);
            }
            catch (Exception ex)
            { }
            displayGetValue();
        }

        /// <summary>
        /// Getting all values from device which is display on form
        /// </summary>
        public void displayGetValue()
        {
            try
            {
                devComm.sendCommandGetProgramLog();
                devComm.sendCommandGetProgLogger();
                devComm.sendCommandLoggerTime();
                devComm.sendCommandGetMinMaxRead();
                devComm.sendCommandGetLEDValue();

                txtboxremark.Text = Global.deviceNameStr.TrimEnd();

                //if (Global.deviceModeStr == "19")
                //{
                //    comboBoxType.Text = "ENDLESS WITH KEY PRESS ";
                //    dateTimePickstart1.Value = DateTime.Parse(Global.dateStrDecimalC, new CultureInfo("en-GB"));
                //    dateTimePickstart2.Value = DateTime.Parse(Global.timeStrDecimalC, new CultureInfo("en-GB"));
                //    dateTimePickstop1.Value = DateTime.Parse(Global.StDateStrDecimalC, new CultureInfo("en-GB"));
                //    dateTimePickstop2.Value = DateTime.Parse(Global.StTimeStrDecimalC, new CultureInfo("en-GB"));
                //}
                if (Global.deviceModeStr == "02")
                {
                    comboBoxType.Text = "MEASURE UPON START TIME ";
                    dateTimePickstart1.Value = DateTime.Parse(Global.dateStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstart2.Value = DateTime.Parse(Global.timeStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstop1.Value = DateTime.Parse(Global.StDateStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstop2.Value = DateTime.Parse(Global.StTimeStrDecimalC, new CultureInfo("en-GB"));
                }
                if (Global.deviceModeStr == "00")
                {
                    comboBoxType.Text = "START IMMEDIATELY UNTILL END OF MEMORY ";
                    dateTimePickstart1.Value = DateTime.Parse(Global.dateStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstart2.Value = DateTime.Parse(Global.timeStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstop1.Value = DateTime.Parse(Global.StDateStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstop2.Value = DateTime.Parse(Global.StTimeStrDecimalC, new CultureInfo("en-GB"));
                }
                if (Global.deviceModeStr == "06")
                {
                    comboBoxType.Text = "START / STOP MEASUREMENT WITH TIME ";
                    dateTimePickstart1.Value = DateTime.Parse(Global.dateStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstart2.Value = DateTime.Parse(Global.timeStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstop1.Value = DateTime.Parse(Global.StDateStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstop2.Value = DateTime.Parse(Global.StTimeStrDecimalC, new CultureInfo("en-GB"));

                }
                if (Global.deviceModeStr == "11")
                {
                    comboBoxType.Text = "START UPON KEY PRESS";
                    dateTimePickstart1.Value = DateTime.Parse(Global.dateStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstart2.Value = DateTime.Parse(Global.timeStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstop1.Value = DateTime.Parse(Global.StDateStrDecimalC, new CultureInfo("en-GB"));
                    dateTimePickstop2.Value = DateTime.Parse(Global.StTimeStrDecimalC, new CultureInfo("en-GB"));
                }

                numUpDwninterval.Value = Convert.ToDecimal(Global.intervalStrDecimal);

                lblLoggerDateTimeValue.Text = Global.loggerDateTime;

                if (Global.responseStringW == "" || Global.responseStringW == null)
                {
                    lblFirmwareValue.Text = "";
                }

                lblFirmwareValue.Text = Global.responseFirmVer;
                lblSerialNoValue.Text = Global.serialNoStr;

                numUpDwnstartdelay.Value = Convert.ToDecimal(Global.delayStrDecimal);

                numUpDwnTempMin.Value = Convert.ToDecimal(Global.tempLowStrDecimal);
                numUpDwnTempMax.Value = Convert.ToDecimal(Global.tempHighStrDecimal);

                if (Global.model == "PDL-K01")
                {
                    lblHumidity.Visible = false;
                    numUpDwnHumiMin.Visible = false;
                    lblRH1.Visible = false;
                    numUpDwnHumiMax.Visible = false;
                    lblRH2.Visible = false;
                }
                if (Global.model == "PDL-K03")
                {
                    lblHumidity.Visible = true;
                    numUpDwnHumiMin.Visible = true;
                    lblRH1.Visible = true;
                    numUpDwnHumiMax.Visible = true;
                    lblRH2.Visible = true;
                    numUpDwnHumiMin.Value = Convert.ToDecimal(Global.humiLowStrDecimal);
                    numUpDwnHumiMax.Value = Convert.ToDecimal(Global.humiHighStrDecimal);
                }

                if (Global.responseLEDValue == "00")
                {
                    chkBoxLED.Checked = true;
                }
                else if (Global.responseLEDValue == "01")
                {
                    chkBoxLED.Checked = false;
                }

                if (Convert.ToDecimal(Global.dispTimeStrDecimal) != 0)
                {
                    numUpDwnDispOnTime.Value = Convert.ToDecimal(Global.dispTimeStrDecimal);
                    chkBoxDispOnTime.Checked = true;
                }
                else
                {
                    chkBoxDispOnTime.Checked = false;
                }

                txtBoxCompName.Text = config.AppSettings.Settings["defaultCompanyName"].Value;
                txtBoxCompLoc.Text = config.AppSettings.Settings["defaultCompanyLoc"].Value;

                if (System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"] == "")
                {
                    txtBoxCompLogo.Text = Application.StartupPath + "\\enviro_logo.jpg";
                }
                else
                {
                    txtBoxCompLogo.Text = config.AppSettings.Settings["defaultCompanyLogo"].Value;
                }

            }
            catch (Exception e)
            {
                FileLog.ErrorLog(e.Message + e.StackTrace);
            }
        }

        private void lblreaddata_Click(object sender, EventArgs e)
        {
            lblReadData.Enabled = false;
            lblSelect.Enabled = true;
            lblProgram.Enabled = true;
            lblShowDataTab.Enabled = true;
            lblShowDataChart.Enabled = true;

            naviBandInfo.Visible = false;

            tabControl.TabPages.Add(tabPageReadData);
            tabControl.TabPages.Remove(tabPageSelect);
            tabControl.TabPages.Remove(tabPageProgram);
            tabControl.TabPages.Remove(tabPageAdminSett);
            tabControl.TabPages.Remove(tabPageShowDataChart);
            tabControl.TabPages.Remove(tabPageShowDataTab);
        }

        private void lblshowdatatab_Click(object sender, EventArgs e)
        {
            lblShowDataTab.Enabled = false;
            lblSelect.Enabled = true;
            lblProgram.Enabled = true;
            lblReadData.Enabled = true;
            lblShowDataChart.Enabled = true;

            Global.lblShowDataTabWasClicked = true;

            tabControl.TabPages.Add(tabPageShowDataTab);
            tabControl.TabPages.Remove(tabPageSelect);
            tabControl.TabPages.Remove(tabPageProgram);
            tabControl.TabPages.Remove(tabPageAdminSett);
            tabControl.TabPages.Remove(tabPageReadData);
            tabControl.TabPages.Remove(tabPageShowDataChart);
            dataInGridView();
            GridViewMinMaxCal();
            calculateMKT();

            if (dataGridView.RowCount != 0)
            {
                naviBandInfo.Visible = true;
                naviBar1.VisibleLargeButtons = 2;
                try
                {
                    Global.transactionCount = Global.readDataAllSplit.Length - Global.dateTimeTrasactionList.Count;

                    lblSerialNoInfoValue.Text = Global.serialNoStr;
                    lblMeasurementsInfoValue.Text = Global.transactionCount.ToString();
                    lblIntervalInfoValue.Text = Global.intervalStrDecimal;

                    if (Global.deviceModeStr == "02" || Global.deviceModeStr == "00" || Global.deviceModeStr == "06" || Global.deviceModeStr == "11")
                    {
                        lblFromInfoValue.Text = Global.dateStrDecimalC + "  " + Global.timeStrDecimalC;
                        lblToInfoValue.Text = Global.StDateStrDecimalC + "  " + Global.StTimeStrDecimalC;
                    }

                    lblMinTempInfoValue.Text = (Convert.ToDouble(Global.tempLowStrDecimal)).ToString("0.0") + "  " + "°C";
                    lblMaxTempInfoValue.Text = (Convert.ToDouble(Global.tempHighStrDecimal)).ToString("0.0") + "  " + "°C";

                    if (Global.model == "PDL-K01")
                    {
                        lblInfoMinHumi.Visible = false;
                        lblInfoMaxHumi.Visible = false;

                        lblMinHumiInfoValue.Visible = false;
                        lblMaxHumiInfoValue.Visible = false;
                    }

                    if (Global.model == "PDL-K03")
                    {
                        lblInfoMinHumi.Visible = true;
                        lblInfoMaxHumi.Visible = true;
                        lblInfoMinHumi.Text = ConstantVariables.InfoMinHumiLabel;
                        lblInfoMaxHumi.Text = ConstantVariables.InfoMaxHumiLabel;

                        lblMinHumiInfoValue.Visible = true;
                        lblMaxHumiInfoValue.Visible = true;
                        lblMinHumiInfoValue.Text = (Convert.ToDouble(Global.humiLowStrDecimal)).ToString("0.0") + "  " + "%RH";
                        lblMaxHumiInfoValue.Text = (Convert.ToDouble(Global.humiHighStrDecimal)).ToString("0.0") + "  " + "%RH";
                    }
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);
                }
            }
        }

        private void lblshowdatachart_Click(object sender, EventArgs e)
        {
            lblShowDataChart.Enabled = false;
            lblSelect.Enabled = true;
            lblProgram.Enabled = true;
            lblReadData.Enabled = true;
            lblShowDataTab.Enabled = true;

            naviBandInfo.Visible = true;
            naviBar1.VisibleLargeButtons = 2;

            Global.lblShowDataChartWasClicked = true;

            if (Global.lblShowDataTabWasClicked == false || Global.btnReadDataWasClicked == true)
            {
                lblshowdatatab_Click(sender, e);
                lblShowDataTab.Enabled = true;
                tabControl.TabPages.Remove(tabPageShowDataTab);
                lblShowDataChart.Enabled = false;
            }

            try
            {
                Global.transactionCount = Global.readDataAllSplit.Length - Global.dateTimeTrasactionList.Count;

                lblSerialNoInfoValue.Text = Global.serialNoStr;
                lblMeasurementsInfoValue.Text = Global.transactionCount.ToString();
                lblIntervalInfoValue.Text = Global.intervalStrDecimal;

                if (Global.deviceModeStr == "02" || Global.deviceModeStr == "00" || Global.deviceModeStr == "06" || Global.deviceModeStr == "11")
                {
                    lblFromInfoValue.Text = Global.dateStrDecimalC + "  " + Global.timeStrDecimalC;
                    lblToInfoValue.Text = Global.StDateStrDecimalC + "  " + Global.StTimeStrDecimalC;
                }

                lblMinTempInfoValue.Text = (Convert.ToDouble(Global.tempLowStrDecimal)).ToString("0.0") + "  " + "°C";
                lblMaxTempInfoValue.Text = (Convert.ToDouble(Global.tempHighStrDecimal)).ToString("0.0") + "  " + "°C";

                if (Global.model == "PDL-K01")
                {
                    lblInfoMinHumi.Visible = false;
                    lblInfoMaxHumi.Visible = false;

                    lblMinHumiInfoValue.Visible = false;
                    lblMaxHumiInfoValue.Visible = false;
                }

                if (Global.model == "PDL-K03")
                {
                    lblInfoMinHumi.Visible = true;
                    lblInfoMaxHumi.Visible = true;
                    lblInfoMinHumi.Text = ConstantVariables.InfoMinHumiLabel;
                    lblInfoMaxHumi.Text = ConstantVariables.InfoMaxHumiLabel;

                    lblMinHumiInfoValue.Visible = true;
                    lblMaxHumiInfoValue.Visible = true;
                    lblMinHumiInfoValue.Text = (Convert.ToDouble(Global.humiLowStrDecimal)).ToString("0.0") + "  " + "%RH";
                    lblMaxHumiInfoValue.Text = (Convert.ToDouble(Global.humiHighStrDecimal)).ToString("0.0") + "  " + "%RH";
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }

            GenerateGraph();

            tabControl.TabPages.Add(tabPageShowDataChart);
            tabControl.TabPages.Remove(tabPageSelect);
            tabControl.TabPages.Remove(tabPageProgram);
            tabControl.TabPages.Remove(tabPageAdminSett);
            tabControl.TabPages.Remove(tabPageReadData);
            tabControl.TabPages.Remove(tabPageShowDataTab);
        }

        #region Button

        private void btnsearchlog_Click(object sender, EventArgs e)
        {
            Global.btnSearchWasClicked = true;
            devComm.sendCommandDeviceSearch();
            devComm.sendCommandCheckSerialNo();
            try
            {
                if (Global.responseStringW == "" || Global.responseStringW == null)
                {
                    lblLoggerName.Text = "";
                    MessageBox.Show("Device is not connected", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else if (Global.responseDevName == "PDL-K03")
                {
                    lblLoggerName.Text = ConstantVariables.LoggerNameValueTR + "    /    " + Global.chkSerialNoStr;
                }
                else if (Global.responseDevName == "PDL-K01")
                {
                    lblLoggerName.Text = ConstantVariables.LoggerNameValueT + "    /    " + Global.chkSerialNoStr;
                }
            }

            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }

        }

        private void btnProgramLogger_Click(object sender, EventArgs e)
        {
            try
            {
                Global.CurrentDateTime();
                setStartStopDateTime();
                getTempSign();
                numUpDwnAllValues();
                getDisplayTime();
                devComm.sendCommandCheckSerialNo();

                if (chkBoxLED.Checked == true)
                {
                    Global.enableLEDValue = "00";
                }
                else if (chkBoxLED.Checked == false)
                {
                    Global.enableLEDValue = "10";
                }

                /////set configuration////////

                DialogResult dra = adm.ShowDialog();

                if (dra == System.Windows.Forms.DialogResult.OK)
                {

                    if (Global.chkSerialNoStr != Global.serialNoStr)
                    {
                        DialogResult dr1 = MessageBox.Show("Device is different . Do you want to Save previous device configuration to new device?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dr1 == DialogResult.Yes)
                        {
                            txtboxremark.Clear();
                            Global.serialNoStr = Global.chkSerialNoStr;
                            lblSerialNoValue.Text = Global.serialNoStr;
                        }
                        //else if (dr1 == DialogResult.No)
                        //{
                        //    txtboxremark.Clear();
                        //    dateTimePickstart1.Value = DateTime.Now.Date;
                        //    dateTimePickstop1.Value = DateTime.Now.Date;
                        //    dateTimePickstart2.Value = DateTime.Now;
                        //    dateTimePickstop2.Value = DateTime.Now;
                        //}
                    }
                    else if (Global.chkSerialNoStr == Global.serialNoStr)   /////0
                    {
                        if (Global.responseStringD != "")    /////1
                        {
                            if (!String.IsNullOrEmpty(txtboxremark.Text))    /////2
                            {
                                if (dateTimePickstart1.Value >= DateTime.Now.Date || comboBoxType.SelectedIndex == 1 || comboBoxType.SelectedIndex == 3)    /////3
                                {
                                    if (dateTimePickstop1.Value >= DateTime.Now.Date || comboBoxType.SelectedIndex == 0 || comboBoxType.SelectedIndex == 1 || comboBoxType.SelectedIndex == 3)   /////4
                                    {
                                        if (dateTimePickstart1.Value <= dateTimePickstop1.Value || comboBoxType.SelectedIndex == 0 || comboBoxType.SelectedIndex == 1 || comboBoxType.SelectedIndex == 3)     /////5
                                        {
                                            if (dateTimePickstart2.Value >= DateTime.Now && dateTimePickstart1.Value == DateTime.Now.Date || dateTimePickstart2.Value == dateTimePickstart2.Value && dateTimePickstart1.Value != DateTime.Now.Date || comboBoxType.SelectedIndex == 1 || comboBoxType.SelectedIndex == 3)    /////6
                                            {
                                                if (dateTimePickstop2.Value >= DateTime.Now && dateTimePickstop1.Value == DateTime.Now.Date || dateTimePickstop2.Value == dateTimePickstop2.Value && dateTimePickstop1.Value != DateTime.Now.Date || comboBoxType.SelectedIndex == 0 || comboBoxType.SelectedIndex == 1 || comboBoxType.SelectedIndex == 3)    /////7
                                                {
                                                    if (numUpDwnTempMin.Value <= numUpDwnTempMax.Value)    /////8
                                                    {
                                                        if (numUpDwnTempMax.Value >= numUpDwnTempMin.Value)    /////9
                                                        {

                                                            DialogResult dr = MessageBox.Show("This will erase all old data from device.  Do you want To Save configuration? ", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                                            if (dr == DialogResult.Yes)
                                                            {
                                                                devComm.sendCommandSetRemark();
                                                                devComm.sendCommandSetMode();
                                                                devComm.sendCommandSetInterDelLimi();
                                                                devComm.sendCommandEnableLED();
                                                                devComm.sendCommandDisplayOnTime();
                                                                devComm.sendCommandSyncTime();
                                                                devComm.sendCommandAllSett();
                                                                MessageBox.Show("To save configuration in device Please remove device from docket ", "Important Note", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                                                lblSelect.Enabled = false;
                                                                tabControl.TabPages.Add(tabPageSelect);
                                                                tabControl.TabPages.Remove(tabPageProgram);
                                                                tabControl.TabPages.Remove(tabPageAdminSett);
                                                                tabControl.TabPages.Remove(tabPageReadData);
                                                                tabControl.TabPages.Remove(tabPageShowDataChart);
                                                                tabControl.TabPages.Remove(tabPageShowDataTab);
                                                                lblProgram.Enabled = true;
                                                                lblLoggerName.Text = "";
                                                            }


                                                        }    /////9
                                                        else
                                                        {
                                                            MessageBox.Show("Maximum temperature value should be greater than minimum temperature value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                                            numUpDwnTempMax.Value = numUpDwnTempMin.Value;
                                                        }

                                                    }     /////8
                                                    else
                                                    {
                                                        MessageBox.Show("Minimum temperature value should be less than maximum temperature value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                                        numUpDwnTempMin.Value = numUpDwnTempMax.Value;
                                                    }
                                                }   /////7
                                                else
                                                {
                                                    MessageBox.Show("Please enter stop time after current time", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                                }

                                            }    /////6
                                            else
                                            {
                                                MessageBox.Show("Please enter start time after current time", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                            }

                                        }     /////5
                                        else
                                        {
                                            MessageBox.Show("Stop date should be next from start date", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        }

                                    }    /////4
                                    else
                                    {
                                        MessageBox.Show("Please enter stop date after today's date", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }

                                }    /////3
                                else
                                {
                                    MessageBox.Show("Please enter start date after today's date", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }

                            }   /////2
                            else
                            {
                                MessageBox.Show("Please enter device name", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }

                        }    /////1
                        else
                        {
                            MessageBox.Show("Device is not connected", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            txtboxremark.Clear();
                        }

                    }   /////0

                }
            }

            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        private void btnReadData_Click(object sender, EventArgs e)
        {
            try
            {
                Global.btnReadDataWasClicked = true;
                btnsearchlog_Click(sender, e);
                displayGetValue();

                Global.deviceNameStrList.Add(Global.deviceNameStr);
                Global.serialNoStrList.Add(Global.serialNoStr);
                Global.tempHighStrDecimalList.Add(Global.tempHighStrDecimal);
                Global.tempLowStrDecimalList.Add(Global.tempLowStrDecimal);
                Global.humiHighStrDecimalList.Add(Global.humiHighStrDecimal);
                Global.humiLowStrDecimalList.Add(Global.humiLowStrDecimal);

                if (Global.responseStringD != "")
                {
                    if (Convert.ToInt32(Global.transCountstrDecimal) != 0)
                    {
                        calculateThread();
                        backgroundWorker1.RunWorkerAsync();
                        devComm.sendCommandReadData();
                        this.Activate();
                        MessageBox.Show("Downloading Data Successfully", "Download", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                    }
                    else
                    {
                        MessageBox.Show("Device have no any Data", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            GenerateGraph();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openImage = new OpenFileDialog();
            openImage.Filter = "Image files (*.jpg, *.jpeg, *.jpe, *.jfif, *.png) | *.jpg; *.jpeg; *.jpe; *.jfif; *.png";
            if (openImage.ShowDialog() == DialogResult.OK)
            {
                txtBoxCompLogo.Text = openImage.FileName;
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            try
            {
                if (!String.IsNullOrEmpty(txtBoxCompName.Text))   ///1
                {
                    if (!String.IsNullOrEmpty(txtBoxCompLoc.Text))  ///2
                    {
                        //if (!String.IsNullOrEmpty(txtBoxCompLogo.Text))  ///3
                        //{
                        DialogResult dra = adm.ShowDialog();

                        if (dra == System.Windows.Forms.DialogResult.OK)
                        {
                            config.AppSettings.Settings["defaultCompanyName"].Value = txtBoxCompName.Text;
                            config.AppSettings.Settings["defaultCompanyLoc"].Value = txtBoxCompLoc.Text;

                            Global.companyLogoValue = txtBoxCompLogo.Text;

                            if (File.Exists(Directory.GetCurrentDirectory() + @"\enviro_logo.jpg"))
                            {
                                File.Delete(AppDomain.CurrentDomain.BaseDirectory + @"\enviro_logo.jpg");
                            }

                            System.IO.File.Copy(Global.companyLogoValue, @"enviro_logo.jpg");
                            string logoFilePath = AppDomain.CurrentDomain.BaseDirectory + @"\enviro_logo.jpg";

                            config.AppSettings.Settings["defaultCompanyLogo"].Value = logoFilePath;

                            config.Save(ConfigurationSaveMode.Modified, true);
                            ConfigurationManager.RefreshSection("appSettings");
                            MessageBox.Show("Saved successfully", "Save", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                            //txtBoxCompName.Text = "";
                            //txtBoxCompLoc.Text = "";
                            //txtBoxCompLogo.Text = "";
                        }

                        // }   ///3  
                        //else
                        //{
                        //    MessageBox.Show("Please enter the company logo", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        //}


                    }   ///2
                    else
                    {
                        MessageBox.Show("Please enter the company location", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                }   ///1
                else
                {
                    MessageBox.Show("Please enter the company name", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }


            }
            catch (FileNotFoundException ex)
            {
                txtBoxCompLogo.Text = "";
                MessageBox.Show("Please select the Logo", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ed)
            {
                FileLog.ErrorLog(ed.Message + ed.StackTrace);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Set the Date & time to device selected by user
        /// </summary>
        public void setStartStopDateTime()
        {
            try
            {
                if (comboBoxType.SelectedIndex == 0)
                {
                    string dtStartDay = dateTimePickstart1.Value.Day.ToString("00");
                    string dtStartMonth = dateTimePickstart1.Value.Month.ToString("00");
                    string startYear = dateTimePickstart1.Value.Year.ToString("00");
                    string dtStartYear = startYear.Substring(startYear.Length - 2, 2);
                    string timeStartHour = dateTimePickstart2.Value.Hour.ToString("00");
                    string timeStartMinute = dateTimePickstart2.Value.Minute.ToString("00");
                    Global.dateTimeStart = timeStartHour + timeStartMinute + dtStartDay + dtStartMonth + dtStartYear;

                }
                if (comboBoxType.SelectedIndex == 2)
                {
                    string dtStartDay = dateTimePickstart1.Value.Day.ToString("00");
                    string dtStartMonth = dateTimePickstart1.Value.Month.ToString("00");
                    string startYear = dateTimePickstart1.Value.Year.ToString("00");
                    string dtStartYear = startYear.Substring(startYear.Length - 2, 2);
                    string timeStartHour = dateTimePickstart2.Value.Hour.ToString("00");
                    string timeStartMinute = dateTimePickstart2.Value.Minute.ToString("00");
                    Global.dateTimeStart = timeStartHour + timeStartMinute + dtStartDay + dtStartMonth + dtStartYear;

                    string dtStopDay = dateTimePickstop1.Value.Day.ToString("00");
                    string dtStopMonth = dateTimePickstop1.Value.Month.ToString("00");
                    string stopYear = dateTimePickstop1.Value.Year.ToString("00");
                    string dtStopYear = stopYear.Substring(stopYear.Length - 2, 2);
                    string timeStopHour = dateTimePickstop2.Value.Hour.ToString("00");
                    string timeStopMinute = dateTimePickstop2.Value.Minute.ToString("00");
                    Global.dateTimeStop = timeStopHour + timeStopMinute + dtStopDay + dtStopMonth + dtStopYear;
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }

        }

        /// <summary>
        /// Values from numericUpDown control
        /// </summary>
        public void numUpDwnAllValues()
        {
            try
            {
                Global.intervalValue = (numUpDwninterval.Value).ToString().PadLeft(5, '0');
                Global.delayValue = (numUpDwnstartdelay.Value).ToString().PadLeft(5, '0');

                Global.valueTempMin = Math.Abs(Convert.ToInt32(numUpDwnTempMin.Value));
                Global.tempMinLimit = Global.valueTempMin.ToString().PadLeft(2, '0');

                Global.valueTempMax = Math.Abs(Convert.ToInt32(numUpDwnTempMax.Value));
                Global.tempMaxLimit = Global.valueTempMax.ToString().PadLeft(2, '0');

                if (Global.model == "PDL-K03")
                {
                    Global.humiMinLimit = (numUpDwnHumiMin.Value).ToString().PadLeft(2, '0');
                    Global.humiMaxLimit = (numUpDwnHumiMax.Value).ToString().PadLeft(2, '0');
                }
                if (Global.model == "PDL-K01")
                {
                    Global.humiMinLimit = (44).ToString();
                    Global.humiMaxLimit = (55).ToString();
                }

            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        /// <summary>
        /// Get the sign of min max temperature set by the user
        /// </summary>
        public void getTempSign()
        {
            try
            {
                if (numUpDwnTempMin.Value >= 0 && numUpDwnTempMax.Value >= 0)
                {
                    Global.tempSign = "00";
                }
                if (numUpDwnTempMin.Value < 0 && numUpDwnTempMax.Value >= 0)
                {
                    Global.tempSign = "01";
                }
                if (numUpDwnTempMin.Value < 0 && numUpDwnTempMax.Value < 0)
                {
                    Global.tempSign = "02";
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }

        }

        /// <summary>
        /// displayOnTime minute convert to hex value
        /// </summary>
        public void getDisplayTime()
        {
            try
            {
                string binaryValue;
                string binaryOfMin = Convert.ToString(Global.displayOnValue, 2).PadLeft(6, '0');
                if (Global.displayOnValue == 0)
                {
                    binaryValue = (binaryOfMin + "10").PadLeft(8, '0');
                }
                else
                {
                    binaryValue = (binaryOfMin + "11").PadLeft(8, '0');
                }
                Global.displayOnTimeValue = string.Format("{0:X2}", Convert.ToByte(binaryValue, 2));
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        /// <summary>
        /// Calculate MKT
        /// </summary>
        public void calculateMKT()
        {
            try
            {
                double temp1 = (273.1 + ((Global.MinTemp1 + Global.MaxTemp1) / 2));
                double tempa1 = Math.Exp(-10000 / temp1);
                double tempb1 = Math.Log(tempa1);
                Global.MKTValue1 = (-10000 / tempb1) - 273.1;

                double temp2 = (273.1 + ((Global.MinTemp2 + Global.MaxTemp2) / 2));
                double tempa2 = Math.Exp(-10000 / temp2);
                double tempb2 = Math.Log(tempa2);
                Global.MKTValue2 = (-10000 / tempb2) - 273.1;

                double temp3 = (273.1 + ((Global.MinTemp3 + Global.MaxTemp3) / 2));
                double tempa3 = Math.Exp(-10000 / temp3);
                double tempb3 = Math.Log(tempa3);
                Global.MKTValue3 = (-10000 / tempb3) - 273.1;

                double temp4 = (273.1 + ((Global.MinTemp4 + Global.MaxTemp4) / 2));
                double tempa4 = Math.Exp(-10000 / temp4);
                double tempb4 = Math.Log(tempa4);
                Global.MKTValue4 = (-10000 / tempb4) - 273.1;

            }
            catch (Exception ex)
            {
            }
        }

        /// <summary>
        /// Calculate thread value for read data
        /// </summary>
        public void calculateThread()
        {
            try
            {

                if (Convert.ToInt32(Global.transCountstrDecimal) <= 1000)
                {
                    Global.threadValue = 40000;          ////0.6 minute
                    Global.timerInterval = 1800;          ////0.3 seconds
                }
                if (Convert.ToInt32(Global.transCountstrDecimal) >= 1000 && Convert.ToInt32(Global.transCountstrDecimal) <= 5000)
                {
                    Global.threadValue = 90000;          ////1.5 minute
                    Global.timerInterval = 1200;          ////1.2 seconds
                }
                if (Convert.ToInt32(Global.transCountstrDecimal) >= 5000 && Convert.ToInt32(Global.transCountstrDecimal) <= 10000)
                {
                    Global.threadValue = 120000;          ////2 minute
                    Global.timerInterval = 1700;         ////1.7 seconds
                }
                if (Convert.ToInt32(Global.transCountstrDecimal) >= 10000 && Convert.ToInt32(Global.transCountstrDecimal) <= 15000)
                {
                    Global.threadValue = 165000;         ////2.75 minute
                    Global.timerInterval = 3000;         ////3 seconds                  
                }
                if (Convert.ToInt32(Global.transCountstrDecimal) >= 15000 && Convert.ToInt32(Global.transCountstrDecimal) <= 20000)
                {
                    Global.threadValue = 220000;         ////3.66 minute
                    Global.timerInterval = 4000;         ////4 seconds 
                }
                if (Convert.ToInt32(Global.transCountstrDecimal) >= 20000 && Convert.ToInt32(Global.transCountstrDecimal) <= 25000)
                {
                    Global.threadValue = 260000;        ////4.33 minute
                    Global.timerInterval = 5000;        ////5 seconds
                }
                if (Convert.ToInt32(Global.transCountstrDecimal) >= 25000 && Convert.ToInt32(Global.transCountstrDecimal) <= 35000)
                {
                    Global.threadValue = 360000;        ////6 minute
                    Global.timerInterval = 6000;        ////6 seconds 
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        /// <summary>
        /// Data Display in DataGridView
        /// </summary>
        public void dataInGridView()
        {
            try
            {
                DataRow dr;

                if (Global.filePathLOGList.Count != 0)
                {
                    if (dataGridView.RowCount != 0)
                    {
                        Global.dateStrList.Clear();
                        Global.timeStrList.Clear();
                        Global.temperatureList.Clear();
                        Global.humidityList.Clear();
                        Global.dateTimeTrasactionList.Clear();
                        Global.temperatureGraphList1.Clear();
                        Global.humidityGraphList1.Clear();
                        Global.temperatureGraphList2.Clear();
                        Global.humidityGraphList2.Clear();
                        Global.temperatureGraphList3.Clear();
                        Global.humidityGraphList3.Clear();
                        Global.temperatureGraphList4.Clear();
                        Global.humidityGraphList4.Clear();

                        dataGridView.DataSource = null;
                        dataGridView.Columns.Clear();
                        dataGridView.Rows.Clear();
                        dt.Columns.Clear();
                        dt.Clear();
                        dataGridView.Refresh();
                    }

                    #region DataTable 1

                    try
                    {
                        if (Global.filePathLOGList[0] != null)
                        {
                            try
                            {
                                Global.filePathCFG = Global.filePathCFGList[0];
                                Global.filePathLOG = Global.filePathLOGList[0];
                                devComm.ReadCFGFile();
                                devComm.ReadLOGFile();

                                Global.TCount1 = Global.readDataAllSplit.Length - Global.dateTimeTrasactionList.Count;
                                Global.fromDateTimeReport = Global.dateStrDecimalC + "  " + Global.timeStrDecimalC;
                                Global.toDateTimeReport = Global.StDateStrDecimalC + "  " + Global.StTimeStrDecimalC;

                                dt.Columns.Add(new DataColumn("Date & Time", typeof(string)));
                                dt.Columns.Add(new DataColumn(Global.serialNoStrList[0] + " " + "°C", typeof(float)));
                                if (Global.model == "PDL-K03")
                                {
                                    dt.Columns.Add(new DataColumn(Global.serialNoStrList[0] + " " + "%RH", typeof(float)));
                                }
                                Global.modelValue1 = Global.model;
                            }
                            catch (ArgumentOutOfRangeException ex)
                            {

                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                for (int j = 0; j < Global.TCount1; j++)
                                {
                                    string[] abc1 = Global.toDateTimeReport.Replace("  ", " ").Split(' ');
                                    string abc = abc1[0].Split('-')[2].ToString() + "-" + abc1[0].Split('-')[1].ToString() + "-" + abc1[0].Split('-')[0].ToString() + " " + abc1[1].ToString();
                                    dr = dt.NewRow();
                                    dr[0] = (Global.dateStrList[j].Split('-'))[2].ToString() + "-" + (Global.dateStrList[j].Split('-'))[1].ToString() + "-" + (Global.dateStrList[j].Split('-'))[0].ToString() + " " + Global.timeStrList[j];
                                    dr[1] = Global.temperatureList[j];
                                    Global.temperatureGraphList1.Add(Global.temperatureList[j]);
                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        dr[2] = Global.humidityList[j];
                                        Global.humidityGraphList1.Add(Global.humidityList[j]);
                                    }

                                    // Prevent from Extra Reading
                                    if (Convert.ToDateTime(dr[0]) <= Convert.ToDateTime(abc))
                                    {
                                        dr[0] = Global.dateStrList[j] + " " + Global.timeStrList[j];
                                        dt.Rows.Add(dr);
                                    }

                                }
                                Global.TCount1 = dt.Rows.Count;
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                            dataGridView.DataSource = dt;

                            #region Alarm Color
                            try
                            {
                                for (int k = 0; k < Global.TCount1; k++)
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[k].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        dataGridView.Rows[k].Cells[1].Style.ForeColor = System.Drawing.Color.Blue;
                                    }
                                    if (Convert.ToDecimal(dataGridView.Rows[k].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        dataGridView.Rows[k].Cells[1].Style.ForeColor = System.Drawing.Color.Red;
                                    }
                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                            #endregion

                            try
                            {
                                dataGridView.Columns[Global.serialNoStrList[0] + " " + "°C"].DefaultCellStyle.Format = "N1";
                                if (Global.modelValue1 == "PDL-K03")
                                {
                                    dataGridView.Columns[Global.serialNoStrList[0] + " " + "%RH"].DefaultCellStyle.Format = "N1";
                                }
                                this.dataGridView.CellBorderStyle = DataGridViewCellBorderStyle.Single;

                                this.dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }

                    #endregion;

                    #region DataTable 2

                    try
                    {
                        if (Global.filePathLOGList[1] != null)
                        {
                            try
                            {
                                Global.filePathCFG = Global.filePathCFGList[1];
                                Global.filePathLOG = Global.filePathLOGList[1];
                                devComm.ReadCFGFile();
                                devComm.ReadLOGFile();

                                Global.TCount2 = Global.readDataAllSplit.Length - Global.dateTimeTrasactionList.Count;

                                dt.Columns.Add(new DataColumn(Global.serialNoStrList[1] + " " + "°C", typeof(float)));
                                if (Global.model == "PDL-K03")
                                {
                                    dt.Columns.Add(new DataColumn(Global.serialNoStrList[1] + " " + "%RH", typeof(float)));
                                }

                                Global.modelValue2 = Global.model;
                            }
                            catch (ArgumentOutOfRangeException ex)
                            {

                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                for (int j = 0; j < Global.TCount2; j++)
                                {
                                    string cde = Global.toDateTimeReport;
                                    dr = dt.NewRow();
                                    if (Global.modelValue1 == "PDL-K01")
                                    {
                                        dataGridView.Rows[j].Cells[2].Value = Global.temperatureList[j];
                                        Global.temperatureGraphList2.Add(Global.temperatureList[j]);
                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            dataGridView.Rows[j].Cells[3].Value = Global.humidityList[j];
                                            Global.humidityGraphList2.Add(Global.humidityList[j]);
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K03")
                                    {
                                        dataGridView.Rows[j].Cells[3].Value = Global.temperatureList[j];
                                        Global.temperatureGraphList2.Add(Global.temperatureList[j]);
                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            dataGridView.Rows[j].Cells[4].Value = Global.humidityList[j];
                                            Global.humidityGraphList2.Add(Global.humidityList[j]);

                                        }
                                    }

                                    // Prevent from Extra Reading
                                    //if (Convert.ToDateTime(dr[0]) <= Convert.ToDateTime(cde))
                                    //{
                                    dt.Rows.Add(dr);
                                    //}

                                }
                                //  Global.TCount2 = dt.Rows.Count;
                            }
                            catch (ArgumentOutOfRangeException ax)
                            {

                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                            dataGridView.DataSource = dt;

                            #region Alarm Color
                            try
                            {
                                for (int k = 0; k < Global.TCount1; k++)
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[k].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        dataGridView.Rows[k].Cells[1].Style.ForeColor = System.Drawing.Color.Blue;
                                    }
                                    if (Convert.ToDecimal(dataGridView.Rows[k].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        dataGridView.Rows[k].Cells[1].Style.ForeColor = System.Drawing.Color.Red;
                                    }
                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                if (Global.modelValue1 == "PDL-K01")
                                {
                                    for (int k = 0; k < Global.TCount2; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03")
                                {
                                    for (int k = 0; k < Global.TCount2; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                            #endregion

                            try
                            {
                                dataGridView.Columns[Global.serialNoStrList[1] + " " + "°C"].DefaultCellStyle.Format = "N1";
                                if (Global.modelValue2 == "PDL-K03")
                                {
                                    dataGridView.Columns[Global.serialNoStrList[1] + " " + "%RH"].DefaultCellStyle.Format = "N1";
                                }
                                this.dataGridView.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                                this.dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }

                    #endregion

                    #region DataTable 3

                    try
                    {
                        if (Global.filePathLOGList[2] != null)
                        {
                            try
                            {
                                Global.filePathCFG = Global.filePathCFGList[2];
                                Global.filePathLOG = Global.filePathLOGList[2];
                                devComm.ReadCFGFile();
                                devComm.ReadLOGFile();

                                Global.TCount3 = Global.readDataAllSplit.Length - Global.dateTimeTrasactionList.Count;

                                dt.Columns.Add(new DataColumn(Global.serialNoStrList[2] + " " + "°C", typeof(float)));
                                if (Global.model == "PDL-K03")
                                {
                                    dt.Columns.Add(new DataColumn(Global.serialNoStrList[2] + " " + "%RH", typeof(float)));
                                }
                                Global.modelValue3 = Global.model;
                            }
                            catch (ArgumentOutOfRangeException ex)
                            {

                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                for (int j = 0; j < Global.TCount3; j++)
                                {
                                    //string abc = Global.toDateTimeReport;
                                    dr = dt.NewRow();

                                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                    {
                                        dataGridView.Rows[j].Cells[3].Value = Global.temperatureList[j];
                                        Global.temperatureGraphList3.Add(Global.temperatureList[j]);

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            dataGridView.Rows[j].Cells[4].Value = Global.humidityList[j];
                                            Global.humidityGraphList3.Add(Global.humidityList[j]);
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                    {
                                        dataGridView.Rows[j].Cells[5].Value = Global.temperatureList[j];
                                        Global.temperatureGraphList3.Add(Global.temperatureList[j]);

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            dataGridView.Rows[j].Cells[6].Value = Global.humidityList[j];
                                            Global.humidityGraphList3.Add(Global.humidityList[j]);
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                    {
                                        dataGridView.Rows[j].Cells[4].Value = Global.temperatureList[j];
                                        Global.temperatureGraphList3.Add(Global.temperatureList[j]);

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            dataGridView.Rows[j].Cells[5].Value = Global.humidityList[j];
                                            Global.humidityGraphList3.Add(Global.humidityList[j]);
                                        }
                                    }

                                    // Prevent from Extra Reading
                                    //if (Convert.ToDateTime(dr[0]) <= Convert.ToDateTime(abc))
                                    //{
                                    dt.Rows.Add(dr);
                                    //}
                                    // Global.TCount3 = dt.Rows.Count;
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                            dataGridView.DataSource = dt;

                            #region Alarm Color
                            try
                            {
                                for (int k = 0; k < Global.TCount1; k++)
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[k].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        dataGridView.Rows[k].Cells[1].Style.ForeColor = System.Drawing.Color.Blue;
                                    }
                                    if (Convert.ToDecimal(dataGridView.Rows[k].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        dataGridView.Rows[k].Cells[1].Style.ForeColor = System.Drawing.Color.Red;
                                    }
                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                if (Global.modelValue1 == "PDL-K01")
                                {
                                    for (int k = 0; k < Global.TCount2; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03")
                                {
                                    for (int k = 0; k < Global.TCount2; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                {
                                    for (int k = 0; k < Global.TCount3; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                {
                                    for (int k = 0; k < Global.TCount3; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[6].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[6].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                {
                                    for (int k = 0; k < Global.TCount3; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }

                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                            #endregion

                            try
                            {
                                dataGridView.Columns[Global.serialNoStrList[2] + " " + "°C"].DefaultCellStyle.Format = "N1";
                                if (Global.modelValue3 == "PDL-K03")
                                {
                                    dataGridView.Columns[Global.serialNoStrList[2] + " " + "%RH"].DefaultCellStyle.Format = "N1";
                                }
                                this.dataGridView.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                                this.dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }

                    #endregion

                    #region DataTable 4

                    try
                    {
                        if (Global.filePathLOGList[3] != null)
                        {
                            try
                            {
                                Global.filePathCFG = Global.filePathCFGList[3];
                                Global.filePathLOG = Global.filePathLOGList[3];
                                devComm.ReadCFGFile();
                                devComm.ReadLOGFile();

                                Global.TCount4 = Global.readDataAllSplit.Length - Global.dateTimeTrasactionList.Count;

                                dt.Columns.Add(new DataColumn(Global.serialNoStrList[3] + " " + "°C", typeof(float)));
                                if (Global.model == "PDL-K03")
                                {
                                    dt.Columns.Add(new DataColumn(Global.serialNoStrList[3] + " " + "%RH", typeof(float)));
                                }
                                Global.modelValue4 = Global.model;
                            }
                            catch (ArgumentOutOfRangeException ex)
                            {

                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                for (int j = 0; j < Global.TCount4; j++)
                                {
                                    //string abc = Global.toDateTimeReport;

                                    dr = dt.NewRow();

                                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                                    {
                                        dataGridView.Rows[j].Cells[4].Value = Global.temperatureList[j];
                                        Global.temperatureGraphList4.Add(Global.temperatureList[j]);

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            dataGridView.Rows[j].Cells[5].Value = Global.humidityList[j];
                                            Global.humidityGraphList4.Add(Global.humidityList[j]);
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                                    {
                                        dataGridView.Rows[j].Cells[7].Value = Global.temperatureList[j];
                                        Global.temperatureGraphList4.Add(Global.temperatureList[j]);

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            dataGridView.Rows[j].Cells[8].Value = Global.humidityList[j];
                                            Global.humidityGraphList4.Add(Global.humidityList[j]);
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                                    {
                                        dataGridView.Rows[j].Cells[6].Value = Global.temperatureList[j];
                                        Global.temperatureGraphList4.Add(Global.temperatureList[j]);

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            dataGridView.Rows[j].Cells[7].Value = Global.humidityList[j];
                                            Global.humidityGraphList4.Add(Global.humidityList[j]);
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                                    {
                                        dataGridView.Rows[j].Cells[5].Value = Global.temperatureList[j];
                                        Global.temperatureGraphList4.Add(Global.temperatureList[j]);

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            dataGridView.Rows[j].Cells[6].Value = Global.humidityList[j];
                                            Global.humidityGraphList4.Add(Global.humidityList[j]);
                                        }
                                    }

                                    // Prevent from Extra Reading
                                    //if (Convert.ToDateTime(dr[0]) <= Convert.ToDateTime(abc))
                                    //{
                                    dt.Rows.Add(dr);
                                    //}
                                    // Global.TCount4 = dt.Rows.Count;
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                            dataGridView.DataSource = dt;

                            #region Alarm Color
                            try
                            {
                                for (int k = 0; k < Global.TCount1; k++)
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[k].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        dataGridView.Rows[k].Cells[1].Style.ForeColor = System.Drawing.Color.Blue;
                                    }
                                    if (Convert.ToDecimal(dataGridView.Rows[k].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        dataGridView.Rows[k].Cells[1].Style.ForeColor = System.Drawing.Color.Red;
                                    }
                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                if (Global.modelValue1 == "PDL-K01")
                                {
                                    for (int k = 0; k < Global.TCount2; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[2].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03")
                                {
                                    for (int k = 0; k < Global.TCount2; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }

                            try
                            {
                                if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                {
                                    for (int k = 0; k < Global.TCount3; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[3].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                {
                                    for (int k = 0; k < Global.TCount3; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[6].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[6].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                {
                                    for (int k = 0; k < Global.TCount3; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }

                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                            try
                            {
                                if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                                {
                                    for (int k = 0; k < Global.TCount4; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            dataGridView.Rows[k].Cells[4].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                                {
                                    for (int k = 0; k < Global.TCount4; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[7].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            dataGridView.Rows[k].Cells[7].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[7].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            dataGridView.Rows[k].Cells[7].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[8].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                dataGridView.Rows[k].Cells[8].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[8].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                dataGridView.Rows[k].Cells[8].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                                {
                                    for (int k = 0; k < Global.TCount4; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[6].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            dataGridView.Rows[k].Cells[6].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[6].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            dataGridView.Rows[k].Cells[6].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[7].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                dataGridView.Rows[k].Cells[7].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[7].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                dataGridView.Rows[k].Cells[7].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                                {
                                    for (int k = 0; k < Global.TCount4; k++)
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Blue;
                                        }
                                        if (Convert.ToDecimal(dataGridView.Rows[k].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            dataGridView.Rows[k].Cells[5].Style.ForeColor = System.Drawing.Color.Red;
                                        }
                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                dataGridView.Rows[k].Cells[6].Style.ForeColor = System.Drawing.Color.Blue;
                                            }
                                            if (Convert.ToDecimal(dataGridView.Rows[k].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                dataGridView.Rows[k].Cells[6].Style.ForeColor = System.Drawing.Color.Red;
                                            }
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                            #endregion

                            try
                            {
                                dataGridView.Columns["Date & Time"].Width = 120;
                                dataGridView.Columns[Global.serialNoStrList[3] + " " + "°C"].DefaultCellStyle.Format = "N1";
                                if (Global.modelValue4 == "PDL-K03")
                                {
                                    dataGridView.Columns[Global.serialNoStrList[3] + " " + "%RH"].DefaultCellStyle.Format = "N1";
                                }
                                this.dataGridView.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                                this.dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                                for (int j = 1; j <= dataGridView.ColumnCount; j++)
                                {
                                    this.dataGridView.Columns[j].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
                                    dataGridView.Columns[j].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                                }

                            }
                            catch (Exception ex)
                            {
                                FileLog.ErrorLog(ex.Message + ex.StackTrace);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                    #endregion

                    #region DeleteEmptyRows
                    try
                    {
                        for (int i = 1; i < dataGridView.RowCount; i++)
                        {
                            Boolean isEmpty = true;
                            for (int j = 0; j < dataGridView.Columns.Count; j++)
                            {
                                if (dataGridView.Rows[i].Cells[j].Value.ToString() != "" && dataGridView.Rows[i].Cells[j].Value != null)
                                {
                                    isEmpty = false;
                                    break;
                                }
                            }
                            if (isEmpty)
                            {
                                dataGridView.Rows.RemoveAt(i);
                                i--;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }

                    #endregion
                }
                else
                {
                    MessageBox.Show("Please download the data", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            catch (Exception ex)
            {
            }

        }

        /// <summary>
        /// Calculate the Min,Max,Avg Temperature & Humidity
        /// </summary>
        public void GridViewMinMaxCal()
        {
            double sumTemp1 = 0;
            double sumHumi1 = 0;

            double sumTemp2 = 0;
            double sumHumi2 = 0;

            double sumTemp3 = 0;
            double sumHumi3 = 0;

            double sumTemp4 = 0;
            double sumHumi4 = 0;
            int dataCount = dataGridView.Rows.Count;

            #region GridViewMinMax 1

            try
            {

                for (int i = 0; i < Global.TCount1; i++)
                {
                    if (Global.MaxTemp1 < double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString()))
                    {
                        Global.MaxTemp1 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString());
                    }
                    else if (Global.MaxTemp1 == 0 && double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString()) < 0)
                    {
                        Global.MaxTemp1 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString());
                    }

                    if (i == 0)
                    {
                        Global.MinTemp1 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString());
                    }
                    if (Global.MinTemp1 > double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString()))
                    {
                        Global.MinTemp1 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString());
                    }

                    if (Global.modelValue1 == "PDL-K03")
                    {
                        if (Global.MaxHumi1 < double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "%RH"].Value.ToString()))
                        {
                            Global.MaxHumi1 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "%RH"].Value.ToString());
                        }

                        if (i == 0)
                        {
                            Global.MinHumi1 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "%RH"].Value.ToString());
                        }
                        if (Global.MinHumi1 > double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "%RH"].Value.ToString()))
                        {
                            Global.MinHumi1 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "%RH"].Value.ToString());
                        }

                    }
                    sumTemp1 = sumTemp1 + double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString());
                    Global.AvgTemp1 = (double)sumTemp1 / dataCount;
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        sumHumi1 = sumHumi1 + double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "%RH"].Value.ToString());
                        Global.AvgHumi1 = (double)sumHumi1 / dataCount;
                    }
                }

            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }

            #endregion

            #region GridViewMinMax 2

            try
            {

                for (int i = 0; i < Global.TCount2; i++)
                {
                    if (Global.MaxTemp2 < double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "°C"].Value.ToString()))
                    {
                        Global.MaxTemp2 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "°C"].Value.ToString());
                    }
                    else if (Global.MaxTemp2 == 0 && double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString()) < 0)
                    {
                        Global.MaxTemp2 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString());
                    }
                    if (i == 0)
                    {
                        Global.MinTemp2 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "°C"].Value.ToString());
                    }
                    if (Global.MinTemp2 > double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "°C"].Value.ToString()))
                    {
                        Global.MinTemp2 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "°C"].Value.ToString());
                    }

                    if (Global.modelValue2 == "PDL-K03")
                    {
                        if (Global.MaxHumi2 < double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "%RH"].Value.ToString()))
                        {
                            Global.MaxHumi2 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "%RH"].Value.ToString());
                        }

                        if (i == 0)
                        {
                            Global.MinHumi2 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "%RH"].Value.ToString());
                        }
                        if (Global.MinHumi2 > double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "%RH"].Value.ToString()))
                        {
                            Global.MinHumi2 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "%RH"].Value.ToString());
                        }

                    }
                    sumTemp2 = sumTemp2 + double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "°C"].Value.ToString());
                    Global.AvgTemp2 = (double)sumTemp2 / dataCount;
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        sumHumi2 = sumHumi2 + double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[1] + " " + "%RH"].Value.ToString());
                        Global.AvgHumi2 = (double)sumHumi2 / dataCount;
                    }
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }

            #endregion

            #region GridViewMinMax 3

            try
            {

                for (int i = 0; i < Global.TCount3; i++)
                {
                    if (Global.MaxTemp3 < double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "°C"].Value.ToString()))
                    {
                        Global.MaxTemp3 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "°C"].Value.ToString());
                    }
                    else if (Global.MaxTemp3 == 0 && double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString()) < 0)
                    {
                        Global.MaxTemp3 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString());
                    }

                    if (i == 0)
                    {
                        Global.MinTemp3 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "°C"].Value.ToString());
                    }
                    if (Global.MinTemp3 > double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "°C"].Value.ToString()))
                    {
                        Global.MinTemp3 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "°C"].Value.ToString());
                    }

                    if (Global.modelValue3 == "PDL-K03")
                    {
                        if (Global.MaxHumi3 < double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "%RH"].Value.ToString()))
                        {
                            Global.MaxHumi3 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "%RH"].Value.ToString());
                        }

                        if (i == 0)
                        {
                            Global.MinHumi3 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "%RH"].Value.ToString());
                        }
                        if (Global.MinHumi3 > double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "%RH"].Value.ToString()))
                        {
                            Global.MinHumi3 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "%RH"].Value.ToString());
                        }

                    }
                    sumTemp3 = sumTemp3 + double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "°C"].Value.ToString());
                    Global.AvgTemp3 = (double)sumTemp3 / dataCount;
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        sumHumi3 = sumHumi3 + double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[2] + " " + "%RH"].Value.ToString());
                        Global.AvgHumi3 = (double)sumHumi3 / dataCount;
                    }
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }

            #endregion

            #region GridViewMinMax 4

            try
            {

                for (int i = 0; i < Global.TCount4; i++)
                {
                    if (Global.MaxTemp4 < double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "°C"].Value.ToString()))
                    {
                        Global.MaxTemp4 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "°C"].Value.ToString());
                    }
                    else if (Global.MaxTemp4 == 0 && double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString()) < 0)
                    {
                        Global.MaxTemp4 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[0] + " " + "°C"].Value.ToString());
                    }

                    if (i == 0)
                    {
                        Global.MinTemp4 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "°C"].Value.ToString());
                    }
                    if (Global.MinTemp4 > double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "°C"].Value.ToString()))
                    {
                        Global.MinTemp4 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "°C"].Value.ToString());
                    }

                    if (Global.modelValue4 == "PDL-K03")
                    {
                        if (Global.MaxHumi4 < double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "%RH"].Value.ToString()))
                        {
                            Global.MaxHumi4 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "%RH"].Value.ToString());
                        }

                        if (i == 0)
                        {
                            Global.MinHumi4 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "%RH"].Value.ToString());
                        }
                        if (Global.MinHumi4 > double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "%RH"].Value.ToString()))
                        {
                            Global.MinHumi4 = double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "%RH"].Value.ToString());
                        }

                    }
                    sumTemp4 = sumTemp4 + double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "°C"].Value.ToString());
                    Global.AvgTemp4 = (double)sumTemp4 / dataCount;
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        sumHumi4 = sumHumi4 + double.Parse(dataGridView.Rows[i].Cells[Global.serialNoStrList[3] + " " + "%RH"].Value.ToString());
                        Global.AvgHumi4 = (double)sumHumi4 / dataCount;
                    }
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }

            #endregion

            Global.calMinTempList.Add(Global.MinTemp1);
            Global.calMinTempList.Add(Global.MinTemp2);
            Global.calMinTempList.Add(Global.MinTemp3);
            Global.calMinTempList.Add(Global.MinTemp4);

            Global.calMaxTempList.Add(Global.MaxTemp1);
            Global.calMaxTempList.Add(Global.MaxTemp2);
            Global.calMaxTempList.Add(Global.MaxTemp3);
            Global.calMaxTempList.Add(Global.MaxTemp4);

            Global.calMinHumiList.Add(Global.MinHumi1);
            Global.calMinHumiList.Add(Global.MinHumi2);
            Global.calMinHumiList.Add(Global.MinHumi3);
            Global.calMinHumiList.Add(Global.MinHumi4);

            Global.calMaxHumiList.Add(Global.MaxHumi1);
            Global.calMaxHumiList.Add(Global.MaxHumi2);
            Global.calMaxHumiList.Add(Global.MaxHumi3);
            Global.calMaxHumiList.Add(Global.MaxHumi4);
        }

        /// <summary>
        /// Generate graph
        /// </summary>
        public void GenerateGraph()
        {
            if (zedGraphControl.GraphPane.CurveList != null)
            {
                zedGraphControl.GraphPane.CurveList.Clear();
                zedGraphControl.GraphPane.GraphObjList.Clear();
            }
            try
            {
                DateTime DTime;
                GraphPane myPane = zedGraphControl.GraphPane;

                graphTempColorList.Add(System.Drawing.Color.DarkGreen);
                graphTempColorList.Add(System.Drawing.Color.HotPink);
                graphTempColorList.Add(System.Drawing.Color.Red);
                graphTempColorList.Add(System.Drawing.Color.Yellow);
                graphHumiColorList.Add(System.Drawing.Color.DarkBlue);
                graphHumiColorList.Add(System.Drawing.Color.Maroon);
                graphHumiColorList.Add(System.Drawing.Color.DeepSkyBlue);
                graphHumiColorList.Add(System.Drawing.Color.Black);

                if (Global.model == "PDL-K03")
                {
                    zedGraphControl.GraphPane.Title.Text = "   Temperature  & Humidity Versus Date & Time Graph";
                    zedGraphControl.GraphPane.Y2Axis.Title.Text = "Humidity";
                }
                else if (Global.model == "PDL-K01")
                {
                    zedGraphControl.GraphPane.Title.Text = "   Temperature Versus Date & Time Graph";

                }
                zedGraphControl.GraphPane.XAxis.Title.Text = "Date & Time";
                zedGraphControl.GraphPane.YAxis.Title.Text = "Temperature";
                //Global.majorStepCount = 1;
                /// <summary>
                /// X Axis For Date & Time
                /// </summary>
                myPane.XAxis.Scale.Format = "dd-MM-yy HH:mm";
                myPane.XAxis.Type = AxisType.Date;
                myPane.XAxis.Scale.MinorUnit = DateUnit.Day;
                myPane.XAxis.Scale.MajorUnit = DateUnit.Minute;
                myPane.XAxis.Scale.MajorStep = Global.majorStepCount;
                myPane.XAxis.Scale.MinAuto = true;
                myPane.XAxis.Scale.MaxAuto = true;
                myPane.XAxis.Scale.FontSpec.Angle = 90;
                try
                {
                    myPane.XAxis.Scale.Min = new XDate(DateTime.Parse(dataGridView.Rows[0].Cells[0].Value.ToString(), new CultureInfo("en-GB")));
                    myPane.XAxis.Scale.Max = new XDate(DateTime.Parse(dataGridView.Rows[dataGridView.Rows.Count - 1].Cells[0].Value.ToString(), new CultureInfo("en-GB")));
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);
                }
                myPane.XAxis.MinorTic.IsAllTics = false;
                myPane.XAxis.MajorTic.IsOpposite = false;
                myPane.XAxis.MinorTic.IsOpposite = false;

                /// <summary>
                /// Y Axis For Temperature
                /// </summary>
                myPane.YAxis.Scale.MaxAuto = true;
                myPane.YAxis.Scale.MinAuto = true;
                myPane.YAxis.Scale.Min = Convert.ToDouble(Global.calMinTempList.Min()) - 10;
                myPane.YAxis.Scale.Max = Convert.ToDouble(Global.calMaxTempList.Max()) + 10;
                myPane.YAxis.MinorTic.IsAllTics = false;
                myPane.YAxis.Scale.MajorStep = 5;
                myPane.YAxis.MinorTic.IsOpposite = false;
                myPane.YAxis.MajorTic.IsOpposite = false;

                PointPairList tempPointList1 = new PointPairList();
                PointPairList tempPointList2 = new PointPairList();
                PointPairList tempPointList3 = new PointPairList();
                PointPairList tempPointList4 = new PointPairList();
                PointPairList humiPointList1 = new PointPairList();
                PointPairList humiPointList2 = new PointPairList();
                PointPairList humiPointList3 = new PointPairList();
                PointPairList humiPointList4 = new PointPairList();

                #region Temperture Graph

                if (Global.temperatureList.Count > 0)
                {
                    if (Global.serialNoStrList.Count >= 1)
                    {
                        for (int i = 0; i < dataGridView.RowCount; i++)
                        {
                            DTime = DateTime.Parse(Global.dateStrList[i] + " " + Global.timeStrList[i], new CultureInfo("en-GB"));

                            double dblTemperature1 = Double.Parse(Global.temperatureGraphList1[i]);
                            tempPointList1.Add(new XDate(DTime), dblTemperature1);

                        }
                        LineItem Temperature1 = myPane.AddCurve(Global.serialNoStrList[0] + " °C", tempPointList1, graphTempColorList[0], SymbolType.None);
                        Temperature1.Line.Width = 2f;
                    }
                    if (Global.serialNoStrList.Count >= 2)
                    {
                        for (int h = 0; h < dataGridView.RowCount; h++)
                        {
                            DTime = DateTime.Parse(Global.dateStrList[h] + " " + Global.timeStrList[h], new CultureInfo("en-GB"));

                            double dblTemperature2 = Double.Parse(Global.temperatureGraphList2[h]);
                            tempPointList2.Add(new XDate(DTime), dblTemperature2);

                        }
                        LineItem Temperature2 = myPane.AddCurve(Global.serialNoStrList[1] + " °C", tempPointList2, graphTempColorList[1], SymbolType.None);
                        Temperature2.Line.Width = 2f;
                    }
                    if (Global.serialNoStrList.Count >= 3)
                    {
                        for (int g = 0; g < dataGridView.RowCount; g++)
                        {
                            DTime = DateTime.Parse(Global.dateStrList[g] + " " + Global.timeStrList[g], new CultureInfo("en-GB"));

                            double dblTemperature3 = Double.Parse(Global.temperatureGraphList3[g]);
                            tempPointList3.Add(new XDate(DTime), dblTemperature3);

                        }
                        LineItem Temperature3 = myPane.AddCurve(Global.serialNoStrList[2] + " °C", tempPointList3, graphTempColorList[2], SymbolType.None);
                        Temperature3.Line.Width = 2f;
                    }
                    if (Global.serialNoStrList.Count == 4)
                    {
                        for (int n = 0; n < dataGridView.RowCount; n++)
                        {
                            DTime = DateTime.Parse(Global.dateStrList[n] + " " + Global.timeStrList[n], new CultureInfo("en-GB"));

                            double dblTemperature4 = Double.Parse(Global.temperatureGraphList4[n]);
                            tempPointList4.Add(new XDate(DTime), dblTemperature4);

                        }
                        LineItem Temperature4 = myPane.AddCurve(Global.serialNoStrList[3] + " °C", tempPointList4, graphTempColorList[3], SymbolType.None);
                        Temperature4.Line.Width = 2f;
                    }

                }

                #endregion

                /// <summary>
                /// Y2 Axis For Humidity
                /// </summary>
                /// 
                if (Global.model == "PDL-K03")
                {
                    myPane.Y2Axis.IsVisible = true;
                    myPane.Y2Axis.Scale.MaxAuto = true;
                    myPane.Y2Axis.Scale.MinAuto = true;
                    myPane.Y2Axis.Scale.Min = Convert.ToDouble(Global.calMinHumiList.Min()) - 10;
                    myPane.Y2Axis.Scale.Max = Convert.ToDouble(Global.calMaxHumiList.Max()) + 10;
                    myPane.Y2Axis.MinorTic.IsAllTics = false;
                    myPane.Y2Axis.Scale.MajorStep = 5;
                    myPane.Y2Axis.MinorTic.IsOpposite = false;
                    myPane.Y2Axis.MajorTic.IsOpposite = false;

                    #region Humidity Graph

                    if (Global.temperatureList.Count > 0)
                    {
                        if (Global.serialNoStrList.Count >= 1)
                        {
                            if (Global.modelValue1 == "PDL-K03")
                            {
                                for (int j = 0; j < dataGridView.RowCount; j++)
                                {
                                    DTime = DateTime.Parse(Global.dateStrList[j] + " " + Global.timeStrList[j], new CultureInfo("en-GB"));

                                    double dblHumidity1 = Double.Parse(Global.humidityGraphList1[j]);
                                    humiPointList1.Add(new XDate(DTime), dblHumidity1);
                                }
                                LineItem Humidity1 = myPane.AddCurve(Global.serialNoStrList[0] + " %RH", humiPointList1, graphHumiColorList[0], SymbolType.None);
                                Humidity1.IsY2Axis = true;
                                Humidity1.Line.Width = 2f;
                            }
                        }
                        if (Global.serialNoStrList.Count >= 2)
                        {
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                for (int k = 0; k < dataGridView.RowCount; k++)
                                {
                                    DTime = DateTime.Parse(Global.dateStrList[k] + " " + Global.timeStrList[k], new CultureInfo("en-GB"));

                                    double dblHumidity2 = Double.Parse(Global.humidityGraphList2[k]);
                                    humiPointList2.Add(new XDate(DTime), dblHumidity2);
                                }
                                LineItem Humidity2 = myPane.AddCurve(Global.serialNoStrList[1] + " %RH", humiPointList2, graphHumiColorList[1], SymbolType.None);
                                Humidity2.IsY2Axis = true;
                                Humidity2.Line.Width = 2f;
                            }
                        }
                        if (Global.serialNoStrList.Count >= 3)
                        {
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                for (int p = 0; p < dataGridView.RowCount; p++)
                                {
                                    DTime = DateTime.Parse(Global.dateStrList[p] + " " + Global.timeStrList[p], new CultureInfo("en-GB"));

                                    double dblHumidity3 = Double.Parse(Global.humidityGraphList3[p]);
                                    humiPointList3.Add(new XDate(DTime), dblHumidity3);
                                }
                                LineItem Humidity3 = myPane.AddCurve(Global.serialNoStrList[2] + " %RH", humiPointList3, graphHumiColorList[2], SymbolType.None);
                                Humidity3.IsY2Axis = true;
                                Humidity3.Line.Width = 2f;
                            }
                        }
                        if (Global.serialNoStrList.Count == 4)
                        {
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                for (int q = 0; q < dataGridView.RowCount; q++)
                                {
                                    DTime = DateTime.Parse(Global.dateStrList[q] + " " + Global.timeStrList[q], new CultureInfo("en-GB"));

                                    double dblHumidity4 = Double.Parse(Global.humidityGraphList4[q]);
                                    humiPointList4.Add(new XDate(DTime), dblHumidity4);
                                }
                                LineItem Humidity4 = myPane.AddCurve(Global.serialNoStrList[3] + " %RH", humiPointList4, graphHumiColorList[3], SymbolType.None);
                                Humidity4.IsY2Axis = true;
                                Humidity4.Line.Width = 2f;
                            }
                        }
                    }

                    #endregion
                }

                zedGraphControl.Refresh();
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        class MyHeaderFooterEvent : iTextSharp.text.pdf.PdfPageEventHelper
        {
            // iTextSharp.text.Font FONT = iTextSharp.text.FontFactory.GetFont("Arial", 12);      //Times New Roman     
            public override void OnStartPage(iTextSharp.text.pdf.PdfWriter writer, iTextSharp.text.Document document)
            {
                base.OnStartPage(writer, document);
                iTextSharp.text.Rectangle pageSize = document.PageSize;
                pageSize.GetLeft(280);
                pageSize.GetTop(80);
            }
            public override void OnEndPage(iTextSharp.text.pdf.PdfWriter writer, iTextSharp.text.Document document)
            {
                base.OnEndPage(writer, document);
                iTextSharp.text.pdf.PdfContentByte cb;
                iTextSharp.text.pdf.PdfTemplate headerTemplate, footerTemplate;
                iTextSharp.text.pdf.BaseFont bf = null;
                DateTime PrintTime = Convert.ToDateTime(DateTime.Now.ToString("dd-MM-yyyy HH:mm"));
                bf = iTextSharp.text.pdf.BaseFont.CreateFont(iTextSharp.text.pdf.BaseFont.HELVETICA, iTextSharp.text.pdf.BaseFont.CP1252, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED);
                cb = writer.DirectContent;
                headerTemplate = cb.CreateTemplate(100, 100);
                footerTemplate = cb.CreateTemplate(50, 50);
                // FOOTER:
                local.Toatalpagecount = writer.PageNumber;
                //String text = "Page No " + pageN;
                //float len = bf.GetWidthPoint(text, 12);
                iTextSharp.text.Rectangle pageSize = document.PageSize;
                cb.BeginText();
                cb.SetFontAndSize(bf, 8);
                cb.SetTextMatrix(pageSize.GetRight(60), pageSize.GetBottom(30));
                //cb.ShowText(text);
                cb.EndText();
                //cb.AddTemplate(footerTemplate, pageSize.GetRight(60) + len, pageSize.GetBottom(30));
            }
            public override void OnCloseDocument(iTextSharp.text.pdf.PdfWriter writer, iTextSharp.text.Document document)
            {
                base.OnCloseDocument(writer, document);
            }
        }




        /// <summary>
        /// Create PDF 
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="FileName"></param>
        public void CreateUploadGraphNData(DataGridView dataGridView, string FileName)
        {
            # region Old PDF report
            //System.IO.FileStream fs = null;
            //try
            //{
            //    SaveFileDialog sfd = new SaveFileDialog();
            //    sfd.Filter = "PDF Documents (*.pdf)|*.pdf";
            //    sfd.FileName = FileName;
            //    if (sfd.ShowDialog() == DialogResult.OK)
            //    //DialogResult dlgResult = saveDlg.ShowDialog();
            //    {
            //        Global.printPDFName = sfd.FileName;
            //        fs = new FileStream(sfd.FileName, FileMode.Create);
            //        // Create an instance of the document class which represents the PDF document itself.
            //        iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30, 30, 30, 30);
            //        // Create an instance to the PDF file by creating an instance of the PDF 
            //        // Writer class using the document and the filestrem in the constructor.
            //        iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, fs);
            //        HeaderFooter header = new HeaderFooter(new Phrase("Portable Data Logger "), false); //+ SerialNumberTextBox.Text + (" + txtboxremark + ")
            //        header.Border = iTextSharp.text.Rectangle.NO_BORDER; ;
            //        document.Header = header;
            //        HeaderFooter footer = new HeaderFooter(new Phrase(string.Format("Prepared By :{0}                      Reviewed By :{1}                      Date :{2}     Page :  ", "", "", DateTime.Now.ToString("dd/MM/yyyy HH:mm"))), true);
            //        footer.Border = iTextSharp.text.Rectangle.NO_BORDER; ;
            //        document.Footer = footer;
            //        // ------------------------ Adding Logo in First Page------------------------------------------------------------------------------------------------------
            //        // Adding Logo in PDf//
            //        System.Drawing.Image image = System.Drawing.Image.FromFile(Application.StartupPath + "\\enviro_logo.jpg");
            //        Document doc = new Document(PageSize.A4);
            //        //PdfWriter.GetInstance(doc, new FileStream("image3.pdf", FileMode.Create));
            //        document.Open();
            //        iTextSharp.text.Image pdfImage = iTextSharp.text.Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Jpeg);
            //        pdfImage.ScaleToFit(100, 50);
            //        //If you want to choose image as background then,
            //        // pdfImage.Alignment = iTextSharp.text.Image.UNDERLYING;
            //        //If you want to give absolute/specified fix position to image.
            //        pdfImage.SetAbsolutePosition(400, 800);
            //        document.Add(pdfImage);
            //        // Add a simple and wellknown phrase to the document in a flow layout manner
            //        iTextSharp.text.Font font5 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 18, 1);
            //        iTextSharp.text.Font font4 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 12, 1);
            //        iTextSharp.text.Font font3 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 0);
            //        iTextSharp.text.Font font2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1);
            //        iTextSharp.text.Font font1 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 0);
            //        iTextSharp.text.Font Redfont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.RED);
            //        iTextSharp.text.Font Bluefont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.BLUE);
            //        // cb.SetLineWidth(1.0f);   // Make a bit thicker than 1.0 default
            //        //cb.SetGrayStroke(1.0f); // 1 = black, 0 = white
            //        PdfContentByte cb = writer.DirectContent;
            //        document.Add(new iTextSharp.text.Paragraph("Data Report", font5));
            //        // ------------------------ Device Information------------------------------------------------------------------------------------------------------
            //        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        document.Add(new iTextSharp.text.Paragraph("Device Information", font4));
            //        cb.MoveTo(30, document.Top - 103);
            //        cb.LineTo(550, document.Top - 103);
            //        cb.Stroke();
            //        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        PdfPTable table1 = null;
            //        float[] columnDefinitionSize1 = { 15F, 15F, 15F, 15F };
            //        table1 = new PdfPTable(columnDefinitionSize1);
            //        table1.WidthPercentage = 100;
            //        table1.DefaultCell.BorderColor = iTextSharp.text.Color.WHITE;
            //        table1.AddCell(new Phrase(String.Format("{0}", "Device Model No/Version:"), font2));
            //        if (Global.responseDevName == "PDL-K03")
            //        {
            //            table1.AddCell(new Phrase(String.Format("{0}", ConstantVariables.LoggerNameValueTR), font1));
            //        }
            //        else if (Global.responseDevName == "PDL-K01")
            //        {
            //            table1.AddCell(new Phrase(String.Format("{0}", ConstantVariables.LoggerNameValueT), font1));
            //        }
            //        if (Global.deviceModeStr == "11")
            //        {
            //            table1.AddCell(new Phrase(String.Format("{0}", "Start Delay:"), font2));
            //            table1.AddCell(new Phrase(String.Format("{0} minute", Global.delayStrDecimal), font1));
            //        }
            //        else
            //        {
            //            table1.AddCell(new Phrase(String.Format("{0}", "Start Delay:"), font2));
            //            table1.AddCell(new Phrase(String.Format("{0}", "-"), font1));
            //        }
            //        table1.AddCell(new Phrase(String.Format("{0}", "Serial Number:"), font2));
            //        table1.AddCell(new Phrase(String.Format("{0}", Global.serialNoStr), Redfont2));
            //        table1.AddCell(new Phrase(String.Format("{0}", "Sampling Period:"), font2));
            //        table1.AddCell(new Phrase(String.Format("{0} minute", Global.intervalStrDecimal), font1));
            //        table1.AddCell(new Phrase(String.Format("{0}", "Minimum Temperature Alarm Limit:"), font2));
            //        table1.AddCell(new Phrase(String.Format("{0} °C", Convert.ToDouble(Global.tempLowStrDecimal).ToString("0.0")), font1));
            //        if (Global.model == "PDL-K03")
            //        {
            //            table1.AddCell(new Phrase(String.Format("{0}", "Minimum Humidity Alarm Limit:"), font2));
            //            table1.AddCell(new Phrase(String.Format("{0} %RH", Convert.ToDouble(Global.humiLowStrDecimal).ToString("0.0")), font1));
            //        }
            //        table1.AddCell(new Phrase(String.Format("{0}", "Maximum  Temperature Alarm Limit:"), font2));
            //        table1.AddCell(new Phrase(String.Format("{0} °C", Convert.ToDouble(Global.tempHighStrDecimal).ToString("0.0")), font1));
            //        if (Global.model == "PDL-K03")
            //        {
            //            table1.AddCell(new Phrase(String.Format("{0}", "Maximum Humidity Alarm Limit:"), font2));
            //            table1.AddCell(new Phrase(String.Format("{0} %RH ", Convert.ToDouble(Global.humiHighStrDecimal).ToString("0.0")), font1));
            //        }
            //        table1.AddCell(new Phrase(String.Format("{0}", "Device Name:"), font2));
            //        table1.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStr), font1));
            //        table1.AddCell(new Phrase(String.Format(" "), font2));
            //        table1.AddCell(new Phrase(String.Format(" "), font1));
            //        document.Add(table1);
            //        if (Global.model == "PDL-K01")
            //        {
            //            document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        }
            //        // ------------------------ Reading Summary------------------------------------------------------------------------------------------------------
            //        document.Add(new iTextSharp.text.Paragraph("Reading Summary", font4));
            //        cb.MoveTo(30, document.Top - 230);
            //        cb.LineTo(550, document.Top - 230);
            //        cb.Stroke();
            //        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        PdfPTable table2 = null;
            //        float[] columnDefinitionSize2 = { 15F, 15F, 15F, 15F };
            //        table2 = new PdfPTable(columnDefinitionSize2);
            //        table2.WidthPercentage = 95;
            //        table2.DefaultCell.BorderColor = iTextSharp.text.Color.WHITE;
            //        table2.AddCell(new Phrase(String.Format("{0}", "Transaction Count:"), font2));
            //        table2.AddCell(new Phrase(String.Format("{0}", Global.transactionCount), font1));
            //        table2.AddCell(new Phrase(String.Format("{0}", "Minimum Temperature:"), font2));
            //        table2.AddCell(new Phrase(String.Format("{0:0.0} °C", Global.MinTemp), font1));
            //        if (Global.deviceModeStr == "02" || Global.deviceModeStr == "00" || Global.deviceModeStr == "06" || Global.deviceModeStr == "11")
            //        {
            //            table2.AddCell(new Phrase(String.Format("{0}", "Start Date & Time:"), font2));
            //            //table2.AddCell(new Phrase(String.Format("{0}", StartTimeTextBox.Text+" "+"(GMT)"), font1));
            //            table2.AddCell(new Phrase(String.Format("{0}", Global.dateStrDecimalC + " " + Global.timeStrDecimalC), font1));
            //        }
            //        table2.AddCell(new Phrase(String.Format("{0}", "Maximum Temperature:"), font2));
            //        table2.AddCell(new Phrase(String.Format("{0:0.0} °C", Global.MaxTemp), font1));
            //        if (Global.deviceModeStr == "02" || Global.deviceModeStr == "00" || Global.deviceModeStr == "06" || Global.deviceModeStr == "11")
            //        {
            //            table2.AddCell(new Phrase(String.Format("{0}", "Stop Date & Time:"), font2));
            //            //table2.AddCell(new Phrase(String.Format("{0}", StartTimeTextBox.Text+" "+"(GMT)"), font1));
            //            table2.AddCell(new Phrase(String.Format("{0}", Global.StDateStrDecimalC + " " + Global.StTimeStrDecimalC), font1));
            //        }
            //        table2.AddCell(new Phrase(String.Format("{0}", "Average Temperature:"), font2));
            //        table2.AddCell(new Phrase(String.Format("{0:0.00} °C", Global.AvgTemp), font1));
            //        if (Global.model == "PDL-K03")
            //        {
            //            table2.AddCell(new Phrase(String.Format("{0}", "Minimum Humidity:"), font2));
            //            table2.AddCell(new Phrase(String.Format("{0:0.0} %RH", Global.MinHumi), font1));
            //        }
            //        table2.AddCell(new Phrase(String.Format("{0}", "MKT:"), font2));
            //        table2.AddCell(new Phrase(String.Format("{0:0.00} °C", Global.MKTValue), font1));
            //        if (Global.model == "PDL-K01")
            //        {
            //            table2.AddCell(new Phrase(String.Format(" "), font2));
            //            table2.AddCell(new Phrase(String.Format(" "), font1));
            //        }
            //        if (Global.model == "PDL-K03")
            //        {
            //            table2.AddCell(new Phrase(String.Format("{0}", "Maximum Humidity:"), font2));
            //            table2.AddCell(new Phrase(String.Format("{0:0.0} %RH", Global.MaxHumi), font1));
            //            table2.AddCell(new Phrase(String.Format("{0}", "Average Humidity:"), font2));
            //            table2.AddCell(new Phrase(String.Format("{0:0.00} %RH", Global.AvgHumi), font1));
            //        }
            //        document.Add(table2);
            //        // ------------------------Charts------------------------------------------------------------------------------------------------------
            //        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        if (Global.model == "PDL-K03")
            //        {
            //            document.Add(new iTextSharp.text.Paragraph("Temperature & Humidity Chart", font4));
            //        }
            //        if (Global.model == "PDL-K01")
            //        {
            //            document.Add(new iTextSharp.text.Paragraph("Temperature Versus Date & Time Graph", font4));
            //        }
            //        System.Drawing.Image img = zedGraphControl.GraphPane.GetImage(530, 330, 1);
            //        iTextSharp.text.Image j = iTextSharp.text.Image.GetInstance(img, iTextSharp.text.Color.GREEN);
            //        document.Add(j);
            //        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        if (Global.model == "PDL-K01")
            //        {
            //            document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        }
            //        // ------------------------ Data Table------------------------------------------------------------------------------------------------------
            //        PdfPTable table3 = null;
            //        PdfPCell cell = null;
            //        if (Global.model == "PDL-K03")
            //        {
            //            float[] acolumnDefinitionSize = { 2F, 4F, 3F, 3F, 2F, 2F };
            //            table3 = new PdfPTable(acolumnDefinitionSize);
            //        }
            //        if (Global.model == "PDL-K01")
            //        {
            //            float[] acolumnDefinitionSize = { 2F, 4F, 3F, 3F };
            //            table3 = new PdfPTable(acolumnDefinitionSize);
            //        }
            //        table3.WidthPercentage = 90;
            //        if (Global.model == "PDL-K03")
            //        {
            //            cell = new PdfPCell(new Phrase("Sr No." + " Date & Time" + " Temperature" + " Temperature Alarm" + "Humidity" + "Humidity Alarm"));
            //        }
            //        if (Global.model == "PDL-K01")
            //        {
            //            cell = new PdfPCell(new Phrase("Sr No." + " Date & Time" + " Temperature" + " Temperature Alarm"));
            //        }
            //        table3.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            //        cell.BackgroundColor = new iTextSharp.text.Color(0xC3, 0xC2, 0xC1, 0xC2);
            //        table3.AddCell(new Phrase(String.Format("{0}", "Sr No."), font2));
            //        table3.AddCell(new Phrase(String.Format("{0}", "Date and Time"), font2));
            //        table3.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font2));
            //        table3.AddCell(new Phrase(String.Format("{0}", "Temperature Alarm"), font2));
            //        if (Global.model == "PDL-K03")
            //        {
            //            table3.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font2));
            //            table3.AddCell(new Phrase(String.Format("{0}", "Humidity Alarm"), font2));
            //        }
            //        for (int i = 0; i < dataGridView.Rows.Count; i++)
            //        {
            //            table3.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));
            //            table3.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[1].Value.ToString()), font1));
            //            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimal))
            //            {
            //                table3.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
            //            }
            //            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimal))
            //            {
            //                table3.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
            //            }
            //            else
            //            {
            //                table3.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
            //            }
            //            table3.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[3].Value.ToString()), font1));
            //            if (Global.model == "PDL-K03")
            //            {
            //                if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimal))
            //                {
            //                    table3.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
            //                }
            //                else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimal))
            //                {
            //                    table3.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
            //                }
            //                else
            //                {
            //                    table3.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
            //                }
            //                table3.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[5].Value.ToString()), font1));
            //            }
            //        }
            //        document.Add(table3);
            //        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
            //        document.Close();
            //        writer.Close();
            //        fs.Close();
            //        MessageBox.Show("File Saved Successfully", "File Saved", MessageBoxButtons.OK);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    if (ex.Message != null)
            //    {
            //        fs.Dispose();
            //    }
            //}
            //finally
            //{
            //}
            #endregion

            #region New PDF report

            System.IO.FileStream fs = null;
            try
            {
                local.Toatalpagecount = 0;
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "PDF Documents (*.pdf)|*.pdf";
                sfd.FileName = FileName;
                Global.pdfFileName = FileName;
                Global.printPDFName = sfd.FileName;

                if (sfd.ShowDialog() == DialogResult.OK)
                {

                    #region Dummy

                    try
                    {
                        if (!Directory.Exists(System.IO.Path.GetTempPath() + @"enviLOG Basic" + @"\PDFDATA"))
                        {
                            Directory.CreateDirectory(System.IO.Path.GetTempPath() + @"enviLOG Basic" + @"\PDFDATA");
                        }
                    }
                    catch { }

                    fs = new FileStream(System.IO.Path.GetTempPath() + @"enviLOG Basic" + @"\PDFDATA" + @"\" + Path.GetFileName(Global.printPDFName), FileMode.Create);

                    iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30, 30, 30, 30);
                    iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, fs);
                    writer.PageEvent = new MyHeaderFooterEvent();
                    document.SetPageSize(iTextSharp.text.PageSize.A4);

                    #region Header - Footer
                    //------------------------Adding Company Name & Location----------------------------------------------------------------------------------------------------
                    HeaderFooter header;
                    // string sysFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
                    if (Global.radioButtonValue == "Tabular" || Global.radioButtonValue == "Tabular & Chart")
                    {
                        header = new HeaderFooter(new Phrase(
                            System.Configuration.ConfigurationManager.AppSettings["defaultCompanyName"] + "\n" +
                            System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLoc"] + "\n\n" +
                            "Data Report" + "\n" + "From Date & Time :" + Global.fromDateTimeReport + "                To Date & Time :"
                            + Global.toDateTimeReport + "                                                                                                         Print Date & Time :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm")), false);

                        header.Border = iTextSharp.text.Rectangle.NO_BORDER;
                        header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        document.Header = header;
                    }
                    if (Global.radioButtonValue == "Chart" || Global.radioButtonValue == "Tabular & Chart")
                    {
                        header = new HeaderFooter(new Phrase(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyName"] + "\n" + System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLoc"] + "\n\n" + "Chart Report" + "\n" + "From Date & Time :" + Global.fromDateTimeReport + "              To Date & Time :" + Global.toDateTimeReport + "                                                                                                         Print Date & Time :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm")), false);
                        header.Border = iTextSharp.text.Rectangle.NO_BORDER;
                        header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        document.Header = header;
                    }

                    string d = "" + local.Toatalpagecount.ToString() + " Current Page : ";
                    HeaderFooter footer = new HeaderFooter(new Phrase(string.Format(" " + PDFWithLogo.PreparedAndPrintedBy + " :{0}                          Reviewed By :{1}                       Total Pages :{2} ", "", "", " " + d + "")), true);
                    footer.Border = iTextSharp.text.Rectangle.NO_BORDER; ;
                    document.Footer = footer;

                    #endregion

                    #region LOGO
                    // ------------------------ Adding Logo in First Page------------------------------------------------------------------------------------------------------

                    System.Drawing.Image image;
                    if (System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"] == "")
                    {
                        image = System.Drawing.Image.FromFile(Application.StartupPath + "\\enviro_logo.jpg");
                    }
                    else
                    {
                        image = System.Drawing.Image.FromFile(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"]);
                    }
                    Document doc = new Document(PageSize.A4);
                    document.Open();
                    iTextSharp.text.Image pdfImage = iTextSharp.text.Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Png);
                    pdfImage.ScaleToFit(100, 50);
                    pdfImage.SetAbsolutePosition(50, 800);
                    document.Add(pdfImage);

                    #endregion

                    // Add a simple and wellknown phrase to the document in a flow layout manner
                    iTextSharp.text.Font font5 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 12, 0);
                    iTextSharp.text.Font font4 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 13, 1);
                    iTextSharp.text.Font font3 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 11, 1);
                    iTextSharp.text.Font font2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1);
                    iTextSharp.text.Font font1 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 0);
                    iTextSharp.text.Font Redfont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.RED);
                    iTextSharp.text.Font Bluefont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.BLUE);

                    PdfContentByte cb = writer.DirectContent;

                    #region MIN,MAX,AVG
                    //------------------------Table 2------------------------------------------------------------------------------------------------------

                    PdfPTable table2 = null;
                    document.Add(new iTextSharp.text.Paragraph("Device Information", font4));
                    cb.MoveTo(30, document.Top - 103);
                    cb.Stroke();
                    document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                    #region Column Size

                    if (Global.filePathLOGList.Count == 1)
                    {
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 2)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 3)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 4)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                            Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                            Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                    }

                    #endregion

                    table2.WidthPercentage = 100;
                    table2.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

                    table2.AddCell(new Phrase(string.Format("{0}", "Device Name / Serial No"), font3));

                    #region Device Name / Serial No

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        }

                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3]), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3]), font1));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Operating Range"), font3));

                    #region Operating Range

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[3]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[3]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[3]).ToString("0.0") + " %RH"), font1));
                        }
                    }


                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Transaction Count"), font3));

                    #region Trascation Count

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        }
                    }

                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        }
                    }

                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount4), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount4), font1));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(string.Format("{0}", " "), font1));

                    #region Unit

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Minimum"), font3));

                    #region Minimum

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi3), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi3), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp4), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi4), font1));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Maximum"), font3));

                    #region Maximum

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi3), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi3), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp4), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi4), font1));
                        }
                    }


                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Average"), font3));

                    #region Average

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp3.ToString("0.00")), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi3.ToString("0.00")), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp3.ToString("0.00")), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi3.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp4.ToString("0.00")), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi4.ToString("0.00")), font1));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "MKT"), font3));

                    #region MKT

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue4), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }
                    }

                    #endregion

                    document.Add(table2);
                    document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                    #endregion

                    #region Chart
                    // ------------------------Charts------------------------------------------------------------------------------------------------------
                    if (Global.radioButtonValue == "Chart" || Global.radioButtonValue == "Tabular & Chart")
                    {
                        document.Add(new iTextSharp.text.Paragraph("Chart Report", font4));

                        System.Drawing.Image img = zedGraphControl.GraphPane.GetImage(534, 330, 1);
                        iTextSharp.text.Image j = iTextSharp.text.Image.GetInstance(img, iTextSharp.text.Color.GREEN);
                        document.Add(j);

                        if (Global.radioButtonValue == "Tabular & Chart")
                        {
                            document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                            document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                            document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                        }
                        else if (Global.radioButtonValue == "Chart")
                        {

                        }

                    }

                    #endregion

                    #region Data Table
                    // ------------------------ Data Table------------------------------------------------------------------------------------------------------
                    if (Global.radioButtonValue == "Tabular" || Global.radioButtonValue == "Tabular & Chart")
                    {
                        PdfPTable table1 = null;
                        PdfPCell cell = null;

                        document.Add(new iTextSharp.text.Paragraph("Device Report", font4));
                        cb.MoveTo(30, document.Top - 103);
                        cb.Stroke();
                        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                        #region Column Size

                        if (Global.filePathLOGList.Count == 1)
                        {
                            if (Global.modelValue1 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                        }
                        else if (Global.filePathLOGList.Count == 2)
                        {
                            if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                        }
                        else if (Global.filePathLOGList.Count == 3)
                        {
                            if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                        }
                        else if (Global.filePathLOGList.Count == 4)
                        {
                            if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                            if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                            if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                            if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                                Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                                Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                            if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                               Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                               Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                               Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                        }

                        #endregion

                        table1.WidthPercentage = 100;
                        table1.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

                        for (int j = 0; j < dataGridView.Columns.Count - 0; j++)
                        {
                            cell = new PdfPCell(new Phrase(dataGridView.Columns[j].HeaderText, font3));
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            table1.AddCell(cell);
                        }


                        #region DataGirdView Device 1

                        try
                        {
                            if (Global.filePathLOGList.Count == 1)
                            {
                                for (int i = 0; i < Global.TCount1; i++)
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                    }
                                }
                            }
                        }
                        catch (ArgumentOutOfRangeException e)
                        {
                        }
                        catch (Exception ex)
                        {
                            FileLog.ErrorLog(ex.Message + ex.StackTrace);
                        }

                        #endregion

                        #region DataGridView Device 2

                        try
                        {
                            if (Global.filePathLOGList.Count == 2)
                            {
                                #region 1

                                for (int i = 0; i < Global.TCount1; i++)
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }
                                    }

                                #endregion

                                    #region 2

                                    if (Global.modelValue1 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                }
                                    #endregion
                            }
                        }
                        catch (ArgumentOutOfRangeException e)
                        {
                        }
                        catch (Exception ex)
                        {
                            FileLog.ErrorLog(ex.Message + ex.StackTrace);
                        }

                        #endregion

                        #region DataGridView Device 3

                        try
                        {
                            if (Global.filePathLOGList.Count == 3)
                            {
                                #region 1

                                for (int i = 0; i < Global.TCount1; i++)
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                    }

                                #endregion

                                    #region 2

                                    if (Global.modelValue1 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion

                                    #region 3

                                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion
                                }
                            }
                        }
                        catch (ArgumentOutOfRangeException e)
                        {
                        }
                        catch (Exception ex)
                        {
                            FileLog.ErrorLog(ex.Message + ex.StackTrace);
                        }
                        #endregion

                        #region DataGridView Device 4

                        try
                        {
                            if (Global.filePathLOGList.Count == 4)
                            {
                                #region 1

                                for (int i = 0; i < Global.TCount1; i++)
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                    }

                                #endregion

                                    #region 2

                                    if (Global.modelValue1 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion

                                    #region 3

                                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion

                                    #region 4

                                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[8].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[8].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion
                                }
                            }
                        }
                        catch (ArgumentOutOfRangeException e)
                        {
                        }
                        catch (Exception ex)
                        {
                            FileLog.ErrorLog(ex.Message + ex.StackTrace);
                        }

                        #endregion

                        document.Add(table1);
                        //  document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    }

                    #endregion

                    document.Close();
                    writer.Close();
                    fs.Close();

                    #endregion

                    #region Real Pdf

                    Global.printPDFName = sfd.FileName;

                    fs = new FileStream(System.IO.Path.GetTempPath() + @"enviLOG Basic" + @"\PDFDATA" + @"\" + Global.pdfFileName + "_New.pdf", FileMode.Create);

                    document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30, 30, 30, 30);

                    writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, fs);
                    writer.PageEvent = new MyHeaderFooterEvent();
                    document.SetPageSize(iTextSharp.text.PageSize.A4);

                    #region Header - Footer
                    //------------------------Adding Company Name & Location----------------------------------------------------------------------------------------------------

                    if (Global.radioButtonValue == "Tabular" || Global.radioButtonValue == "Tabular & Chart")
                    {
                        header = new HeaderFooter(new Phrase(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyName"] + "\n" + System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLoc"] + "\n\n" + "Data Report" + "\n" + "From Date & Time :" + Global.fromDateTimeReport + "                                       To Date & Time :" + Global.toDateTimeReport + "                                                                                                         Print Date & Time :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm")), false);
                        // image = System.Drawing.Image.FromFile(Application.StartupPath + "\\enviro_logo.jpg");
                        // document.Open();

                        //image = System.Drawing.Image.FromFile(Application.StartupPath + "\\enviro_logo.jpg");
                        //iTextSharp.text.Image pdfImage = iTextSharp.text.Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Jpeg);
                        //pdfImage.ScaleToFit(100, 50);
                        //pdfImage.SetAbsolutePosition(50, 800);
                        //header.Chunks.Add(new Chunk(pdfImage, 100, 50, true));

                        header.Border = iTextSharp.text.Rectangle.NO_BORDER;
                        header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        document.Header = header;
                    }
                    if (Global.radioButtonValue == "Chart")
                    {
                        header = new HeaderFooter(new Phrase(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyName"] + "\n" + System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLoc"] + "\n\n" + "Chart Report" + "\n" + "From Date & Time :" + Global.fromDateTimeReport + "                                       To Date & Time :" + Global.toDateTimeReport + "                                                                                                         Print Date & Time :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm")), false);
                        header.Border = iTextSharp.text.Rectangle.NO_BORDER;
                        header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                        document.Header = header;
                    }

                    d = "" + local.Toatalpagecount.ToString() + " Current Page : ";
                    footer = new HeaderFooter(new Phrase(string.Format(" " + PDFWithLogo.PreparedAndPrintedBy + " :{0}                          Reviewed By :{1}                       Total Pages :{2} ", "", "", " " + d + "")), true);
                    footer.Border = iTextSharp.text.Rectangle.NO_BORDER; ;
                    document.Footer = footer;

                    #endregion

                    #region LOGO
                    // ------------------------ Adding Logo in First Page------------------------------------------------------------------------------------------------------


                    if (System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"] == "")
                    {
                        image = System.Drawing.Image.FromFile(Application.StartupPath + "\\enviro_logo.jpg");
                    }
                    else
                    {
                        image = System.Drawing.Image.FromFile(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"]);
                    }
                    doc = new Document(PageSize.A4);
                    document.Open();
                    //pdfImage = iTextSharp.text.Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Png);
                    //pdfImage.ScaleToFit(100, 50);
                    //pdfImage.SetAbsolutePosition(50, 800);
                    //document.Add(pdfImage);

                    #endregion

                    // Add a simple and wellknown phrase to the document in a flow layout manner
                    font5 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 12, 0);
                    font4 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 13, 1);
                    font3 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 11, 1);
                    font2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1);
                    font1 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 0);
                    Redfont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.RED);
                    Bluefont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.BLUE);

                    cb = writer.DirectContent;

                    #region MIN,MAX,AVG
                    //------------------------Table 2------------------------------------------------------------------------------------------------------

                    table2 = null;
                    document.Add(new iTextSharp.text.Paragraph("Device Information", font4));
                    cb.MoveTo(30, document.Top - 103);
                    cb.Stroke();
                    document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                    #region Column Size

                    if (Global.filePathLOGList.Count == 1)
                    {
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 2)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 3)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 4)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                            Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                            Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03")
                        {
                            float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F };
                            table2 = new PdfPTable(columnDefinitionSize2);
                        }

                    }

                    #endregion

                    table2.WidthPercentage = 100;
                    table2.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

                    table2.AddCell(new Phrase(string.Format("{0}", "Device Name / Serial No"), font3));

                    #region Device Name / Serial No

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        }

                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3]), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3]), font1));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Operating Range"), font3));

                    #region Operating Range

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[3]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[3]).ToString("0.0") + " °C", font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[3]).ToString("0.0") + " %RH"), font1));
                        }
                    }


                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Transaction Count"), font3));

                    #region Trascation Count

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        }
                    }

                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        }
                    }

                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount4), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.TCount4), font1));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(string.Format("{0}", " "), font1));

                    #region Unit

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Minimum"), font3));

                    #region Minimum

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi3), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi3), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp4), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi4), font1));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Maximum"), font3));

                    #region Maximum

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi3), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi3), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp4), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi4), font1));
                        }
                    }


                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "Average"), font3));

                    #region Average

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp3.ToString("0.00")), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi3.ToString("0.00")), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp3.ToString("0.00")), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi3.ToString("0.00")), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp4.ToString("0.00")), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi4.ToString("0.00")), font1));
                        }
                    }

                    #endregion

                    table2.AddCell(new Phrase(String.Format("{0}", "MKT"), font3));

                    #region MKT

                    if (Global.filePathLOGList.Count == 1)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 2)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 3)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }
                    }
                    if (Global.filePathLOGList.Count == 4)
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                        if (Global.modelValue2 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue3), font1));
                        if (Global.modelValue3 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }

                        table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue4), font1));
                        if (Global.modelValue4 == "PDL-K03")
                        {
                            table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                        }
                    }

                    #endregion

                    document.Add(table2);
                    document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                    #endregion

                    #region Chart
                    // ------------------------Charts------------------------------------------------------------------------------------------------------
                    if (Global.radioButtonValue == "Chart" || Global.radioButtonValue == "Tabular & Chart")
                    {
                        document.Add(new iTextSharp.text.Paragraph("Chart Report", font4));

                        System.Drawing.Image img = zedGraphControl.GraphPane.GetImage(534, 330, 1);
                        iTextSharp.text.Image j = iTextSharp.text.Image.GetInstance(img, iTextSharp.text.Color.GREEN);
                        document.Add(j);

                        if (Global.radioButtonValue == "Tabular & Chart")
                        {
                            document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                            document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                            document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                        }
                        else if (Global.radioButtonValue == "Chart")
                        {

                        }

                    }

                    #endregion

                    #region Data Table
                    // ------------------------ Data Table------------------------------------------------------------------------------------------------------
                    if (Global.radioButtonValue == "Tabular" || Global.radioButtonValue == "Tabular & Chart")
                    {
                        PdfPTable table1 = null;
                        PdfPCell cell = null;

                        document.Add(new iTextSharp.text.Paragraph("Device Report", font4));
                        cb.MoveTo(30, document.Top - 103);
                        cb.Stroke();
                        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                        #region Column Size

                        if (Global.filePathLOGList.Count == 1)
                        {
                            if (Global.modelValue1 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                        }
                        else if (Global.filePathLOGList.Count == 2)
                        {
                            if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                        }
                        else if (Global.filePathLOGList.Count == 3)
                        {
                            if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                            else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }
                        }
                        else if (Global.filePathLOGList.Count == 4)
                        {
                            if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                            if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                            if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                            if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                                Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                                Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                                Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                            if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                               Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                               Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                               Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03")
                            {
                                float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F };
                                table1 = new PdfPTable(columnDefinitionSize1);
                            }

                        }

                        #endregion

                        table1.WidthPercentage = 100;
                        table1.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

                        for (int j = 0; j < dataGridView.Columns.Count - 0; j++)
                        {
                            cell = new PdfPCell(new Phrase(dataGridView.Columns[j].HeaderText, font3));
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            table1.AddCell(cell);
                        }


                        #region DataGirdView Device 1

                        try
                        {
                            if (Global.filePathLOGList.Count == 1)
                            {
                                for (int i = 0; i < Global.TCount1; i++)
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                    }
                                }
                            }
                        }
                        catch (ArgumentOutOfRangeException e)
                        {
                        }
                        catch (Exception ex)
                        {
                            FileLog.ErrorLog(ex.Message + ex.StackTrace);
                        }

                        #endregion

                        #region DataGridView Device 2

                        try
                        {
                            if (Global.filePathLOGList.Count == 2)
                            {
                                #region 1

                                for (int i = 0; i < Global.TCount1; i++)
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }
                                    }

                                #endregion

                                    #region 2

                                    if (Global.modelValue1 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                }
                                    #endregion
                            }
                        }
                        catch (ArgumentOutOfRangeException e)
                        {
                        }
                        catch (Exception ex)
                        {
                            FileLog.ErrorLog(ex.Message + ex.StackTrace);
                        }

                        #endregion

                        #region DataGridView Device 3

                        try
                        {
                            if (Global.filePathLOGList.Count == 3)
                            {
                                #region 1

                                for (int i = 0; i < Global.TCount1; i++)
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                    }

                                #endregion

                                    #region 2

                                    if (Global.modelValue1 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion

                                    #region 3

                                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion
                                }
                            }
                        }
                        catch (ArgumentOutOfRangeException e)
                        {
                        }
                        catch (Exception ex)
                        {
                            FileLog.ErrorLog(ex.Message + ex.StackTrace);
                        }
                        #endregion

                        #region DataGridView Device 4

                        try
                        {
                            if (Global.filePathLOGList.Count == 4)
                            {
                                #region 1

                                for (int i = 0; i < Global.TCount1; i++)
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                    }

                                #endregion

                                    #region 2

                                    if (Global.modelValue1 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue2 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion

                                    #region 3

                                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }
                                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue3 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion

                                    #region 4

                                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[8].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[8].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }

                                        if (Global.modelValue4 == "PDL-K03")
                                        {
                                            if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                            }
                                            else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                            }
                                            else
                                            {
                                                table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                            }
                                        }
                                    }

                                    #endregion
                                }
                            }
                        }
                        catch (ArgumentOutOfRangeException e)
                        {
                        }
                        catch (Exception ex)
                        {
                            FileLog.ErrorLog(ex.Message + ex.StackTrace);
                        }

                        #endregion

                        document.Add(table1);
                        //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    }

                    #endregion

                    document.Close();
                    writer.Close();
                    fs.Close();

                    PDFWithLogo.LogoEveryPage(System.IO.Path.GetTempPath() + @"enviLOG Basic" + @"\PDFDATA" + @"\" + Global.pdfFileName + "_New.pdf", sfd.FileName);

                    #endregion
                    local.Toatalpagecount = 0;
                    MessageBox.Show("File Saved Successfully", "Save", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
            catch (Exception ex)
            {
                if (ex.Message != null)
                {
                    fs.Dispose();
                }
            }
            finally
            {
            }
            #endregion
        }

        /// <summary>
        /// Export the tabuler report to Excel file.Global.serialNoStr
        /// </summary>
        public void ExportExcelFile()
        {
            try
            {
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 2;




                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Excel.Application xlApp;
                    Excel.Workbook xlWorkBook;
                    Excel.Worksheet xlWorkSheet;
                    Excel.Range range;
                    object misValue = System.Reflection.Missing.Value;

                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                    //  xlWorkBook1.Worksheets.Add((DataTable)dataGridView.DataSource);

                    //  xlWorkSheet.Copy(xlWorkSheet.Range["A1:E1"], xlWorkBook1.Worksheets[0].Range["A1:E1"]);

                    int i = 0;
                    int j = 0;



                    #region Device Information
                    xlWorkSheet.Cells[10, 1] = "Device Information";
                    xlWorkSheet.Cells[10, 1].Font.Size = 11;
                    xlWorkSheet.Cells[10, 1].Font.Bold = true;
                    xlWorkSheet.Cells[10, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[10, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    #endregion

                    #region Logo, Company Name,Location

                    if (System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"] == "")
                    {
                        xlWorkSheet.Shapes.AddPicture(Application.StartupPath + "\\enviro_logo.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 10, 10, 100, 45); //C:\\csharp-xl-picture.JPG
                    }
                    else
                    {
                        xlWorkSheet.Shapes.AddPicture(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"], Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 10, 10, 100, 45); //C:\\csharp-xl-picture.JPG                      
                    }

                    range = xlWorkSheet.get_Range("D1:F2");
                    range.Merge();
                    range.Font.Bold = true;
                    range.Font.Size = 13;
                    range.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xlWorkSheet.Cells[1, 4] = System.Configuration.ConfigurationManager.AppSettings["defaultCompanyName"];

                    range = xlWorkSheet.get_Range("D3:F3");
                    range.Merge();
                    range.Font.Bold = true;
                    range.Font.Size = 13;
                    range.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xlWorkSheet.Cells[3, 4] = System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLoc"];

                    range = xlWorkSheet.get_Range("D5:F5");
                    range.Merge();
                    range.Font.Bold = true;
                    range.Font.Size = 13;
                    range.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xlWorkSheet.Cells[5, 4] = System.Configuration.ConfigurationManager.AppSettings["defaultCompanyDataReport"];

                    #endregion

                    #region From & To Datetime,Print Datetime

                    range = xlWorkSheet.get_Range("A7:B7");
                    range.Merge();
                    range.Font.Bold = true;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    xlWorkSheet.Cells[7, 1] = "From Date & Time : " + Global.fromDateTimeReport;

                    range = xlWorkSheet.get_Range("F7:I7");
                    range.Merge();
                    range.Font.Bold = true;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    xlWorkSheet.Cells[7, 6] = "To Date & Time : " + Global.toDateTimeReport;

                    range = xlWorkSheet.get_Range("F8:I8");
                    range.Merge();
                    range.Font.Bold = true;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    xlWorkSheet.Cells[8, 6] = "Print Date & Time : " + DateTime.Now.ToString("dd-MM-yyyy HH:mm");

                    #endregion

                    #region Device name & Serial No

                    xlWorkSheet.Cells[11, 1] = "Device name / Serial No";
                    xlWorkSheet.Cells[11, 1].Font.Size = 10;
                    xlWorkSheet.Cells[11, 1].Font.Bold = true;
                    xlWorkSheet.Cells[11, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[11, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;

                    #region Device 1

                    if (Global.filePathLOGList.Count == 1)
                    {
                        xlWorkSheet.Cells[11, 2] = Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0];

                        xlWorkSheet.Cells[11, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[11, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[11, 3] = Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0];
                            xlWorkSheet.Cells[11, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }
                    }

                    #endregion

                    #region  Device 2

                    if (Global.filePathLOGList.Count == 2)
                    {
                        #region 1

                        xlWorkSheet.Cells[11, 2] = Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0];
                        xlWorkSheet.Cells[11, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[11, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[11, 3] = Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0];
                            xlWorkSheet.Cells[11, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[11, 3] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                            xlWorkSheet.Cells[11, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 4] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                                xlWorkSheet.Cells[11, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {

                            xlWorkSheet.Cells[11, 4] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                            xlWorkSheet.Cells[11, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 5] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                                xlWorkSheet.Cells[11, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 3

                    if (Global.filePathLOGList.Count == 3)
                    {
                        #region 1

                        xlWorkSheet.Cells[11, 2] = Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0];
                        xlWorkSheet.Cells[11, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[11, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[11, 3] = Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0];
                            xlWorkSheet.Cells[11, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[11, 3] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                            xlWorkSheet.Cells[11, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 4] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                                xlWorkSheet.Cells[11, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {

                            xlWorkSheet.Cells[11, 4] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                            xlWorkSheet.Cells[11, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 5] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                                xlWorkSheet.Cells[11, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[11, 4] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                            xlWorkSheet.Cells[11, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 5] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                                xlWorkSheet.Cells[11, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[11, 6] = Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2];
                            xlWorkSheet.Cells[11, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 7] = Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2];
                                xlWorkSheet.Cells[11, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[11, 5] = Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2];
                            xlWorkSheet.Cells[11, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 6] = Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2];
                                xlWorkSheet.Cells[11, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 4

                    if (Global.filePathLOGList.Count == 4)
                    {
                        #region 1

                        xlWorkSheet.Cells[11, 2] = Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0];
                        xlWorkSheet.Cells[11, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[11, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[11, 3] = Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0];
                            xlWorkSheet.Cells[11, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[11, 3] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                            xlWorkSheet.Cells[11, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 4] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                                xlWorkSheet.Cells[11, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[11, 4] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                            xlWorkSheet.Cells[11, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 5] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                                xlWorkSheet.Cells[11, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }


                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[11, 4] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                            xlWorkSheet.Cells[11, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 5] = Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1];
                                xlWorkSheet.Cells[11, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[11, 6] = Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2];
                            xlWorkSheet.Cells[11, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 7] = Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2];
                                xlWorkSheet.Cells[11, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[11, 5] = Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2];
                            xlWorkSheet.Cells[11, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 6] = Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2];
                                xlWorkSheet.Cells[11, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 4

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[11, 5] = Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3];
                            xlWorkSheet.Cells[11, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 6] = Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3];
                                xlWorkSheet.Cells[11, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[11, 8] = Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3];
                            xlWorkSheet.Cells[11, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 9] = Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3];
                                xlWorkSheet.Cells[11, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 9].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[11, 7] = Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3];
                            xlWorkSheet.Cells[11, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 8] = Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3];
                                xlWorkSheet.Cells[11, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[11, 6] = Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3];
                            xlWorkSheet.Cells[11, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[11, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[11, 7] = Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3];
                                xlWorkSheet.Cells[11, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[11, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }
                    #endregion

                    #endregion

                    #region Operating Range

                    xlWorkSheet.Cells[12, 1] = "Operating Range";
                    xlWorkSheet.Cells[12, 1].Font.Size = 10;
                    xlWorkSheet.Cells[12, 1].Font.Bold = true;
                    xlWorkSheet.Cells[12, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[12, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;

                    #region Device 1

                    if (Global.filePathLOGList.Count == 1)
                    {
                        xlWorkSheet.Cells[12, 2] = Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C";
                        xlWorkSheet.Cells[12, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[12, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 3] = Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH";
                            xlWorkSheet.Cells[12, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }
                    }

                    #endregion

                    #region Device 2

                    if (Global.filePathLOGList.Count == 2)
                    {
                        #region 1

                        xlWorkSheet.Cells[12, 2] = Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C";
                        xlWorkSheet.Cells[12, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[12, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 3] = Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH";
                            xlWorkSheet.Cells[12, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[12, 3] = Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 4] = Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 4] = Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 5] = Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 3

                    if (Global.filePathLOGList.Count == 3)
                    {
                        #region 1

                        xlWorkSheet.Cells[12, 2] = Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C";
                        xlWorkSheet.Cells[12, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[12, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 3] = Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH";
                            xlWorkSheet.Cells[12, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[12, 3] = Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 4] = Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 4] = Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 5] = Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[12, 4] = Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 5] = Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 6] = Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 7] = Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[12, 5] = Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 6] = Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 4

                    if (Global.filePathLOGList.Count == 4)
                    {
                        #region 1

                        xlWorkSheet.Cells[12, 2] = Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C";
                        xlWorkSheet.Cells[12, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[12, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 3] = Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH";
                            xlWorkSheet.Cells[12, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[12, 3] = Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 4] = Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 4] = Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 5] = Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[12, 4] = Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 5] = Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 6] = Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 7] = Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[12, 5] = Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 6] = Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 4

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[12, 5] = Convert.ToDouble(Global.tempLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[3]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 6] = Convert.ToDouble(Global.humiLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[3]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 8] = Convert.ToDouble(Global.tempLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[3]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 9] = Convert.ToDouble(Global.humiLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[3]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 9].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[12, 7] = Convert.ToDouble(Global.tempLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[3]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 8] = Convert.ToDouble(Global.humiLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[3]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[12, 6] = Convert.ToDouble(Global.tempLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[3]).ToString("0.0") + " °C";
                            xlWorkSheet.Cells[12, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[12, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[12, 7] = Convert.ToDouble(Global.humiLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[3]).ToString("0.0") + " %RH";
                                xlWorkSheet.Cells[12, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[12, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }
                    #endregion

                    #endregion

                    #region Transaction Count

                    xlWorkSheet.Cells[13, 1] = "Transaction Count";
                    xlWorkSheet.Cells[13, 1].Font.Size = 10;
                    xlWorkSheet.Cells[13, 1].Font.Bold = true;
                    xlWorkSheet.Cells[13, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[13, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;

                    #region Device 1

                    if (Global.filePathLOGList.Count == 1)
                    {
                        xlWorkSheet.Cells[13, 2] = Global.TCount1;
                        xlWorkSheet.Cells[13, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[13, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 3] = Global.TCount1;
                            xlWorkSheet.Cells[13, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }
                    }

                    #endregion

                    #region Device 2

                    if (Global.filePathLOGList.Count == 2)
                    {
                        #region 1

                        xlWorkSheet.Cells[13, 2] = Global.TCount1;
                        xlWorkSheet.Cells[13, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[13, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 3] = Global.TCount1;
                            xlWorkSheet.Cells[13, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[13, 3] = Global.TCount2;
                            xlWorkSheet.Cells[13, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 4] = Global.TCount2;
                                xlWorkSheet.Cells[13, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 4] = Global.TCount2;
                            xlWorkSheet.Cells[13, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 5] = Global.TCount2;
                                xlWorkSheet.Cells[13, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }

                    #endregion

                    #region Device 3

                    if (Global.filePathLOGList.Count == 3)
                    {
                        #region 1

                        xlWorkSheet.Cells[13, 2] = Global.TCount1;
                        xlWorkSheet.Cells[13, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[13, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 3] = Global.TCount1;
                            xlWorkSheet.Cells[13, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[13, 3] = Global.TCount2;
                            xlWorkSheet.Cells[13, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 4] = Global.TCount2;
                                xlWorkSheet.Cells[13, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 4] = Global.TCount2;
                            xlWorkSheet.Cells[13, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 5] = Global.TCount2;
                                xlWorkSheet.Cells[13, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[13, 4] = Global.TCount3;
                            xlWorkSheet.Cells[13, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 5] = Global.TCount3;
                                xlWorkSheet.Cells[13, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 6] = Global.TCount3;
                            xlWorkSheet.Cells[13, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 7] = Global.TCount3;
                                xlWorkSheet.Cells[13, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[13, 5] = Global.TCount3;
                            xlWorkSheet.Cells[13, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 6] = Global.TCount3;
                                xlWorkSheet.Cells[13, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 4

                    if (Global.filePathLOGList.Count == 4)
                    {
                        #region 1

                        xlWorkSheet.Cells[13, 2] = Global.TCount1;
                        xlWorkSheet.Cells[13, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[13, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 3] = Global.TCount1;
                            xlWorkSheet.Cells[13, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[13, 3] = Global.TCount2;
                            xlWorkSheet.Cells[13, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 4] = Global.TCount2;
                                xlWorkSheet.Cells[13, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 4] = Global.TCount2;
                            xlWorkSheet.Cells[13, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 5] = Global.TCount2;
                                xlWorkSheet.Cells[13, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[13, 4] = Global.TCount3;
                            xlWorkSheet.Cells[13, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 5] = Global.TCount3;
                                xlWorkSheet.Cells[13, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 6] = Global.TCount3;
                            xlWorkSheet.Cells[13, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 7] = Global.TCount3;
                                xlWorkSheet.Cells[13, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[13, 5] = Global.TCount3;
                            xlWorkSheet.Cells[13, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 6] = Global.TCount3;
                                xlWorkSheet.Cells[13, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 4

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[13, 5] = Global.TCount4;
                            xlWorkSheet.Cells[13, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 6] = Global.TCount4;
                                xlWorkSheet.Cells[13, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 8] = Global.TCount4;
                            xlWorkSheet.Cells[13, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 9] = Global.TCount4;
                                xlWorkSheet.Cells[13, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 9].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[13, 7] = Global.TCount4;
                            xlWorkSheet.Cells[13, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 8] = Global.TCount4;
                                xlWorkSheet.Cells[13, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[13, 6] = Global.TCount4;
                            xlWorkSheet.Cells[13, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[13, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[13, 7] = Global.TCount4;
                                xlWorkSheet.Cells[13, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[13, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }
                    #endregion

                    #endregion

                    #region Temperature (°C) and Humidity (%RH)

                    xlWorkSheet.Cells[14, 1] = "";
                    xlWorkSheet.Cells[14, 1].Font.Size = 10;
                    xlWorkSheet.Cells[14, 1].Font.Bold = true;
                    xlWorkSheet.Cells[14, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[14, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;

                    #region Device 1

                    if (Global.filePathLOGList.Count == 1)
                    {
                        xlWorkSheet.Cells[14, 2] = "Temperature °C";
                        xlWorkSheet.Cells[14, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[14, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        xlWorkSheet.Cells[14, 2].Font.Bold = true;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 3] = "Humidity %RH";
                            xlWorkSheet.Cells[14, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 3].Font.Bold = true;
                        }
                    }

                    #endregion

                    #region Device 2

                    if (Global.filePathLOGList.Count == 2)
                    {
                        #region 1

                        xlWorkSheet.Cells[14, 2] = "Temperature °C";
                        xlWorkSheet.Cells[14, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[14, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        xlWorkSheet.Cells[14, 2].Font.Bold = true;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 3] = "Humidity %RH";
                            xlWorkSheet.Cells[14, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 3].Font.Bold = true;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[14, 3] = "Temperature °C";
                            xlWorkSheet.Cells[14, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 3].Font.Bold = true;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 4] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 4].Font.Bold = true;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 4] = "Temperature °C";
                            xlWorkSheet.Cells[14, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 4].Font.Bold = true;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 5] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 5].Font.Bold = true;
                            }
                        }

                        #endregion

                    }

                    #endregion

                    #region Device 3

                    if (Global.filePathLOGList.Count == 3)
                    {
                        #region 1

                        xlWorkSheet.Cells[14, 2] = "Temperature °C";
                        xlWorkSheet.Cells[14, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[14, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        xlWorkSheet.Cells[14, 2].Font.Bold = true;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 3] = "Humidity %RH";
                            xlWorkSheet.Cells[14, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 3].Font.Bold = true;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[14, 3] = "Temperature °C";
                            xlWorkSheet.Cells[14, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 3].Font.Bold = true;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 4] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 4].Font.Bold = true;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 4] = "Temperature °C";
                            xlWorkSheet.Cells[14, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 4].Font.Bold = true;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 5] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 5].Font.Bold = true;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[14, 4] = "Temperature °C";
                            xlWorkSheet.Cells[14, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 4].Font.Bold = true;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 5] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 5].Font.Bold = true;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 6] = "Temperature °C";
                            xlWorkSheet.Cells[14, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 6].Font.Bold = true;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 7] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 7].Font.Bold = true;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[14, 5] = "Temperature °C";
                            xlWorkSheet.Cells[14, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 5].Font.Bold = true;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 6] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 6].Font.Bold = true;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 4

                    if (Global.filePathLOGList.Count == 4)
                    {
                        #region 1

                        xlWorkSheet.Cells[14, 2] = "Temperature °C";
                        xlWorkSheet.Cells[14, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[14, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        xlWorkSheet.Cells[14, 2].Font.Bold = true;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 3] = "Humidity %RH";
                            xlWorkSheet.Cells[14, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 3].Font.Bold = true;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[14, 3] = "Temperature °C";
                            xlWorkSheet.Cells[14, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 3].Font.Bold = true;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 4] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 4].Font.Bold = true;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 4] = "Temperature °C";
                            xlWorkSheet.Cells[14, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 4].Font.Bold = true;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 5] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 5].Font.Bold = true;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[14, 4] = "Temperature °C";
                            xlWorkSheet.Cells[14, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 4].Font.Bold = true;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 5] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 5].Font.Bold = true;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 6] = "Temperature °C";
                            xlWorkSheet.Cells[14, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 6].Font.Bold = true;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 7] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 7].Font.Bold = true;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[14, 5] = "Temperature °C";
                            xlWorkSheet.Cells[14, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 5].Font.Bold = true;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 6] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 6].Font.Bold = true;
                            }
                        }

                        #endregion

                        #region 4

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[14, 5] = "Temperature °C";
                            xlWorkSheet.Cells[14, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 5].Font.Bold = true;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 6] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 6].Font.Bold = true;
                            }
                        }

                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 8] = "Temperature °C";
                            xlWorkSheet.Cells[14, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 8].Font.Bold = true;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 9] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 9].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 9].Font.Bold = true;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[14, 7] = "Temperature °C";
                            xlWorkSheet.Cells[14, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 7].Font.Bold = true;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 8] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 8].Font.Bold = true;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[14, 6] = "Temperature °C";
                            xlWorkSheet.Cells[14, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[14, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            xlWorkSheet.Cells[14, 6].Font.Bold = true;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[14, 7] = "Humidity %RH";
                                xlWorkSheet.Cells[14, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[14, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                                xlWorkSheet.Cells[14, 7].Font.Bold = true;
                            }
                        }

                        #endregion

                    }
                    #endregion

                    #endregion

                    #region Min

                    xlWorkSheet.Cells[15, 1] = "Minimum";
                    xlWorkSheet.Cells[15, 1].Font.Size = 10;
                    xlWorkSheet.Cells[15, 1].Font.Bold = true;
                    xlWorkSheet.Cells[15, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[15, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;

                    #region Device 1

                    if (Global.filePathLOGList.Count == 1)
                    {
                        xlWorkSheet.Cells[15, 2] = Global.MinTemp1;
                        xlWorkSheet.Cells[15, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[15, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 3] = Global.MinHumi1;
                            xlWorkSheet.Cells[15, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }
                    }

                    #endregion

                    #region Device 2

                    if (Global.filePathLOGList.Count == 2)
                    {
                        #region 1

                        xlWorkSheet.Cells[15, 2] = Global.MinTemp1;
                        xlWorkSheet.Cells[15, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[15, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 3] = Global.MinHumi1;
                            xlWorkSheet.Cells[15, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[15, 3] = Global.MinTemp2;
                            xlWorkSheet.Cells[15, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 4] = Global.MinHumi2;
                                xlWorkSheet.Cells[15, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 4] = Global.MinTemp2;
                            xlWorkSheet.Cells[15, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 5] = Global.MinHumi2;
                                xlWorkSheet.Cells[15, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }

                    #endregion

                    #region Device 3

                    if (Global.filePathLOGList.Count == 3)
                    {
                        #region 1

                        xlWorkSheet.Cells[15, 2] = Global.MinTemp1;
                        xlWorkSheet.Cells[15, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[15, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 3] = Global.MinHumi1;
                            xlWorkSheet.Cells[15, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[15, 3] = Global.MinTemp2;
                            xlWorkSheet.Cells[15, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 4] = Global.MinHumi2;
                                xlWorkSheet.Cells[15, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 4] = Global.MinTemp2;
                            xlWorkSheet.Cells[15, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 5] = Global.MinHumi2;
                                xlWorkSheet.Cells[15, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[15, 4] = Global.MinTemp3;
                            xlWorkSheet.Cells[15, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 5] = Global.MinHumi3;
                                xlWorkSheet.Cells[15, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 6] = Global.MinTemp3;
                            xlWorkSheet.Cells[15, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 7] = Global.MinHumi3;
                                xlWorkSheet.Cells[15, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[15, 5] = Global.MinTemp3;
                            xlWorkSheet.Cells[15, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 6] = Global.MinHumi3;
                                xlWorkSheet.Cells[15, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 4

                    if (Global.filePathLOGList.Count == 4)
                    {
                        #region 1

                        xlWorkSheet.Cells[15, 2] = Global.MinTemp1;
                        xlWorkSheet.Cells[15, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[15, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 3] = Global.MinHumi1;
                            xlWorkSheet.Cells[15, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[15, 3] = Global.MinTemp2;
                            xlWorkSheet.Cells[15, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 4] = Global.MinHumi2;
                                xlWorkSheet.Cells[15, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 4] = Global.MinTemp2;
                            xlWorkSheet.Cells[15, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 5] = Global.MinHumi2;
                                xlWorkSheet.Cells[15, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[15, 4] = Global.MinTemp3;
                            xlWorkSheet.Cells[15, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 5] = Global.MinHumi3;
                                xlWorkSheet.Cells[15, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 6] = Global.MinTemp3;
                            xlWorkSheet.Cells[15, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 7] = Global.MinHumi3;
                                xlWorkSheet.Cells[15, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[15, 5] = Global.MinTemp3;
                            xlWorkSheet.Cells[15, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 6] = Global.MinHumi3;
                                xlWorkSheet.Cells[15, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 4

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[15, 5] = Global.MinTemp4;
                            xlWorkSheet.Cells[15, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 6] = Global.MinHumi4;
                                xlWorkSheet.Cells[15, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 8] = Global.MinTemp4;
                            xlWorkSheet.Cells[15, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 9] = Global.MinHumi4;
                                xlWorkSheet.Cells[15, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 9].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[15, 7] = Global.MinTemp4;
                            xlWorkSheet.Cells[15, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 8] = Global.MinHumi4;
                                xlWorkSheet.Cells[15, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[15, 6] = Global.MinTemp4;
                            xlWorkSheet.Cells[15, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[15, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[15, 7] = Global.MinHumi4;
                                xlWorkSheet.Cells[15, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[15, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }
                    #endregion

                    #endregion

                    #region Max

                    xlWorkSheet.Cells[16, 1] = "Maximum";
                    xlWorkSheet.Cells[16, 1].Font.Size = 10;
                    xlWorkSheet.Cells[16, 1].Font.Bold = true;
                    xlWorkSheet.Cells[16, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[16, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;

                    #region Device 1

                    if (Global.filePathLOGList.Count == 1)
                    {
                        xlWorkSheet.Cells[16, 2] = Global.MaxTemp1;
                        xlWorkSheet.Cells[16, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[16, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 3] = Global.MaxHumi1;
                            xlWorkSheet.Cells[16, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }
                    }

                    #endregion

                    #region Device 2

                    if (Global.filePathLOGList.Count == 2)
                    {
                        #region 1

                        xlWorkSheet.Cells[16, 2] = Global.MaxTemp1;
                        xlWorkSheet.Cells[16, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[16, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 3] = Global.MaxHumi1;
                            xlWorkSheet.Cells[16, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[16, 3] = Global.MaxTemp2;
                            xlWorkSheet.Cells[16, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 4] = Global.MaxHumi2;
                                xlWorkSheet.Cells[16, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 4] = Global.MaxTemp2;
                            xlWorkSheet.Cells[16, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 5] = Global.MaxHumi2;
                                xlWorkSheet.Cells[16, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }

                    #endregion

                    #region Device 3

                    if (Global.filePathLOGList.Count == 3)
                    {
                        #region 1

                        xlWorkSheet.Cells[16, 2] = Global.MaxTemp1;
                        xlWorkSheet.Cells[16, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[16, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 3] = Global.MaxHumi1;
                            xlWorkSheet.Cells[16, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[16, 3] = Global.MaxTemp2;
                            xlWorkSheet.Cells[16, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 4] = Global.MaxHumi2;
                                xlWorkSheet.Cells[16, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 4] = Global.MaxTemp2;
                            xlWorkSheet.Cells[16, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 5] = Global.MaxHumi2;
                                xlWorkSheet.Cells[16, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[16, 4] = Global.MaxTemp3;
                            xlWorkSheet.Cells[16, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 5] = Global.MaxHumi3;
                                xlWorkSheet.Cells[16, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 6] = Global.MaxTemp3;
                            xlWorkSheet.Cells[16, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 7] = Global.MaxHumi3;
                                xlWorkSheet.Cells[16, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[16, 5] = Global.MaxTemp3;
                            xlWorkSheet.Cells[16, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 6] = Global.MaxHumi3;
                                xlWorkSheet.Cells[16, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 4

                    if (Global.filePathLOGList.Count == 4)
                    {
                        #region 1

                        xlWorkSheet.Cells[16, 2] = Global.MaxTemp1;
                        xlWorkSheet.Cells[16, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[16, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 3] = Global.MaxHumi1;
                            xlWorkSheet.Cells[16, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[16, 3] = Global.MaxTemp2;
                            xlWorkSheet.Cells[16, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 4] = Global.MaxHumi2;
                                xlWorkSheet.Cells[16, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 4] = Global.MaxTemp2;
                            xlWorkSheet.Cells[16, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 5] = Global.MaxHumi2;
                                xlWorkSheet.Cells[16, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[16, 4] = Global.MaxTemp3;
                            xlWorkSheet.Cells[16, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 5] = Global.MaxHumi3;
                                xlWorkSheet.Cells[16, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 6] = Global.MaxTemp3;
                            xlWorkSheet.Cells[16, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 7] = Global.MaxHumi3;
                                xlWorkSheet.Cells[16, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[16, 5] = Global.MaxTemp3;
                            xlWorkSheet.Cells[16, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 6] = Global.MaxHumi3;
                                xlWorkSheet.Cells[16, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 4

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[16, 5] = Global.MaxTemp4;
                            xlWorkSheet.Cells[16, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 6] = Global.MaxHumi4;
                                xlWorkSheet.Cells[16, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 8] = Global.MaxTemp4;
                            xlWorkSheet.Cells[16, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 9] = Global.MaxHumi4;
                                xlWorkSheet.Cells[16, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 9].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[16, 7] = Global.MaxTemp4;
                            xlWorkSheet.Cells[16, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 8] = Global.MaxHumi4;
                                xlWorkSheet.Cells[16, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[16, 6] = Global.MaxTemp4;
                            xlWorkSheet.Cells[16, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[16, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[16, 7] = Global.MaxHumi4;
                                xlWorkSheet.Cells[16, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[16, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }
                    #endregion

                    #endregion

                    #region Average

                    xlWorkSheet.Cells[17, 1] = "Average";
                    xlWorkSheet.Cells[17, 1].Font.Size = 10;
                    xlWorkSheet.Cells[17, 1].Font.Bold = true;
                    xlWorkSheet.Cells[17, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[17, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;

                    #region Device 1

                    if (Global.filePathLOGList.Count == 1)
                    {
                        xlWorkSheet.Cells[17, 2] = Math.Round(Global.AvgTemp1, 2);
                        xlWorkSheet.Cells[17, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[17, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 3] = Math.Round(Global.AvgHumi1, 2);
                            xlWorkSheet.Cells[17, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }
                    }

                    #endregion

                    #region Device 2

                    if (Global.filePathLOGList.Count == 2)
                    {
                        #region 1

                        xlWorkSheet.Cells[17, 2] = Math.Round(Global.AvgTemp1, 2);
                        xlWorkSheet.Cells[17, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[17, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 3] = Math.Round(Global.AvgHumi1, 2);
                            xlWorkSheet.Cells[17, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[17, 3] = Math.Round(Global.AvgTemp2, 2);
                            xlWorkSheet.Cells[17, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 4] = Math.Round(Global.AvgHumi2, 2);
                                xlWorkSheet.Cells[17, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 4] = Math.Round(Global.AvgTemp2, 2);
                            xlWorkSheet.Cells[17, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 5] = Math.Round(Global.AvgHumi2, 2);
                                xlWorkSheet.Cells[17, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }

                    #endregion

                    #region Device 3

                    if (Global.filePathLOGList.Count == 3)
                    {
                        #region 1

                        xlWorkSheet.Cells[17, 2] = Math.Round(Global.AvgTemp1, 2);
                        xlWorkSheet.Cells[17, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[17, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 3] = Math.Round(Global.AvgHumi1, 2);
                            xlWorkSheet.Cells[17, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[17, 3] = Math.Round(Global.AvgTemp2, 2);
                            xlWorkSheet.Cells[17, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 4] = Math.Round(Global.AvgHumi2, 2);
                                xlWorkSheet.Cells[17, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 4] = Math.Round(Global.AvgTemp2, 2);
                            xlWorkSheet.Cells[17, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 5] = Math.Round(Global.AvgHumi2, 2);
                                xlWorkSheet.Cells[17, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[17, 4] = Math.Round(Global.AvgTemp3, 2);
                            xlWorkSheet.Cells[17, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 5] = Math.Round(Global.AvgHumi3, 2);
                                xlWorkSheet.Cells[17, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 6] = Math.Round(Global.AvgTemp3, 2);
                            xlWorkSheet.Cells[17, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 7] = Math.Round(Global.AvgHumi3, 2);
                                xlWorkSheet.Cells[17, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[17, 5] = Math.Round(Global.AvgTemp3, 2);
                            xlWorkSheet.Cells[17, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 6] = Math.Round(Global.AvgHumi3, 2);
                                xlWorkSheet.Cells[17, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 4

                    if (Global.filePathLOGList.Count == 4)
                    {
                        #region 1

                        xlWorkSheet.Cells[17, 2] = Math.Round(Global.AvgTemp1, 2);
                        xlWorkSheet.Cells[17, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[17, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 3] = Math.Round(Global.AvgHumi1, 2);
                            xlWorkSheet.Cells[17, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[17, 3] = Math.Round(Global.AvgTemp2, 2);
                            xlWorkSheet.Cells[17, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 4] = Math.Round(Global.AvgHumi2, 2);
                                xlWorkSheet.Cells[17, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 4] = Math.Round(Global.AvgTemp2, 2);
                            xlWorkSheet.Cells[17, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 5] = Math.Round(Global.AvgHumi2, 2);
                                xlWorkSheet.Cells[17, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[17, 4] = Math.Round(Global.AvgTemp3, 2);
                            xlWorkSheet.Cells[17, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 5] = Math.Round(Global.AvgHumi3, 2);
                                xlWorkSheet.Cells[17, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 6] = Math.Round(Global.AvgTemp3, 2);
                            xlWorkSheet.Cells[17, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 7] = Math.Round(Global.AvgHumi3, 2);
                                xlWorkSheet.Cells[17, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[17, 5] = Math.Round(Global.AvgTemp3, 2);
                            xlWorkSheet.Cells[17, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 6] = Math.Round(Global.AvgHumi3, 2);
                                xlWorkSheet.Cells[17, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 4

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[17, 5] = Math.Round(Global.AvgTemp4, 2);
                            xlWorkSheet.Cells[17, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 6] = Math.Round(Global.AvgHumi4, 2);
                                xlWorkSheet.Cells[17, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 8] = Math.Round(Global.AvgTemp4, 2);
                            xlWorkSheet.Cells[17, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 9] = Math.Round(Global.AvgHumi4, 2);
                                xlWorkSheet.Cells[17, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 9].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[17, 7] = Math.Round(Global.AvgTemp4, 2);
                            xlWorkSheet.Cells[17, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 8] = Math.Round(Global.AvgHumi4, 2);
                                xlWorkSheet.Cells[17, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[17, 6] = Math.Round(Global.AvgTemp4, 2);
                            xlWorkSheet.Cells[17, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[17, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                xlWorkSheet.Cells[17, 7] = Math.Round(Global.AvgHumi4, 2);
                                xlWorkSheet.Cells[17, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[17, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }
                    #endregion

                    #endregion

                    #region MKT

                    xlWorkSheet.Cells[18, 1] = "MKT";
                    xlWorkSheet.Cells[18, 1].Font.Size = 10;
                    xlWorkSheet.Cells[18, 1].Font.Bold = true;
                    xlWorkSheet.Cells[18, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[18, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;

                    #region Device 1

                    if (Global.filePathLOGList.Count == 1)
                    {
                        xlWorkSheet.Cells[18, 2] = Global.MKTValue1;
                        xlWorkSheet.Cells[18, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[18, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            //xlWorkSheet.Cells[17, 3] = Global.MKTValue1;
                            xlWorkSheet.Cells[18, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }
                    }

                    #endregion

                    #region Device 2

                    if (Global.filePathLOGList.Count == 2)
                    {
                        #region 1

                        xlWorkSheet.Cells[18, 2] = Global.MKTValue1;
                        xlWorkSheet.Cells[18, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[18, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            //xlWorkSheet.Cells[17, 3] = Global.TCount1;
                            xlWorkSheet.Cells[18, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[18, 3] = Global.MKTValue2;
                            xlWorkSheet.Cells[18, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                // xlWorkSheet.Cells[17, 4] = Global.TCount2;
                                xlWorkSheet.Cells[18, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[18, 4] = Global.MKTValue2;
                            xlWorkSheet.Cells[18, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                //xlWorkSheet.Cells[17, 5] = Global.TCount2;
                                xlWorkSheet.Cells[18, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }

                    #endregion

                    #region Device 3

                    if (Global.filePathLOGList.Count == 3)
                    {
                        #region 1

                        xlWorkSheet.Cells[18, 2] = Global.MKTValue1;
                        xlWorkSheet.Cells[18, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[18, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            // xlWorkSheet.Cells[17, 3] = Global.TCount1;
                            xlWorkSheet.Cells[18, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[18, 3] = Global.MKTValue2;
                            xlWorkSheet.Cells[18, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                //  xlWorkSheet.Cells[17, 4] = Global.TCount2;
                                xlWorkSheet.Cells[18, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[18, 4] = Global.MKTValue2;
                            xlWorkSheet.Cells[18, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                //xlWorkSheet.Cells[17, 5] = Global.TCount2;
                                xlWorkSheet.Cells[18, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[18, 4] = Global.MKTValue3;
                            xlWorkSheet.Cells[18, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                // xlWorkSheet.Cells[17, 5] = Global.TCount3;
                                xlWorkSheet.Cells[18, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[18, 6] = Global.MKTValue3;
                            xlWorkSheet.Cells[18, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                // xlWorkSheet.Cells[17, 7] = Global.TCount3;
                                xlWorkSheet.Cells[18, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[18, 5] = Global.MKTValue3;
                            xlWorkSheet.Cells[18, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                //xlWorkSheet.Cells[17, 6] = Global.TCount3;
                                xlWorkSheet.Cells[18, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion
                    }

                    #endregion

                    #region Device 4

                    if (Global.filePathLOGList.Count == 4)
                    {
                        #region 1

                        xlWorkSheet.Cells[18, 2] = Global.MKTValue1;
                        xlWorkSheet.Cells[18, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[18, 2].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            //xlWorkSheet.Cells[17, 3] = Global.TCount1;
                            xlWorkSheet.Cells[18, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        }

                        #endregion

                        #region 2

                        if (Global.modelValue1 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[18, 3] = Global.MKTValue2;
                            xlWorkSheet.Cells[18, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 3].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                //xlWorkSheet.Cells[17, 4] = Global.TCount2;
                                xlWorkSheet.Cells[18, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[18, 4] = Global.MKTValue2;
                            xlWorkSheet.Cells[18, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue2 == "PDL-K03")
                            {
                                //xlWorkSheet.Cells[17, 5] = Global.TCount2;
                                xlWorkSheet.Cells[18, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 3

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[18, 4] = Global.MKTValue3;
                            xlWorkSheet.Cells[18, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 4].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                // xlWorkSheet.Cells[17, 5] = Global.TCount3;
                                xlWorkSheet.Cells[18, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[18, 6] = Global.MKTValue3;
                            xlWorkSheet.Cells[18, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                // xlWorkSheet.Cells[17, 7] = Global.TCount3;
                                xlWorkSheet.Cells[18, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[18, 5] = Global.MKTValue3;
                            xlWorkSheet.Cells[18, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue3 == "PDL-K03")
                            {
                                // xlWorkSheet.Cells[17, 6] = Global.TCount3;
                                xlWorkSheet.Cells[18, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                        #region 4

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[18, 5] = Global.MKTValue4;
                            xlWorkSheet.Cells[18, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 5].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                //xlWorkSheet.Cells[17, 6] = Global.TCount4;
                                xlWorkSheet.Cells[18, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[18, 8] = Global.MKTValue4;
                            xlWorkSheet.Cells[18, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                // xlWorkSheet.Cells[17, 9] = Global.TCount4;
                                xlWorkSheet.Cells[18, 9].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 9].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            xlWorkSheet.Cells[18, 7] = Global.MKTValue4;
                            xlWorkSheet.Cells[18, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                // xlWorkSheet.Cells[17, 8] = Global.TCount4;
                                xlWorkSheet.Cells[18, 8].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 8].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            xlWorkSheet.Cells[18, 6] = Global.MKTValue4;
                            xlWorkSheet.Cells[18, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            xlWorkSheet.Cells[18, 6].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            if (Global.modelValue4 == "PDL-K03")
                            {
                                //xlWorkSheet.Cells[17, 7] = Global.TCount4;
                                xlWorkSheet.Cells[18, 7].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                xlWorkSheet.Cells[18, 7].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            }
                        }

                        #endregion

                    }
                    #endregion

                    #endregion

                    #region Device Report
                    xlWorkSheet.Cells[20, 1] = "Device Report";
                    xlWorkSheet.Cells[20, 1].Font.Size = 11;
                    xlWorkSheet.Cells[20, 1].Font.Bold = true;
                    xlWorkSheet.Cells[20, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlWorkSheet.Cells[20, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    #endregion

                    for (int k = 1; k < dataGridView.Columns.Count + 1; k++)
                    {
                        xlWorkSheet.Cells[21, k] = dataGridView.Columns[k - 1].HeaderText;
                        xlWorkSheet.Cells[21, k].Font.Size = 10;
                        xlWorkSheet.Cells[21, k].Font.Bold = true;
                        xlWorkSheet.Cells[21, k].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlWorkSheet.Cells[21, k].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    }


                    copyAlltoClipboard();
                    Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[22, 1];
                    CR.Select();
                    xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);


                    //for (i = 0; i < dataGridView.Rows.Count; i++)
                    //{
                    //    for (j = 0; j < dataGridView.Columns.Count; j++)
                    //    {


                    //        xlWorkSheet.Cells[i + 22, j + 1] = dataGridView.Rows[i].Cells[j].Value.ToString();
                    //        xlWorkSheet.Cells[i + 22, j + 1].Font.Size = 9;
                    //        xlWorkSheet.Cells[i + 22, j + 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    //        xlWorkSheet.Cells[i + 22, j + 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    //    }
                    //}



                    string Count = dataGridView.RowCount.ToString();
                    Int32 Cell = Convert.ToInt32(Count) + 24;

                    range.Merge();
                    range.Font.Bold = true;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    xlWorkSheet.Cells[Cell, 1] = "" + PDFWithLogo.PreparedAndPrintedBy + " : " + "                                          ";

                    range.Merge();
                    range.Font.Bold = true;
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    xlWorkSheet.Cells[Cell, 5] = "Reviewed By :" + "                                          ";



                    range = xlWorkSheet.get_Range("B11:I17");
                    range.Font.Size = 9;

                    range = xlWorkSheet.get_Range("B19:Z100000");
                    range.NumberFormat = "0.0";
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    range = xlWorkSheet.get_Range("A11:Z100000");
                    range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xlWorkSheet.Columns["A:K"].AutoFit();

                    xlWorkBook.SaveAs(saveDialog.FileName);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                    MessageBox.Show("Excel report created", "Save", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    Clipboard.Clear();

                }
            }
            catch (Exception e)
            {
                FileLog.ErrorLog(e.Message + e.StackTrace);
                //SaveFileDialog saveDialog = new SaveFileDialog();
                //saveDialog.Filter = "Excel files (*.csv)|All files (*.*)|*.*";
            }
        }
        private void copyAlltoClipboard()
        {


            dataGridView.SelectAll();
            DataObject dataObj = dataGridView.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        /// <summary>
        /// Import Hex file and generate the tabuler report and graph.
        /// </summary>
        public void LoadHexFile(string FileName)
        {
            try
            {
                OpenFileDialog fldDialog = new OpenFileDialog();
                string targetDirectory = string.Empty;

                fldDialog.FileName = "Folder Selection.";
                if (fldDialog.ShowDialog() == DialogResult.OK)
                {
                    targetDirectory = Path.GetDirectoryName(fldDialog.FileName);

                    string[] fileEntries = Directory.GetFiles(targetDirectory, "*.txt");

                    foreach (string fileName in fileEntries)
                    {

                        string fileNameBoth = Path.GetFileName(fileName);

                        if (fileNameBoth == "CFG.txt")
                        {
                            Global.filePathCFG = fileName;
                            devComm.ReadCFGFile();
                        }
                        if (fileNameBoth == "LOG.txt")
                        {
                            Global.filePathLOG = fileName;
                        }

                    }
                }
            }
            catch (Exception e)
            {
                FileLog.ErrorLog(e.Message + e.StackTrace);
            }

        }
        /// <summary>
        /// Show loading form
        /// </summary>
        public void showForm()
        {
            try
            {
                load.ShowDialog();
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }
        public void printPDF()
        {
            System.IO.FileStream fs = null;
            try
            {
                #region dummy
                fs = new FileStream(System.IO.Path.GetTempPath() + @"enviLOG Basic" + "\\printPDF1.pdf", FileMode.Create);
                //    fs = new FileStream(Application.StartupPath + @"\PDFDATA" + @"\" + Global.pdfFileName + "_New.pdf", FileMode.Create);
                iTextSharp.text.Document document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30, 30, 30, 30);
                iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, fs);
                writer.PageEvent = new MyHeaderFooterEvent();
                document.SetPageSize(iTextSharp.text.PageSize.A4);

                #region Header - Footer
                //------------------------Adding Company Name & Location----------------------------------------------------------------------------------------------------
                string sysFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
                HeaderFooter header = new HeaderFooter(new Phrase(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyName"] + "\n" + System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLoc"] + "\n\n" + "Data Report" + "\n\n" + "From Date & Time :" + Global.fromDateTimeReport + "                                       To Date & Time :" + Global.toDateTimeReport + "                                                                                                         Print Date & Time :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm")), false);
                header.Border = iTextSharp.text.Rectangle.NO_BORDER;
                header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                document.Header = header;


                string d = "" + local.Toatalpagecount.ToString() + " Current Page : ";
                HeaderFooter footer = new HeaderFooter(new Phrase(string.Format(" " + PDFWithLogo.PreparedAndPrintedBy + " :{0}                          Reviewed By :{1}                       Total Pages :{2} ", "", "", " " + d + "")), true);
                footer.Border = iTextSharp.text.Rectangle.NO_BORDER; ;
                document.Footer = footer;

                #endregion

                #region LOGO
                // ------------------------ Adding Logo in First Page------------------------------------------------------------------------------------------------------

                System.Drawing.Image image;
                if (System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"] == "")
                {
                    image = System.Drawing.Image.FromFile(System.IO.Path.GetTempPath() + @"enviLOG Basic" + "\\enviro_logo.jpg");
                }
                else
                {
                    image = System.Drawing.Image.FromFile(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"]);
                }
                Document doc = new Document(PageSize.A4);

                document.Open();
                iTextSharp.text.Image pdfImage = iTextSharp.text.Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Png);
                pdfImage.ScaleToFit(100, 50);

                pdfImage.SetAbsolutePosition(50, 800);
                document.Add(pdfImage);

                #endregion

                // Add a simple and wellknown phrase to the document in a flow layout manner
                iTextSharp.text.Font font5 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 12, 0);
                iTextSharp.text.Font font4 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 13, 1);
                iTextSharp.text.Font font3 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 11, 1);
                iTextSharp.text.Font font2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1);
                iTextSharp.text.Font font1 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 0);
                iTextSharp.text.Font Redfont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.RED);
                iTextSharp.text.Font Bluefont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.BLUE);

                PdfContentByte cb = writer.DirectContent;

                #region MIN,MAX,AVG
                //------------------------Table 2------------------------------------------------------------------------------------------------------

                PdfPTable table2 = null;
                document.Add(new iTextSharp.text.Paragraph("Device Information", font4));
                cb.MoveTo(30, document.Top - 103);
                cb.Stroke();
                document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                #region Column Size

                if (Global.filePathLOGList.Count == 1)
                {
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                }
                else if (Global.filePathLOGList.Count == 2)
                {
                    if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                }
                else if (Global.filePathLOGList.Count == 3)
                {
                    if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                }
                else if (Global.filePathLOGList.Count == 4)
                {
                    if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                        Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                        Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                    if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                       Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                       Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                       Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                }

                #endregion

                table2.WidthPercentage = 100;
                table2.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

                table2.AddCell(new Phrase(string.Format("{0}", "Device Name / Serial No"), font3));

                #region Device Name / Serial No

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    }

                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3]), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3]), font1));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Operating Range"), font3));

                #region Operating Range

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[3]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[3]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[3]).ToString("0.0") + " %RH"), font1));
                    }
                }


                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Transaction Count"), font3));

                #region Trascation Count

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    }
                }

                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    }
                }

                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount4), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount4), font1));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(string.Format("{0}", " "), font1));

                #region Unit

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Minimum"), font3));

                #region Minimum

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi3), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi3), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp4), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi4), font1));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Maximum"), font3));

                #region Maximum

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi3), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi3), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp4), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi4), font1));
                    }
                }


                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Average"), font3));

                #region Average

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp3.ToString("0.00")), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi3.ToString("0.00")), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp3.ToString("0.00")), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi3.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp4.ToString("0.00")), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi4.ToString("0.00")), font1));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "MKT"), font3));

                #region MKT

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue4), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }
                }

                #endregion

                document.Add(table2);
                document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                #endregion

                GenerateGraph();

                //Global.radioButtonValue = "Tabular & Chart";

                #region Chart
                // ------------------------Charts------------------------------------------------------------------------------------------------------
                if (Global.radioButtonValue == "Chart" || Global.radioButtonValue == "Tabular & Chart")
                {
                    document.Add(new iTextSharp.text.Paragraph("Chart Report", font4));

                    System.Drawing.Image img = zedGraphControl.GraphPane.GetImage(534, 330, 1);
                    iTextSharp.text.Image j = iTextSharp.text.Image.GetInstance(img, iTextSharp.text.Color.GREEN);
                    document.Add(j);
                    if (Global.radioButtonValue == "Tabular & Chart")
                    {
                        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                        // document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                        //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    }
                    else if (Global.radioButtonValue == "Chart")
                    {

                    }
                }

                #endregion

                //if (Global.radioButtonValue != "Chart")
                //{

                #region Data Table
                // ------------------------ Data Table------------------------------------------------------------------------------------------------------
                if (Global.radioButtonValue == "Tabular" || Global.radioButtonValue == "Tabular & Chart")
                {
                    PdfPTable table1 = null;
                    PdfPCell cell = null;
                    document.Add(new iTextSharp.text.Paragraph("Device Report", font4));
                    cb.MoveTo(30, document.Top - 103);
                    cb.Stroke();
                    document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                    #region Column Size

                    if (Global.filePathLOGList.Count == 1)
                    {
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 2)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 3)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 4)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                            Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                            Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                    }

                    #endregion

                    table1.WidthPercentage = 100;
                    table1.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

                    for (int j = 0; j < dataGridView.Columns.Count - 0; j++)
                    {
                        cell = new PdfPCell(new Phrase(dataGridView.Columns[j].HeaderText, font3));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        table1.AddCell(cell);
                    }


                    #region DataGirdView Device 1

                    try
                    {
                        if (Global.filePathLOGList.Count == 1)
                        {
                            for (int i = 0; i < Global.TCount1; i++)
                            {
                                table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                }
                                else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                }
                                else
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                }

                                if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                }
                            }
                        }
                    }
                    catch (ArgumentOutOfRangeException e)
                    {
                    }
                    catch (Exception ex)
                    {
                        FileLog.ErrorLog(ex.Message + ex.StackTrace);
                    }

                    #endregion

                    #region DataGridView Device 2

                    try
                    {
                        if (Global.filePathLOGList.Count == 2)
                        {
                            #region 1

                            for (int i = 0; i < Global.TCount1; i++)
                            {
                                table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                }
                                else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                }
                                else
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                }

                                if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }
                                }

                            #endregion

                                #region 2

                                if (Global.modelValue1 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                            }
                                #endregion
                        }
                    }
                    catch (ArgumentOutOfRangeException e)
                    {
                    }
                    catch (Exception ex)
                    {
                        FileLog.ErrorLog(ex.Message + ex.StackTrace);
                    }

                    #endregion

                    #region DataGridView Device 3

                    try
                    {
                        if (Global.filePathLOGList.Count == 3)
                        {
                            #region 1

                            for (int i = 0; i < Global.TCount1; i++)
                            {
                                table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                }
                                else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                }
                                else
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                }

                                if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                }

                            #endregion

                                #region 2

                                if (Global.modelValue1 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion

                                #region 3

                                if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion
                            }
                        }
                    }
                    catch (ArgumentOutOfRangeException e)
                    {
                    }
                    catch (Exception ex)
                    {
                        FileLog.ErrorLog(ex.Message + ex.StackTrace);
                    }
                    #endregion

                    #region DataGridView Device 4

                    try
                    {
                        if (Global.filePathLOGList.Count == 4)
                        {
                            #region 1

                            for (int i = 0; i < Global.TCount1; i++)
                            {
                                table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                }
                                else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                }
                                else
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                }

                                if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                }

                            #endregion

                                #region 2

                                if (Global.modelValue1 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion

                                #region 3

                                if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion

                                #region 4

                                if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue4 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue4 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[8].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[8].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue4 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue4 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion
                            }
                        }
                    }
                    catch (ArgumentOutOfRangeException e)
                    {
                    }
                    catch (Exception ex)
                    {
                        FileLog.ErrorLog(ex.Message + ex.StackTrace);
                    }

                    #endregion

                    document.Add(table1);
                    //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                #endregion

                }

                document.Close();
                writer.Close();
                fs.Close();
                #endregion

                #region Real

                File.Delete(System.IO.Path.GetTempPath() + @"enviLOG Basic" + "\\printPDF1.pdf");
                fs = new FileStream(System.IO.Path.GetTempPath() + @"enviLOG Basic" + "\\printPDF.pdf", FileMode.Create);
                document = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 30, 30, 30, 30);
                writer = iTextSharp.text.pdf.PdfWriter.GetInstance(document, fs);
                writer.PageEvent = new MyHeaderFooterEvent();
                document.SetPageSize(iTextSharp.text.PageSize.A4);

                #region Header - Footer
                //------------------------Adding Company Name & Location----------------------------------------------------------------------------------------------------

                header = new HeaderFooter(new Phrase(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyName"] + "\n" + System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLoc"] + "\n\n" + "Data Report" + "\n\n" + "From Date & Time :" + Global.fromDateTimeReport + "                                       To Date & Time :" + Global.toDateTimeReport + "                                                                                                         Print Date & Time :" + DateTime.Now.ToString("dd-MM-yyyy HH:mm")), false);
                header.Border = iTextSharp.text.Rectangle.NO_BORDER;
                header.Alignment = iTextSharp.text.Element.ALIGN_CENTER;
                document.Header = header;


                d = "" + local.Toatalpagecount.ToString() + " Current Page : ";
                footer = new HeaderFooter(new Phrase(string.Format(" " + PDFWithLogo.PreparedAndPrintedBy + " :{0}                          Reviewed By :{1}                       Total Pages :{2} ", "", "", " " + d + "")), true);
                footer.Border = iTextSharp.text.Rectangle.NO_BORDER; ;
                document.Footer = footer;

                #endregion

                #region LOGO
                // ------------------------ Adding Logo in First Page------------------------------------------------------------------------------------------------------


                if (System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"] == "")
                {
                    image = System.Drawing.Image.FromFile(System.IO.Path.GetTempPath() + @"enviLOG Basic" + "\\enviro_logo.jpg");
                }
                else
                {
                    image = System.Drawing.Image.FromFile(System.Configuration.ConfigurationManager.AppSettings["defaultCompanyLogo"]);
                }
                doc = new Document(PageSize.A4);

                document.Open();
                //pdfImage = iTextSharp.text.Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Png);
                //pdfImage.ScaleToFit(100, 50);

                //pdfImage.SetAbsolutePosition(50, 800);
                //document.Add(pdfImage);

                #endregion

                // Add a simple and wellknown phrase to the document in a flow layout manner
                font5 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 12, 0);
                font4 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 13, 1);
                font3 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 11, 1);
                font2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1);
                font1 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 0);
                Redfont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.RED);
                Bluefont2 = iTextSharp.text.FontFactory.GetFont(FontFactory.HELVETICA, 10, 1, iTextSharp.text.Color.BLUE);

                cb = writer.DirectContent;

                #region MIN,MAX,AVG
                //------------------------Table 2------------------------------------------------------------------------------------------------------

                table2 = null;
                document.Add(new iTextSharp.text.Paragraph("Device Information", font4));
                cb.MoveTo(30, document.Top - 103);
                cb.Stroke();
                document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                #region Column Size

                if (Global.filePathLOGList.Count == 1)
                {
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                }
                else if (Global.filePathLOGList.Count == 2)
                {
                    if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                }
                else if (Global.filePathLOGList.Count == 3)
                {
                    if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                    else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }
                }
                else if (Global.filePathLOGList.Count == 4)
                {
                    if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                    if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                        Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                        Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                        Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                    if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                       Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                       Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                       Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03")
                    {
                        float[] columnDefinitionSize2 = { 3F, 3F, 3F, 3F, 3F, 3F };
                        table2 = new PdfPTable(columnDefinitionSize2);
                    }

                }

                #endregion

                table2.WidthPercentage = 100;
                table2.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

                table2.AddCell(new Phrase(string.Format("{0}", "Device Name / Serial No"), font3));

                #region Device Name / Serial No

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    }

                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[0] + "   /   " + Global.serialNoStrList[0]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[1] + "   /   " + Global.serialNoStrList[1]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[2] + "   /   " + Global.serialNoStrList[2]), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3]), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.deviceNameStrList[3] + "   /   " + Global.serialNoStrList[3]), font1));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Operating Range"), font3));

                #region Operating Range

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[0]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[0]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[0]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[0]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[1]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[1]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[1]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[1]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[2]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[2]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[2]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[2]).ToString("0.0") + " %RH"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.tempLowStrDecimalList[3]).ToString("0.0")) + " - " + Convert.ToDouble(Global.tempHighStrDecimalList[3]).ToString("0.0") + " °C", font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(Global.humiLowStrDecimalList[3]).ToString("0.0") + " - " + Convert.ToDouble(Global.humiHighStrDecimalList[3]).ToString("0.0") + " %RH"), font1));
                    }
                }


                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Transaction Count"), font3));

                #region Trascation Count

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    }
                }

                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    }
                }

                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount3), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.TCount4), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.TCount4), font1));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(string.Format("{0}", " "), font1));

                #region Unit

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", "Temperature  °C"), font3));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "Humidity  %RH"), font3));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Minimum"), font3));

                #region Minimum

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi3), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi3), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinTemp4), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MinHumi4), font1));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Maximum"), font3));

                #region Maximum

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi3), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi1), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi2), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi3), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxTemp4), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0:0.0}", Global.MaxHumi4), font1));
                    }
                }


                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "Average"), font3));

                #region Average

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp3.ToString("0.00")), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi3.ToString("0.00")), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp1.ToString("0.00")), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi1.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp2.ToString("0.00")), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi2.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp3.ToString("0.00")), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi3.ToString("0.00")), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0}", Global.AvgTemp4.ToString("0.00")), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", Global.AvgHumi4.ToString("0.00")), font1));
                    }
                }

                #endregion

                table2.AddCell(new Phrase(String.Format("{0}", "MKT"), font3));

                #region MKT

                if (Global.filePathLOGList.Count == 1)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 2)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 3)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }
                }
                if (Global.filePathLOGList.Count == 4)
                {
                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue1), font1));
                    if (Global.modelValue1 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue2), font1));
                    if (Global.modelValue2 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue3), font1));
                    if (Global.modelValue3 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }

                    table2.AddCell(new Phrase(String.Format("{0:0.00}", Global.MKTValue4), font1));
                    if (Global.modelValue4 == "PDL-K03")
                    {
                        table2.AddCell(new Phrase(String.Format("{0}", "-"), font1));
                    }
                }

                #endregion

                document.Add(table2);
                document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                //document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                #endregion

                #region Chart
                // ------------------------Charts------------------------------------------------------------------------------------------------------
                if (Global.radioButtonValue == "Chart" || Global.radioButtonValue == "Tabular & Chart")
                {


                    document.Add(new iTextSharp.text.Paragraph("Chart Report", font4));

                    System.Drawing.Image img = zedGraphControl.GraphPane.GetImage(534, 330, 1);
                    iTextSharp.text.Image j = iTextSharp.text.Image.GetInstance(img, iTextSharp.text.Color.GREEN);
                    document.Add(j);

                    if (Global.radioButtonValue == "Tabular & Chart")
                    {
                        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                        document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                        // document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                        // document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));
                    }
                    else if (Global.radioButtonValue == "Chart")
                    {

                    }
                }
                #endregion
                #endregion
                //if (Global.radioButtonValue != "Chart")
                //{

                #region Data Table
                // ------------------------ Data Table------------------------------------------------------------------------------------------------------
                if (Global.radioButtonValue == "Tabular" || Global.radioButtonValue == "Tabular & Chart")
                {
                    PdfPTable table1 = null;
                    PdfPCell cell = null;
                    document.Add(new iTextSharp.text.Paragraph("Device Report", font4));
                    cb.MoveTo(30, document.Top - 103);
                    cb.Stroke();
                    document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                    #region Column Size

                    if (Global.filePathLOGList.Count == 1)
                    {
                        if (Global.modelValue1 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 2)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 3)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                        else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }
                    }
                    else if (Global.filePathLOGList.Count == 4)
                    {
                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                        if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                            Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                            Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03" ||
                            Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                        if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" && Global.modelValue4 == "PDL-K01" ||
                           Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" && Global.modelValue4 == "PDL-K03")
                        {
                            float[] columnDefinitionSize1 = { 5F, 3F, 3F, 3F, 3F, 3F };
                            table1 = new PdfPTable(columnDefinitionSize1);
                        }

                    }

                    #endregion

                    table1.WidthPercentage = 100;
                    table1.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;

                    for (int j = 0; j < dataGridView.Columns.Count - 0; j++)
                    {
                        cell = new PdfPCell(new Phrase(dataGridView.Columns[j].HeaderText, font3));
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        table1.AddCell(cell);
                    }

                    #region DataGirdView Device 1

                    try
                    {
                        if (Global.filePathLOGList.Count == 1)
                        {
                            for (int i = 0; i < Global.TCount1; i++)
                            {
                                table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                }
                                else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                }
                                else
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                }

                                if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                }
                            }
                        }
                    }
                    catch (ArgumentOutOfRangeException e)
                    {
                    }
                    catch (Exception ex)
                    {
                        FileLog.ErrorLog(ex.Message + ex.StackTrace);
                    }

                    #endregion

                    #region DataGridView Device 2

                    try
                    {
                        if (Global.filePathLOGList.Count == 2)
                        {
                            #region 1

                            for (int i = 0; i < Global.TCount1; i++)
                            {
                                table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                }
                                else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                }
                                else
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                }

                                if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }
                                }

                            #endregion

                                #region 2

                                if (Global.modelValue1 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                            }
                                #endregion
                        }
                    }
                    catch (ArgumentOutOfRangeException e)
                    {
                    }
                    catch (Exception ex)
                    {
                        FileLog.ErrorLog(ex.Message + ex.StackTrace);
                    }

                    #endregion

                    #region DataGridView Device 3

                    try
                    {
                        if (Global.filePathLOGList.Count == 3)
                        {
                            #region 1

                            for (int i = 0; i < Global.TCount1; i++)
                            {
                                table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                }
                                else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                }
                                else
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                }

                                if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                }

                            #endregion

                                #region 2

                                if (Global.modelValue1 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion

                                #region 3

                                if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion
                            }
                        }
                    }
                    catch (ArgumentOutOfRangeException e)
                    {
                    }
                    catch (Exception ex)
                    {
                        FileLog.ErrorLog(ex.Message + ex.StackTrace);
                    }
                    #endregion

                    #region DataGridView Device 4

                    try
                    {
                        if (Global.filePathLOGList.Count == 4)
                        {
                            #region 1

                            for (int i = 0; i < Global.TCount1; i++)
                            {
                                table1.AddCell(new Phrase(String.Format("{0}", dataGridView.Rows[i].Cells[0].Value.ToString()), font1));

                                if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Bluefont2));
                                }
                                else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[1].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[0]))
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), Redfont2));
                                }
                                else
                                {
                                    table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[1].Value).ToString("N1")), font1));
                                }

                                if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[0]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                }

                            #endregion

                                #region 2

                                if (Global.modelValue1 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[2].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[2].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[1]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue2 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[1]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion

                                #region 3

                                if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[3].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                        }
                                    }
                                }
                                else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[2]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue3 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[2]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion

                                #region 4

                                if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[4].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[4].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue4 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue4 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[8].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[8].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[8].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03" || Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue4 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[7].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[7].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                else if (Global.modelValue1 == "PDL-K03" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K03" && Global.modelValue3 == "PDL-K01" || Global.modelValue1 == "PDL-K01" && Global.modelValue2 == "PDL-K01" && Global.modelValue3 == "PDL-K03")
                                {
                                    if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) <= Convert.ToDecimal(Global.tempLowStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Bluefont2));
                                    }
                                    else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[5].Value) >= Convert.ToDecimal(Global.tempHighStrDecimalList[3]))
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), Redfont2));
                                    }
                                    else
                                    {
                                        table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[5].Value).ToString("N1")), font1));
                                    }

                                    if (Global.modelValue4 == "PDL-K03")
                                    {
                                        if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) <= Convert.ToDecimal(Global.humiLowStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Bluefont2));
                                        }
                                        else if (Convert.ToDecimal(dataGridView.Rows[i].Cells[6].Value) >= Convert.ToDecimal(Global.humiHighStrDecimalList[3]))
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), Redfont2));
                                        }
                                        else
                                        {
                                            table1.AddCell(new Phrase(String.Format("{0}", Convert.ToDouble(dataGridView.Rows[i].Cells[6].Value).ToString("N1")), font1));
                                        }
                                    }
                                }

                                #endregion
                            }
                        }
                    }
                    catch (ArgumentOutOfRangeException e)
                    {
                    }
                    catch (Exception ex)
                    {
                        FileLog.ErrorLog(ex.Message + ex.StackTrace);
                    }

                    #endregion

                    document.Add(table1);
                    //  document.Add(new iTextSharp.text.Paragraph(Environment.NewLine));

                #endregion
                }


                document.Close();
                writer.Close();
                fs.Close();



                Process p = new Process();
                p.StartInfo = new ProcessStartInfo()
                {
                    CreateNoWindow = false,
                    Verb = "print",
                    FileName = (System.IO.Path.GetTempPath() + @"enviLOG Basic" + "\\printPDF.pdf") //put the correct path here
                };

                PDFWithLogo.LogoEveryPage(System.IO.Path.GetTempPath() + @"enviLOG Basic" + "\\printPDF.pdf", "printPDF1.pdf");
                p.Start();

            }
            catch (Exception ex)
            {
                if (ex.Message != null)
                {

                }
            }
            finally
            {
                //  p.Start();
            }
        }

        #endregion

        private void txtboxremark_TextChanged(object sender, EventArgs e)
        {
            Global.remarkTxt = txtboxremark.Text;
            if (string.IsNullOrWhiteSpace(txtboxremark.Text))
            {
                txtboxremark.Clear();
            }
            if (Global.remarkTxt.Length < 8)
            {
                Global.remarkTxt = (Global.remarkTxt).PadRight(8, ' ');
            }
        }


        private void txtboxremark_KeyPress(object sender, KeyPressEventArgs e)
        {
            lblRemark.Visible = true;
        }

        private void txtboxremark_Leave(object sender, EventArgs e)
        {
            lblRemark.Visible = false;
        }

        private void comboBoxType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxType.SelectedIndex == 0)
            {
                dateTimePickstart1.Enabled = true;
                dateTimePickstart2.Enabled = true;
                dateTimePickstop1.Enabled = false;
                dateTimePickstop2.Enabled = false;
                numUpDwnstartdelay.Enabled = false;
                numUpDwnstartdelay.Value = 0;
            }
            if (comboBoxType.SelectedIndex == 1)
            {
                dateTimePickstart1.Enabled = false;
                dateTimePickstart2.Enabled = false;
                dateTimePickstop1.Enabled = false;
                dateTimePickstop2.Enabled = false;
                numUpDwnstartdelay.Enabled = false;
                numUpDwnstartdelay.Value = 0;
            }
            if (comboBoxType.SelectedIndex == 2)
            {
                dateTimePickstart1.Enabled = true;
                dateTimePickstart2.Enabled = true;
                dateTimePickstop1.Enabled = true;
                dateTimePickstop2.Enabled = true;
                numUpDwnstartdelay.Enabled = false;
                numUpDwnstartdelay.Value = 0;
            }
            if (comboBoxType.SelectedIndex == 3)
            {
                dateTimePickstart1.Enabled = false;
                dateTimePickstart2.Enabled = false;
                dateTimePickstop1.Enabled = false;
                dateTimePickstop2.Enabled = false;
                numUpDwnstartdelay.Enabled = true;
            }
            Global.comBoxSelIndValue = comboBoxType.SelectedIndex;
        }

        private void chkbxenableLED_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBoxLED.Checked == true)
            {
                Global.enableLEDValue = "00";
                lblLEDMsg.Text = "( It may consume more battery )";
            }
            else
            {
                lblLEDMsg.Text = "";
            }
        }

        private void chkBoxDespTime_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBoxDispOnTime.Checked == true)
            {
                numUpDwnDispOnTime.Enabled = true;
            }
            else
            {
                numUpDwnDispOnTime.Enabled = false;
                numUpDwnDispOnTime.Value = 0;
            }
        }

        private void chkBoxShowLimit_CheckedChanged(object sender, EventArgs e)
        {
            GenerateGraph();
        }

        private void createPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = rdbn.ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    string value = rdbn.GetrdbnValue();

                    lblshowdatatab_Click(sender, e);
                    lblShowDataTab.Enabled = true;
                    tabControl.TabPages.Remove(tabPageShowDataTab);


                    lblshowdatachart_Click(sender, e);
                    lblShowDataChart.Enabled = true;
                    tabControl.TabPages.Remove(tabPageShowDataChart);

                    CreateUploadGraphNData(dataGridView, Global.serialNoStr + "_" + DateTime.Now.ToString("dd-MM-yyyy HH.mm"));
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        private void exportExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                lblshowdatatab_Click(sender, e);
                lblShowDataTab.Enabled = true;
                tabControl.TabPages.Remove(tabPageShowDataTab);

                ExportExcelFile();
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        private void menuItemPrint_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = rdbn.ShowDialog();
                if (dr == System.Windows.Forms.DialogResult.OK)
                {
                    //Global.radioButtonValue = "Tabular & Chart";
                    lblshowdatatab_Click(sender, e);
                    lblShowDataTab.Enabled = true;
                    tabControl.TabPages.Remove(tabPageShowDataTab);

                    //lblshowdatachart_Click(sender, e);
                    //lblShowDataChart.Enabled = true;
                    //tabControl.TabPages.Remove(tabPageShowDataChart);

                    //CreateUploadGraphNData(dataGridView, Global.serialNoStr + "_" + DateTime.Now.ToString("dd-MM-yyyy HH.mm"));
                    System.Diagnostics.Process.Start(Global.printPDFName);

                    printPDF();
                }
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        private void importHexFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    if (Global.lblShowDataTabWasClicked == true)
            //    {
            //        tabControl.TabPages.Remove(tabPageShowDataTab);
            //        lblShowDataTab.Enabled = true;
            //    }
            //    if (Global.lblShowDataChartWasClicked == true)
            //    {
            //        tabControl.TabPages.Remove(tabPageShowDataChart);
            //        lblShowDataChart.Enabled = true;
            //    }

            //    LoadHexFile("Report.txt");
            //    lblshowdatatab_Click(sender, e);
            //}
            //catch (Exception ex)
            //{
            //    FileLog.ErrorLog(ex.Message + ex.StackTrace);
            //}
        }

        private void menuItemSetting_Click(object sender, EventArgs e)
        {
            try
            {
                chgpass.ShowDialog();
            }
            catch (Exception ex)
            {
            }

        }

        private void menuItemHelp_Click(object sender, EventArgs e)
        {
            //AboutUs au = new AboutUs();
            //au.Show();
        }

        private void numUpDwnDispOnTime_ValueChanged(object sender, EventArgs e)
        {
            Global.displayOnValue = Convert.ToInt32(numUpDwnDispOnTime.Value);
        }

        private void numUpDwnTempMin_ValueChanged(object sender, EventArgs e)
        {
            if (numUpDwnTempMin == this.ActiveControl)
            {
                if (numUpDwnTempMin.Value > numUpDwnTempMax.Value)
                {
                    MessageBox.Show("Minimum temperature value should be less than maximum temperature value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    numUpDwnTempMin.Value = numUpDwnTempMax.Value;
                }
            }
        }

        private void numUpDwnTempMax_ValueChanged(object sender, EventArgs e)
        {
            if (numUpDwnTempMax == this.ActiveControl)
            {
                if (numUpDwnTempMax.Value < numUpDwnTempMin.Value)
                {
                    MessageBox.Show("Maximum temperature value should be greater than minimum temperature value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    numUpDwnTempMax.Value = numUpDwnTempMin.Value;
                }
            }

        }

        private void numUpDwnHumiMin_ValueChanged(object sender, EventArgs e)
        {
            if (numUpDwnHumiMin == this.ActiveControl)
            {
                if (numUpDwnHumiMin.Value > numUpDwnHumiMax.Value)
                {
                    MessageBox.Show("Minimum humidity value should be less than maximum humidity value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    numUpDwnHumiMin.Value = numUpDwnHumiMax.Value;
                }
            }
        }

        private void numUpDwnHumiMax_ValueChanged(object sender, EventArgs e)
        {
            if (numUpDwnHumiMax == this.ActiveControl)
            {
                if (numUpDwnHumiMax.Value < numUpDwnHumiMin.Value)
                {
                    MessageBox.Show("Maximum humidity value should be greater than minimum humidity value", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    numUpDwnHumiMax.Value = numUpDwnHumiMin.Value;
                }
            }
        }

        private void dateTimePickstart1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePickstart1 == this.ActiveControl)
            {
                if (dateTimePickstart1.Value < DateTime.Now.Date)
                {
                    MessageBox.Show("Start date should be next from current date", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dateTimePickstart1.Value = DateTime.Now.Date;
                }
            }
        }

        private void dateTimePickstop1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePickstop1 == this.ActiveControl)
            {
                if (dateTimePickstop1.Value < dateTimePickstart1.Value)
                {
                    MessageBox.Show("Stop date should be next from start date", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dateTimePickstop1.Value = dateTimePickstart1.Value;
                }
            }
        }

        private void numUpDwnTempMin_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (e.KeyChar == '.')
            {
                e.Handled = true;
            }

        }

        private void numUpDwnHumiMin_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.Handled = true;
            }

        }

        private void numUpDwnTempMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.Handled = true;
            }
        }

        private void numUpDwnHumiMax_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.Handled = true;
            }
        }

        private void numUpDwninterval_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.Handled = true;
            }
        }

        private void numUpDwnDispOnTime_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.Handled = true;
            }
        }

        private void numUpDwnstartdelay_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '.')
            {
                e.Handled = true;
            }
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            Thread.CurrentThread.Priority = ThreadPriority.Highest;
            sf = new Thread(new ThreadStart(showForm));
            sf.Priority = ThreadPriority.Lowest;
            sf.Start();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            try
            {
                this.Enabled = true;
                this.Invoke(new Action(() => { load.Close(); }));
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
        }

        private void menuItemSave_Click(object sender, EventArgs e)
        {

        }

        private void userManualToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string str = Environment.CurrentDirectory + "\\enviLOG UserManual.pdf";
            System.Diagnostics.Process.Start(str);
        }

        private void installationGuideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string str = Environment.CurrentDirectory + "\\INSTALLATION GUIDE OF enviLOG Basic.pdf";
            System.Diagnostics.Process.Start(str);
        }

        private void aboutUsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutUs au = new AboutUs();
            au.Show();
        }

    }

    public class PDFFooter : PdfPageEventHelper
    {
        // write on top of document
        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            base.OnOpenDocument(writer, document);
            PdfPTable tabFot = new PdfPTable(new float[] { 1F });
            tabFot.SpacingAfter = 10F;
            PdfPCell cell;
            tabFot.TotalWidth = 300F;
            cell = new PdfPCell(new Phrase("Header"));
            tabFot.AddCell(cell);
            tabFot.WriteSelectedRows(0, -1, 150, document.Top, writer.DirectContent);
        }

        // write on start of each page
        public override void OnStartPage(PdfWriter writer, Document document)
        {
            base.OnStartPage(writer, document);
        }

        // write on end of each page
        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);
            PdfPTable tabFot = new PdfPTable(new float[] { 1F });
            PdfPCell cell;
            tabFot.TotalWidth = 300F;
            cell = new PdfPCell(new Phrase("Footer"));
            tabFot.AddCell(cell);
            tabFot.WriteSelectedRows(0, -1, 150, document.Bottom, writer.DirectContent);
        }

        //write on close of document
        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);
        }



    }
}
//

