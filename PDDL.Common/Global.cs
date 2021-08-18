using System.IO;
using System;

using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows;
using System.Management;
using System.Windows.Forms;
using System.Configuration;
using System.Globalization;
using System.Security.Principal;
using System.Security.AccessControl;
using System.Collections.Generic;
using System.Security.Cryptography;


namespace PDDL.Common
{
    public static class Global
    {

        public static string model;
        public static string[] AllPortList;
        public static string portName;
        public static string filePath;
        public static string filePathLOG;
        public static string filePathCFG;
        public static string readDataAll;
        public static string[] readDataAllSplit;
        public static string dateTimeTrasactionStr;
        public static List<string> dateTimeTrasactionList = new List<string>();
        public static List<string> filePathLOGList = new List<string>();
        public static List<string> filePathCFGList = new List<string>();

        public static string responseStringW;
        public static string responseStringC;
        public static string responseStringD;
        public static string responseStringP;
        public static string responseStringX;
        public static string responseStringR;
        public static string responseStringF;
        public static string responseStringV;
        public static string responseStringA;
        public static string responseStringY;
        public static string responseStringN;
        public static string responseStringJ;
        public static string responseStringK;
        public static string responseStringB;



        public static string[] responseStringFA;

        public static string[] commaSplit;
        public static string responseDevName;
        public static string responseFirmVer;
        public static string responseProgramLog;
        public static string responseProgLogger;
        public static string responseLoggerTime;
        public static string responseMinMaxRead;
        public static string responseLEDValue;


        public static string dateTimeStr;
        public static string dateStr1;
        public static string dateStr2;
        public static string timeStr1;
        public static string timeStr2;
        public static string dateStrSwap;
        public static string timeStrSwap;
        public static int timeStrDecimalHr;
        public static int timeStrDecimalMin;
        public static int timeStrMinInterval;
        public static int timeIntervalMaxCount;
        public static string timeStrDecimal;
        public static string timeStrDecimalInter;
        public static int dateStrDecimalDay;
        public static int dateStrDecimalMonth;
        public static int dateStrDecimalYear;
        public static string dateStrDecimal;
        public static List<string> timeStrMinList = new List<string>();
        public static List<string> timeStrList = new List<string>();
        public static List<string> dateStrList = new List<string>();

        public static string timeStrDecimalP;  // Get time $P command
        public static string dateStrDecimalP;  // Get date $P command

        public static string timeStrDecimalC;  // Get Starttime $C command
        public static string dateStrDecimalC;  // Get Startdate $C command
        public static string StTimeStrDecimalC;// Get Stoptime $C command
        public static string StDateStrDecimalC;// Get StopDate $C command

        public static string timeStrDecimalD;  // Get Starttime $D command
        public static string dateStrDecimalD;  // Get Startdate $D command
        public static string StTimeStrDecimalD;// Get Stoptime $D command
        public static string StDateStrDecimalD;// Get StopDate $D command

        public static string readingStr;
        public static string tempStr1;
        public static string tempStr2;
        public static string humidiStr1;
        public static string humidiStr2;
        public static string tempStrSwap;
        public static string humidiStrSwap;
        public static double temperature;
        public static double humidity;
        public static List<string> temperatureList = new List<string>();
        public static List<string> humidityList = new List<string>();
        public static List<string> temperatureGraphList1 = new List<string>();
        public static List<string> humidityGraphList1 = new List<string>();
        public static List<string> temperatureGraphList2 = new List<string>();
        public static List<string> humidityGraphList2 = new List<string>();
        public static List<string> temperatureGraphList3 = new List<string>();
        public static List<string> humidityGraphList3 = new List<string>();
        public static List<string> temperatureGraphList4 = new List<string>();
        public static List<string> humidityGraphList4 = new List<string>();
        public static double maxTemperature;
        public static double minTemperature;
        public static double maxHumidity;
        public static double minHumidity;

        public static string intervalStr;
        public static string intervalStr1;
        public static string intervalStr2;
        public static string intervalStrSwap;
        public static string intervalStrDecimal;

        public static string dispTimeStrDecimal;
        public static string delayStrDecimal;
        public static string tempHighStrDecimal;
        public static string tempLowStrDecimal;
        public static string humiHighStrDecimal;
        public static string humiLowStrDecimal;
        public static string deviceModeStr;
        public static string deviceNameStr;
        public static string serialNoStr;
        public static string chkSerialNoStr;

        public static List<string> deviceNameStrList = new List<string>();
        public static List<string> serialNoStrList = new List<string>();
        public static List<string> tempHighStrDecimalList = new List<string>();
        public static List<string> tempLowStrDecimalList = new List<string>();
        public static List<string> humiHighStrDecimalList = new List<string>();
        public static List<string> humiLowStrDecimalList = new List<string>();
        public static List<string> transactionCountList = new List<string>();

        public static string formatStr;
        public static string eventStr;
        public static string dummyByteStr;

        public static string StDateTimeStr;
        public static string StDateStr1;
        public static string StDateStr2;
        public static string StTimeStr1;
        public static string StTimeStr2;
        public static string StDateStrSwap;
        public static string StTimeStrSwap;
        public static string StTimeStrDecimal;
        public static string StDateStrDecimal;

        public static string transCountstr;
        public static string transCountstr1;
        public static string transCountstr2;
        public static string transCountstrSwap;
        public static string transCountstrDecimal;

        public static string devLogStatusStr;

        public static string endOfCommand = "0D";
        public static string endOfCommandAscii = "";
        public static string errorMessage = "";
        public static double tempFormValue1 = -46.85;
        public static double tempFormValue2 = 175.72;
        public static double humidiFormValue1 = -6;
        public static double humidiFormValue2 = 125;
        public static double divideValue = 65536;
        public static string remarkTxt;
        public static int comBoxSelIndValue;
        public static string dateTimeStart;
        public static string dateTimeStop;
        public static string intervalValue;
        public static string delayValue;
        public static string tempMinLimit;
        public static string tempMaxLimit;
        public static string humiMinLimit;
        public static string humiMaxLimit;
        public static string tempSign;
        public static int valueTempMin;
        public static int valueTempMax;
        public static string enableLEDValue = "10";
        public static string displayOnTimeValue = "02";
        public static int displayOnValue;
        public static string loggerDateTime;

        public static string currDateTime;
        public static string printPDFName;

        public static bool btnSearchWasClicked = false;
        public static bool btnReadDataWasClicked = false;
        public static bool lblProgramWasClicked = false;
        public static bool lblShowDataTabWasClicked = false;
        public static bool lblShowDataChartWasClicked = false;

        public static double MaxTemp1 = 0;
        public static double MinTemp1 = 0;
        public static double AvgTemp1 = 0;
        public static double MaxHumi1 = 0;
        public static double MinHumi1 = 0;
        public static double AvgHumi1 = 0;

        public static double MaxTemp2 = 0;
        public static double MinTemp2 = 0;
        public static double AvgTemp2 = 0;
        public static double MaxHumi2 = 0;
        public static double MinHumi2 = 0;
        public static double AvgHumi2 = 0;

        public static double MaxTemp3 = 0;
        public static double MinTemp3 = 0;
        public static double AvgTemp3 = 0;
        public static double MaxHumi3 = 0;
        public static double MinHumi3 = 0;
        public static double AvgHumi3 = 0;

        public static double MaxTemp4 = 0;
        public static double MinTemp4 = 0;
        public static double AvgTemp4 = 0;
        public static double MaxHumi4 = 0;
        public static double MinHumi4 = 0;
        public static double AvgHumi4 = 0;

        public static List<double> calMaxHumiList = new List<double>();
        public static List<double> calMaxTempList = new List<double>();
        public static List<double> calMinHumiList = new List<double>();
        public static List<double> calMinTempList = new List<double>();

        public static string sFolderName;
        public static string dFolderName;

        public static int majorStepCount;
        public static int transactionCount;
        public static int dataGridViewCount = 0;
        public static int generateGraphCount = 0;
        public static double MKTValue1;
        public static double MKTValue2;
        public static double MKTValue3;
        public static double MKTValue4;
        public static int threadValue;

        public static string companyLogoValue;
        public static string compLogoDefName = "Image";
        public static string pdfFileName;

        public static int CountValue = 0;

        public static string radioButtonValue;

        public static string userNameValue;
        public static string passwordValue;
        public static bool correctValue = false;

        public static string newPwdValue;
        public static string masterPwdValue;
        public static int timerInterval;
        public static int TCount1;
        public static int TCount2;
        public static int TCount3;
        public static int TCount4;

        public static string fromDateTimeReport;
        public static string toDateTimeReport;
        public static string modelValue1;
        public static string modelValue2;
        public static string modelValue3;
        public static string modelValue4;
        public static string sMacAddress = string.Empty;
        public static string encyMacAdd;
        public static string decyMacAdd;
        public static string decyKeyMacAdd;

        /// <summary>
        /// convert hex to ASCII of End of Command (0D)
        /// </summary>   
        public static string ConvertHextoAscii()
        {
            string asciiString = "";
            for (int i = 0; i < endOfCommand.Length; i += 2)
            {
                if (endOfCommand.Length >= i + 2)
                {
                    String hs = endOfCommand.Substring(i, 2);
                    asciiString = asciiString + System.Convert.ToChar(System.Convert.ToUInt32(endOfCommand.Substring(i, 2), 16)).ToString();
                }
            }
            endOfCommandAscii = asciiString;
            return asciiString;
        }

        /// <summary>
        /// get the current date time of system
        /// </summary>
        public static void CurrentDateTime()
        {
            string dd = DateTime.Now.ToString("dd");
            string MM = DateTime.Now.ToString("MM");
            string yy = DateTime.Now.ToString("yy");
            string hh = DateTime.Now.ToString("HH");
            string mm = DateTime.Now.ToString("mm");
            Global.currDateTime = hh + mm + dd + MM + yy;
        }

        /// <summary>
        ///  To full access to other users 
        /// </summary>
        public static class Access
        {
            public static void GrantAccess(string fullPath)
            {
                try
                {
                    DirectoryInfo info = new DirectoryInfo(fullPath);
                    WindowsIdentity self = System.Security.Principal.WindowsIdentity.GetCurrent();
                    DirectorySecurity ds = info.GetAccessControl();
                    ds.AddAccessRule(new FileSystemAccessRule(self.Name, FileSystemRights.FullControl, InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit,
                    PropagationFlags.NoPropagateInherit, AccessControlType.Allow));

                    info.SetAccessControl(ds);

                }
                catch (Exception ex)
                {
                }

            }

        }

        /// <summary>
        /// Ths method is used for decrypting the mac address
        /// </summary>
        /// <param name="Source"></param>
        /// <param name="Key"></param>
        /// <returns></returns>
        private static byte[] DESKey = { 200, 5, 78, 232, 9, 6, 0, 4 };
        private static byte[] DESInitializationVector = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };
        public static string Decrypting(string Source, string Key)
        {
            using (var cryptoProvider = new DESCryptoServiceProvider())
            using (var memoryStream = new System.IO.MemoryStream(Convert.FromBase64String(Source)))
            using (var cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(DESKey, DESInitializationVector), CryptoStreamMode.Read))
            using (var reader = new System.IO.StreamReader(cryptoStream))
            {
                return reader.ReadToEnd();
            }
        }

        ////
    }
}
