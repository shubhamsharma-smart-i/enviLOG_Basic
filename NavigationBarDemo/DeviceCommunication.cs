using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.IO.Ports;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Data;
using System.Windows.Forms;
using PDDL.Common;
using PDDL.Interface;
using System.Globalization;


namespace PDDL
{
    public class DeviceCommunication
    {
        #region Service initialization/ de-initialization

        SerialPort serialPort = new SerialPort();
        string path = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);

        public DeviceCommunication()
        {
            Global.ConvertHextoAscii();
            connectAllSerialPort();
            sendCommandCheckSerialNo();
        }

        /// <summary>
        /// Get All Available Ports 
        /// </summary>
        public void GetPorts()
        {
            Global.AllPortList = SerialPort.GetPortNames();
        }

        /// <summary>
        /// This method opens connection to COM port
        /// </summary>
        /// <param name="PortName">COM5</param>
        /// <param name="BaudRate">19200</param>
        /// <param name="Parity">Parity.None</param>
        /// <param name="DataBits">8</param>
        /// 
        /// 
        /// <param name="StopBits">StopBits.One</param>
        public void connectAllSerialPort()
        {
            GetPorts();
            try
            {
                for (int i = 0; i < Global.AllPortList.Length; i++)
                {
                    serialPort.PortName = Global.AllPortList[i];
                    serialPort.BaudRate = 19200;
                    serialPort.Parity = Parity.None;
                    serialPort.DataBits = 8;
                    serialPort.StopBits = StopBits.One;


                    serialPort.ReadBufferSize = 3000000;
                    serialPort.WriteBufferSize = 3000000;

                    try
                    {
                        serialPort.Open();
                        serialPort.DiscardOutBuffer();
                        serialPort.DiscardInBuffer();
                        serialPort.RtsEnable = true;
                        serialPort.DtrEnable = true;
                    }
                    catch
                    {

                    }
                    autoDevSearch();
                }

            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }

        }

        /// <summary>
        /// Send command for device search when application get started
        /// </summary>
        public void autoDevSearch()
        {
            string commandText = string.Format("$W00");
            string message = commandText + Global.endOfCommandAscii;
            Global.responseStringW = string.Empty;

            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.WriteTimeout = 1000;
                    serialPort.Write(message);
                    serialPort.WriteTimeout = -1;

                    Thread.Sleep(200);
                    Global.responseStringW += serialPort.ReadExisting();

                    if (Global.responseStringW == "" || Global.responseStringW == null)
                    {
                        serialPort.Close();
                    }
                    Global.commaSplit = Global.responseStringW.Split(new char[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    string splitFirmVer = Global.commaSplit[0];
                    Global.responseFirmVer = splitFirmVer.Substring(2, splitFirmVer.Length - 2);

                    Global.responseDevName = Global.commaSplit[1];
                    Global.portName = serialPort.PortName;
                    Global.model = Global.responseDevName;
                    FileLog.CommandLog("Sent command :" + "    " + message);  //Shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringW); //Shubham
                }
                catch (Exception ex)
                {
                    serialPort.Close();
                    //MessageBox.Show(ex.ToString());
                    FileLog.ErrorLog(ex.Message + ex.StackTrace); //Shubham
                    //  FileLog.TempLog(ex.Message + ex.StackTrace);
                    serialPort.WriteTimeout = -1;
                    // FileLog.TempLog(ex.Message + ex.StackTrace);
                }

            }
        }

        /// <summary>
        /// This method opens connection to COM port
        /// </summary>
        /// <param name="PortName">COM5</param>
        /// <param name="BaudRate">19200</param>
        /// <param name="Parity">Parity.None</param>
        /// <param name="DataBits">8</param>
        /// <param name="StopBits">StopBits.One</param>
        public void connectSerialPort()
        {
            GetPorts();
            try
            {
                serialPort.PortName = Global.portName;
                serialPort.BaudRate = 19200;
                serialPort.Parity = Parity.None;
                serialPort.DataBits = 8;
                serialPort.StopBits = StopBits.One;


                serialPort.ReadBufferSize = 3000000;
                serialPort.WriteBufferSize = 3000000;

                serialPort.Open();
                serialPort.DiscardOutBuffer();
                serialPort.DiscardInBuffer();
                serialPort.RtsEnable = true;
                serialPort.DtrEnable = true;

            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace); //Shubham
            }

        }

        /// <summary>
        /// Get All Device Configuration
        /// </summary>
        #region GetDeviceConfiguration

        /// <summary>
        /// Get version of device $W
        /// </summary>
        public void sendCommandDeviceSearch()
        {
            string commandText = string.Format("$W00");
            string message = commandText + Global.endOfCommandAscii;
            Global.responseStringW = string.Empty;


            try
            {
                serialPort.Close();

                if (serialPort.IsOpen == false)
                {
                    connectAllSerialPort();
                }
            }

            //if (serialPort.IsOpen)
            //{
            //    try
            //    {
            //        serialPort.Write(message);
            //        Thread.Sleep(200);
            //        Global.responseStringW += serialPort.ReadExisting();

            //        if (Global.responseStringW == "" || Global.responseStringW == null)
            //        {
            //            System.Windows.Forms.MessageBox.Show("Device is not connected");
            //            serialPort.Close();
            //        }

            //        Global.responseDevName = Global.responseStringW.Substring(1, Global.responseStringW.Length - 2);
            //        FileLog.CommandLog("Sent command :" + "    " + message);
            //        FileLog.CommandLog("Received Data :" + "   " + Global.responseStringW);
            //    }

            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace); //Shubham
            }
            // }


        }

        /// <summary>
        /// Get Start datetime,Stop datetime,interval,format,event,dummy,total transaction,total device log status. $C
        /// </summary>
        public void sendCommandGetProgramLog()
        {
            string commandText = string.Format("$C00");
            string message = commandText + Global.endOfCommandAscii;

            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }

            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(200);
                    FileLog.CommandLog("Sent command :" + "    " + message);
                    GetResponseGetProgramLog();

                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);//shubham
                }
            }

        }

        /// <summary>
        /// Response of $C command
        /// </summary>
        public void GetResponseGetProgramLog()
        {
            try
            {
                Global.responseStringC = string.Empty;
                Global.responseStringC += serialPort.ReadExisting();
                Global.responseProgramLog = Global.responseStringC.Substring(2, Global.responseStringC.Length - 3);

                #region StartDateTime

                Global.dateTimeStr = Global.responseProgramLog.Substring(0, 8);

                Global.timeStr1 = Global.dateTimeStr.Substring(0, 2);
                Global.timeStr2 = Global.dateTimeStr.Substring(2, 2);
                Global.dateStr1 = Global.dateTimeStr.Substring(4, 2);
                Global.dateStr2 = Global.dateTimeStr.Substring(6, 2);

                Global.timeStrSwap = string.Concat(Global.timeStr2, Global.timeStr1);
                Global.dateStrSwap = string.Concat(Global.dateStr2, Global.dateStr1);

                string timeStrBinary = Convert.ToString(Convert.ToInt64(Global.timeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                string dateStrBinary = Convert.ToString(Convert.ToInt64(Global.dateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                int timeStrDecimalHr = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 11, 5), 2);
                int timeStrDecimalMin = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 6, 6), 2);
                Global.timeStrDecimalC = timeStrDecimalHr.ToString("00") + ":" + timeStrDecimalMin.ToString("00");

                int dateStrDecimalDay = Convert.ToInt32(dateStrBinary.Substring(0, 5), 2);
                int dateStrDecimalMonth = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 11, 4), 2);
                int dateStrDecimalYear = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 7, 7), 2);
                Global.dateStrDecimalC = dateStrDecimalDay.ToString("00") + "-" + dateStrDecimalMonth.ToString("00") + "-" + "20" + dateStrDecimalYear.ToString("00");

                #endregion

                #region LogInterval

                Global.intervalStr = Global.responseProgramLog.Substring(8, 4);

                Global.intervalStr1 = Global.intervalStr.Substring(0, 2);
                Global.intervalStr2 = Global.intervalStr.Substring(2, 2);
                Global.intervalStrSwap = string.Concat(Global.intervalStr2 + Global.intervalStr1);

                Global.intervalStrDecimal = Convert.ToString(Convert.ToInt32(Global.intervalStrSwap.Substring(0, 4), 16));

                #endregion

                #region Format,Event,Dummybyte
                Global.formatStr = Global.responseProgramLog.Substring(12, 2);
                Global.eventStr = Global.responseProgramLog.Substring(14, 2);
                Global.dummyByteStr = Global.responseProgramLog.Substring(16, 2);
                #endregion

                #region StopDateTime

                Global.StDateTimeStr = Global.responseProgramLog.Substring(Global.responseProgramLog.Length - 14, 8);

                Global.StTimeStr1 = Global.StDateTimeStr.Substring(0, 2);
                Global.StTimeStr2 = Global.StDateTimeStr.Substring(2, 2);
                Global.StDateStr1 = Global.StDateTimeStr.Substring(4, 2);
                Global.StDateStr2 = Global.StDateTimeStr.Substring(6, 2);

                Global.StTimeStrSwap = string.Concat(Global.StTimeStr2, Global.StTimeStr1);
                Global.StDateStrSwap = string.Concat(Global.StDateStr2, Global.StDateStr1);

                string StTimeStrBinary = Convert.ToString(Convert.ToInt64(Global.StTimeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                string StDateStrBinary = Convert.ToString(Convert.ToInt64(Global.StDateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                int StTimeStrDecimalHr = Convert.ToInt32(StTimeStrBinary.Substring(StTimeStrBinary.Length - 11, 5), 2);
                int StTimeStrDecimalMin = Convert.ToInt32(StTimeStrBinary.Substring(StTimeStrBinary.Length - 6, 6), 2);
                Global.StTimeStrDecimalC = StTimeStrDecimalHr.ToString("00") + ":" + StTimeStrDecimalMin.ToString("00");

                int StDateStrDecimalDay = Convert.ToInt32(StDateStrBinary.Substring(0, 5), 2);
                int StDateStrDecimalMonth = Convert.ToInt32(StDateStrBinary.Substring(StDateStrBinary.Length - 11, 4), 2);
                int StDateStrDecimalYear = Convert.ToInt32(StDateStrBinary.Substring(StDateStrBinary.Length - 7, 7), 2);
                Global.StDateStrDecimalC = StDateStrDecimalDay.ToString("00") + "-" + StDateStrDecimalMonth.ToString("00") + "-" + "20" + StDateStrDecimalYear.ToString("00");

                #endregion

                #region TotalTransactionStoreCount

                Global.transCountstr = Global.responseProgramLog.Substring(Global.responseProgramLog.Length - 6, 4);

                Global.transCountstr1 = Global.transCountstr.Substring(0, 2);
                Global.transCountstr2 = Global.transCountstr.Substring(2, 2);

                Global.transCountstrSwap = string.Concat(Global.transCountstr2, Global.transCountstr1);
                Global.transCountstrDecimal = Convert.ToString(Convert.ToInt32(Global.transCountstrSwap.Substring(0, 4), 16));

                #endregion

                #region DeviceLogStatus

                Global.devLogStatusStr = Global.responseProgramLog.Substring(Global.responseProgramLog.Length - 2, 2);

                #endregion

                FileLog.CommandLog("Received Data :" + "   " + Global.responseStringC);//shubham
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);//shubham
            }
        }

        /// <summary>
        /// Get display on time,start time ,stop time, delay ,temperature high & low limit,
        /// humidity high & low limit,temperature sign,device mode,device name set by user,serial no. $D
        /// </summary>
        public void sendCommandGetProgLogger()
        {
            string commandText = string.Format("$D00");
            string message = commandText + Global.endOfCommandAscii;

            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }

            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(200);
                    FileLog.CommandLog("Sent command :" + "    " + message);//shubham
                    GetResponseGetProgLogger();
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace); //shubham
                }
            }
        }

        /// <summary>
        /// Response of $D command`
        /// </summary>
        public void GetResponseGetProgLogger()
        {
            try
            {
                Global.responseStringD = string.Empty;
                Global.responseStringD += serialPort.ReadExisting();
                Global.responseProgLogger = Global.responseStringD.Substring(Global.responseStringD.Length - 55, 54);

                #region DisplayOnTime

                string dispTimeStr = Global.responseProgLogger.Substring(0, 2);
                string dispTimeStrBinary1 = Convert.ToString(Convert.ToInt64(dispTimeStr, 16), 2).PadLeft(8, '0');
                string dispTimeStrBinary2 = dispTimeStrBinary1.Substring(dispTimeStrBinary1.Length - 8, 6).PadLeft(8, '0');
                int dispTimeStrDecimal1 = Convert.ToInt32(dispTimeStrBinary2.Substring(0, 4), 2);
                int dispTimeStrDecimal2 = Convert.ToInt32(dispTimeStrBinary2.Substring(dispTimeStrBinary2.Length - 4, 4), 2);
                Global.dispTimeStrDecimal = Convert.ToString(dispTimeStrDecimal1 + dispTimeStrDecimal2);

                #endregion

                #region StartDateTime

                Global.dateTimeStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 52, 8);

                Global.timeStr1 = Global.dateTimeStr.Substring(0, 2);
                Global.timeStr2 = Global.dateTimeStr.Substring(2, 2);
                Global.dateStr1 = Global.dateTimeStr.Substring(4, 2);
                Global.dateStr2 = Global.dateTimeStr.Substring(6, 2);

                Global.timeStrSwap = string.Concat(Global.timeStr2, Global.timeStr1);
                Global.dateStrSwap = string.Concat(Global.dateStr2, Global.dateStr1);

                string timeStrBinary = Convert.ToString(Convert.ToInt64(Global.timeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                string dateStrBinary = Convert.ToString(Convert.ToInt64(Global.dateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                int timeStrDecimalHr = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 11, 5), 2);
                int timeStrDecimalMin = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 6, 6), 2);
                Global.timeStrDecimalD = timeStrDecimalHr.ToString("00") + ":" + timeStrDecimalMin.ToString("00");

                int dateStrDecimalDay = Convert.ToInt32(dateStrBinary.Substring(0, 5), 2);
                int dateStrDecimalMonth = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 11, 4), 2);
                int dateStrDecimalYear = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 7, 7), 2);
                Global.dateStrDecimalD = dateStrDecimalDay.ToString("00") + "-" + dateStrDecimalMonth.ToString("00") + "-" + "20" + dateStrDecimalYear.ToString("00");

                #endregion

                #region StopDateTime

                Global.StDateTimeStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 44, 8);

                Global.StTimeStr1 = Global.StDateTimeStr.Substring(0, 2);
                Global.StTimeStr2 = Global.StDateTimeStr.Substring(2, 2);
                Global.StDateStr1 = Global.StDateTimeStr.Substring(4, 2);
                Global.StDateStr2 = Global.StDateTimeStr.Substring(6, 2);

                Global.StTimeStrSwap = string.Concat(Global.StTimeStr2, Global.StTimeStr1);
                Global.StDateStrSwap = string.Concat(Global.StDateStr2, Global.StDateStr1);

                string StTimeStrBinary = Convert.ToString(Convert.ToInt64(Global.StTimeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                string StDateStrBinary = Convert.ToString(Convert.ToInt64(Global.StDateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                int StTimeStrDecimalHr = Convert.ToInt32(StTimeStrBinary.Substring(StTimeStrBinary.Length - 11, 5), 2);
                int StTimeStrDecimalMin = Convert.ToInt32(StTimeStrBinary.Substring(StTimeStrBinary.Length - 6, 6), 2);
                Global.StTimeStrDecimalD = StTimeStrDecimalHr.ToString("00") + ":" + StTimeStrDecimalMin.ToString("00");

                int StDateStrDecimalDay = Convert.ToInt32(StDateStrBinary.Substring(0, 5), 2);
                int StDateStrDecimalMonth = Convert.ToInt32(StDateStrBinary.Substring(StDateStrBinary.Length - 11, 4), 2);
                int StDateStrDecimalYear = Convert.ToInt32(StDateStrBinary.Substring(StDateStrBinary.Length - 7, 7), 2);
                Global.StDateStrDecimalD = StDateStrDecimalDay.ToString("00") + "-" + StDateStrDecimalMonth.ToString("00") + "-" + "20" + StDateStrDecimalYear.ToString("00");

                #endregion

                #region LogDelay

                string delayStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 36, 4);
                string delayStr1 = delayStr.Substring(0, 2);
                string delayStr2 = delayStr.Substring(2, 2);
                string delayStrSwap = string.Concat(delayStr2, delayStr1);
                Global.delayStrDecimal = Convert.ToString(Convert.ToInt32(delayStrSwap, 16));

                #endregion

                #region tempHighAlarm

                string tempHighStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 32, 4);
                string tempHighStr1 = tempHighStr.Substring(0, 2);
                string tempHighStr2 = tempHighStr.Substring(2, 2);
                string tempHighStrSwap = string.Concat(tempHighStr2, tempHighStr1);
                Global.tempHighStrDecimal = Convert.ToString(Convert.ToUInt32(tempHighStrSwap, 16));

                #endregion

                #region tempLowAlarm

                string tempLowStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 28, 4);
                string tempLowStr1 = tempLowStr.Substring(0, 2);
                string tempLowStr2 = tempLowStr.Substring(2, 2);
                string tempLowStrSwap = string.Concat(tempLowStr2, tempLowStr1);
                Global.tempLowStrDecimal = Convert.ToString(Convert.ToUInt32(tempLowStrSwap, 16));

                #endregion

                #region humiHighAlarm

                string humiHighStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 24, 2);
                Global.humiHighStrDecimal = Convert.ToString(Convert.ToInt32(humiHighStr, 16));

                #endregion

                #region humiLowAlarm

                string humiLowStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 22, 2);
                Global.humiLowStrDecimal = Convert.ToString(Convert.ToInt32(humiLowStr, 16));

                #endregion

                #region tempAlarmSign

                string tempsignStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 20, 2);
                if (tempsignStr == "00")
                {
                    Global.tempHighStrDecimal = Global.tempHighStrDecimal;
                    Global.tempLowStrDecimal = Global.tempLowStrDecimal;
                }
                if (tempsignStr == "01")
                {
                    Global.tempLowStrDecimal = (65536 - Convert.ToInt32(Global.tempLowStrDecimal)).ToString();
                    Global.tempHighStrDecimal = Global.tempHighStrDecimal;
                    Global.tempLowStrDecimal = "-" + Global.tempLowStrDecimal;
                }
                if (tempsignStr == "02")
                {
                    Global.tempHighStrDecimal = (65536 - Convert.ToInt32(Global.tempHighStrDecimal)).ToString();
                    Global.tempLowStrDecimal = (65536 - Convert.ToInt32(Global.tempLowStrDecimal)).ToString();
                    Global.tempHighStrDecimal = "-" + Global.tempHighStrDecimal;
                    Global.tempLowStrDecimal = "-" + Global.tempLowStrDecimal;
                }
                #endregion

                #region deviceStartMode

                Global.deviceModeStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 18, 2);

                #endregion

                #region deviceName

                Global.deviceNameStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 16, 8);

                #endregion

                #region serialNo

                Global.serialNoStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 8, 8);

                #endregion

                FileLog.CommandLog("Received Data :" + "   " + Global.responseStringD);//shubham
                FileLog.DataLog("Device configuration");//shubham
                File.AppendAllText(Global.filePathCFG, Global.responseDevName + Environment.NewLine);
                File.AppendAllText(Global.filePathCFG, Global.responseProgramLog + Environment.NewLine);
                File.AppendAllText(Global.filePathCFG, Global.responseProgLogger + Environment.NewLine);

            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace); //shubham
            }
        }

        /// <summary>
        /// Get Device Date & Time $P
        /// </summary>
        public void sendCommandLoggerTime()
        {
            string commandText = string.Format("$P00");
            string message = commandText + Global.endOfCommandAscii;
            Global.responseStringP = string.Empty;
            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }
            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(200);
                    Global.responseStringP += serialPort.ReadExisting();
                    Global.responseLoggerTime = Global.responseStringP.Substring(Global.responseStringP.Length - 9, 8);

                    Global.dateTimeStr = Global.responseLoggerTime;
                    Global.timeStr1 = Global.dateTimeStr.Substring(0, 2);
                    Global.timeStr2 = Global.dateTimeStr.Substring(2, 2);
                    Global.dateStr1 = Global.dateTimeStr.Substring(4, 2);
                    Global.dateStr2 = Global.dateTimeStr.Substring(6, 2);

                    Global.timeStrSwap = string.Concat(Global.timeStr2, Global.timeStr1);
                    Global.dateStrSwap = string.Concat(Global.dateStr2, Global.dateStr1);

                    string timeStrBinary = Convert.ToString(Convert.ToInt64(Global.timeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                    string dateStrBinary = Convert.ToString(Convert.ToInt64(Global.dateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                    Global.timeStrDecimalHr = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 11, 5), 2);
                    Global.timeStrDecimalMin = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 6, 6), 2);
                    Global.timeStrDecimalP = Global.timeStrDecimalHr.ToString("00") + " : " + Global.timeStrDecimalMin.ToString("00");

                    int dateStrDecimalDay = Convert.ToInt32(dateStrBinary.Substring(0, 5), 2);
                    int dateStrDecimalMonth = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 11, 4), 2);
                    int dateStrDecimalYear = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 7, 7), 2);
                    Global.dateStrDecimalP = dateStrDecimalDay.ToString("00") + "-" + dateStrDecimalMonth.ToString("00") + "-" + "20" + dateStrDecimalYear.ToString("00");

                    Global.loggerDateTime = Global.dateStrDecimalP + "  " + Global.timeStrDecimalP;

                    FileLog.CommandLog("Sent command :" + "    " + message);//shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringP);//shubham
                    File.AppendAllText(Global.filePathCFG, Global.responseLoggerTime + Environment.NewLine);
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);//shubham
                }
            }

        }

        /// <summary>
        /// Get Min Max temperature & humidity $X
        /// </summary>
        public void sendCommandGetMinMaxRead()
        {
            string commandText = string.Format("$X00");
            string message = commandText + Global.endOfCommandAscii;
            Global.responseStringX = string.Empty;
            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }
            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(200);
                    Global.responseStringX += serialPort.ReadExisting();
                    Global.responseMinMaxRead = Global.responseStringX.Substring(Global.responseStringX.Length - 17, 16);

                    string tempMaxStr = Global.responseMinMaxRead.Substring(0, 4);
                    string tempMinStr = Global.responseMinMaxRead.Substring(Global.responseMinMaxRead.Length - 12, 4);
                    string humiMaxStr = Global.responseMinMaxRead.Substring(Global.responseMinMaxRead.Length - 8, 4);
                    string humiMinStr = Global.responseMinMaxRead.Substring(Global.responseMinMaxRead.Length - 4, 4);

                    string tempMaxStr1 = tempMaxStr.Substring(0, 2);
                    string tempMaxStr2 = tempMaxStr.Substring(tempMaxStr.Length - 2, 2);
                    string tempMinStr1 = tempMinStr.Substring(0, 2);
                    string tempMinStr2 = tempMinStr.Substring(tempMinStr.Length - 2, 2);

                    string humiMaxStr1 = humiMaxStr.Substring(0, 2);
                    string humiMaxStr2 = humiMaxStr.Substring(humiMaxStr.Length - 2, 2);
                    string humiMinStr1 = humiMinStr.Substring(0, 2);
                    string humiMinStr2 = humiMinStr.Substring(humiMinStr.Length - 2, 2);

                    string tempMaxStrSwap = string.Concat(tempMaxStr2, tempMaxStr1);
                    string tempMinStrSwap = string.Concat(tempMinStr2, tempMinStr1);
                    string humiMaxStrSwap = string.Concat(humiMaxStr2, humiMaxStr1);
                    string humiMinStrSwap = string.Concat(humiMinStr2, humiMinStr1);

                    int tempMaxStrDecimal = Convert.ToInt32(tempMaxStrSwap, 16);
                    int tempMinStrDecimal = Convert.ToInt32(tempMinStrSwap, 16);
                    int humiMaxStrDecimal = Convert.ToInt32(humiMaxStrSwap, 16);
                    int humiMinStrDecimal = Convert.ToInt32(humiMinStrSwap, 16);

                    Global.maxTemperature = Math.Round((Global.tempFormValue1 + (Global.tempFormValue2 * (Convert.ToDouble(tempMaxStrDecimal) / 65536))), 1);
                    Global.minTemperature = Math.Round((Global.tempFormValue1 + (Global.tempFormValue2 * (Convert.ToDouble(tempMinStrDecimal) / 65536))), 1);
                    Global.maxHumidity = Math.Round((Global.humidiFormValue1 + (Global.humidiFormValue2 * (Convert.ToDouble(humiMaxStrDecimal) / 65536))), 1);
                    Global.minHumidity = Math.Round((Global.humidiFormValue1 + (Global.humidiFormValue2 * (Convert.ToDouble(humiMinStrDecimal) / 65536))), 1);

                    FileLog.CommandLog("Sent command :" + "    " + message);// shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringX); //shubham
                    File.AppendAllText(Global.filePathCFG, Global.responseMinMaxRead + Environment.NewLine);
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace); //shubham
                }
            }

        }

        /// <summary>
        /// Get value of LED enable or disable $R
        /// </summary>
        public void sendCommandGetLEDValue()
        {
            string commandText = string.Format("$R00");
            string message = commandText + Global.endOfCommandAscii;
            Global.responseStringR = string.Empty;
            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }
            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(200);
                    Global.responseStringR += serialPort.ReadExisting();
                    Global.responseLEDValue = Global.responseStringR.Substring(Global.responseStringR.Length - 3, 2);

                    FileLog.CommandLog("Sent command :" + "    " + message); //shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringR); //shubham
                    File.AppendAllText(Global.filePathCFG, Global.responseLEDValue + Environment.NewLine);
                }
                catch (Exception ex)
                {

                    FileLog.ErrorLog(ex.Message + ex.StackTrace); //shubham
                }
            }
        }

        /// <summary>
        /// Download all transaction from device & save in text file. 
        /// </summary>
        public void sendCommandReadData()
        {
            string commandText = string.Format("$F00");
            string message = commandText + Global.endOfCommandAscii;
            Global.responseStringF = string.Empty;

            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }

            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(Global.threadValue);
                    Global.responseStringF += serialPort.ReadExisting();

                    if (Global.responseStringF == "" || Global.responseStringF == null)
                    {
                        System.Windows.Forms.MessageBox.Show("No any transactions", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        File.WriteAllText(Global.filePathLOG, string.Empty);
                        File.AppendAllText(Global.filePathLOG, Global.responseStringF);
                        if (Global.filePathLOGList.Count < 4)
                        {
                            Global.filePathCFGList.Add(Global.filePathCFG);
                            Global.filePathLOGList.Add(Global.filePathLOG);
                        }
                    }
                }

                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace); //shubham
                }

            }

        }

        /// <summary>
        /// Send $D command for checking the device serial no when setting the new configuration
        /// </summary>
        public void sendCommandCheckSerialNo()
        {
            string commandText = string.Format("$D00");
            string message = commandText + Global.endOfCommandAscii;
            Global.responseStringD = string.Empty;

            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(200);
                    Global.responseStringD += serialPort.ReadExisting();
                    Global.responseProgLogger = Global.responseStringD.Substring(Global.responseStringD.Length - 55, 54);
                    Global.chkSerialNoStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 8, 8);
                    FileLog.CommandLog("Sent command :" + "    " + message);// shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringD);// shubham
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace); //shubham
                }
            }
        }

        #endregion

        /// <summary>
        /// Set All Device Configuration 
        /// </summary>
        #region  SetDeviceConfiguration

        /// <summary>
        /// Set Name of Device $V
        /// </summary>
        public void sendCommandSetRemark()
        {
            string commandText = string.Format("$V");
            string message = commandText + Global.remarkTxt + Global.serialNoStr + Global.endOfCommandAscii;
            Global.responseStringV = string.Empty;

            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }

            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(300);
                    Global.responseStringV += serialPort.ReadExisting();
                    FileLog.CommandLog("Sent command :" + "    " + message); //shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringV); //shubham
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace); //shubham
                }
            }
        }

        /// <summary>
        /// Set Mode of Device $A
        /// </summary>
        public void sendCommandSetMode()
        {
            //#region Mode 1

            //if (Global.comBoxSelIndValue == 0)
            //{
            //    string commandText = string.Format("$A");
            //    string message = commandText + "19" + Global.currDateTime + Global.currDateTime + Global.endOfCommandAscii;
            //    Global.responseStringA = string.Empty;
            //    if (serialPort.IsOpen == false)
            //    {
            //        serialPort.Open();
            //    }
            //    if (serialPort.IsOpen)
            //    {
            //        try
            //        {
            //            serialPort.Write(message);
            //            Thread.Sleep(300);
            //            Global.responseStringA += serialPort.ReadExisting();
            //            FileLog.CommandLog("Sent command :"+"    " + message);
            //            FileLog.CommandLog("Received Data :" +"   "+ Global.responseStringA);
            //        }
            //        catch(Exception e)
            //        {
            //            FileLog.ErrorLog(e.Message + e.StackTrace);
            //        }
            //    }
            //}
            //#endregion

            #region Mode 1

            if (Global.comBoxSelIndValue == 0)
            {
                string commandText = string.Format("$A");
                string message = commandText + "02" + Global.dateTimeStart + Global.currDateTime + Global.endOfCommandAscii;
                Global.responseStringA = string.Empty;
                if (serialPort.IsOpen == false)
                {
                    serialPort.Open();
                }
                if (serialPort.IsOpen)
                {
                    try
                    {
                        serialPort.Write(message);
                        Thread.Sleep(300);
                        Global.responseStringA += serialPort.ReadExisting();
                        FileLog.CommandLog("Sent command :" + "    " + message); //shubham
                        FileLog.CommandLog("Received Data :" + "   " + Global.responseStringA); //shubham
                    }
                    catch (Exception e)
                    {
                        FileLog.ErrorLog(e.Message + e.StackTrace); //shubham
                    }
                }

            }
            #endregion

            #region Mode 2

            if (Global.comBoxSelIndValue == 1)
            {
                string commandText = string.Format("$A");
                string message = commandText + "00" + Global.currDateTime + Global.currDateTime + Global.endOfCommandAscii;
                Global.responseStringA = string.Empty;
                if (serialPort.IsOpen == false)
                {
                    serialPort.Open();
                }
                if (serialPort.IsOpen)
                {
                    try
                    {
                        serialPort.Write(message);
                        Thread.Sleep(300);
                        Global.responseStringA += serialPort.ReadExisting();
                        FileLog.CommandLog("Sent command :" + "    " + message);// shubham
                        FileLog.CommandLog("Received Data :" + "   " + Global.responseStringA);// shubham
                    }
                    catch (Exception e)
                    {
                        FileLog.ErrorLog(e.Message + e.StackTrace);// shubham
                    }
                }
            }

            #endregion

            #region Mode 3

            if (Global.comBoxSelIndValue == 2)
            {
                string commandText = string.Format("$A");
                string message = commandText + "06" + Global.dateTimeStart + Global.dateTimeStop + Global.endOfCommandAscii;
                Global.responseStringA = string.Empty;
                if (serialPort.IsOpen == false)
                {
                    serialPort.Open();
                }
                if (serialPort.IsOpen)
                {
                    try
                    {
                        serialPort.Write(message);
                        Thread.Sleep(300);
                        Global.responseStringA += serialPort.ReadExisting();
                        FileLog.CommandLog("Sent command :" + "    " + message);// shubham
                        FileLog.CommandLog("Received Data :" + "   " + Global.responseStringA); //shubham
                    }
                    catch (Exception e)
                    {
                        FileLog.ErrorLog(e.Message + e.StackTrace); //shubham
                    }
                }
            }

            #endregion

            #region Mode 4

            if (Global.comBoxSelIndValue == 3)
            {
                string commandText = string.Format("$A");
                string message = commandText + "11" + Global.currDateTime + Global.currDateTime + Global.endOfCommandAscii;
                Global.responseStringA = string.Empty;
                if (serialPort.IsOpen == false)
                {
                    serialPort.Open();
                }
                if (serialPort.IsOpen)
                {
                    try
                    {
                        serialPort.Write(message);
                        Thread.Sleep(300);
                        Global.responseStringA += serialPort.ReadExisting();
                        FileLog.CommandLog("Sent command :" + "    " + message);// shubham
                        FileLog.CommandLog("Received Data :" + "   " + Global.responseStringA); //shubham
                    }
                    catch (Exception e)
                    {
                        FileLog.ErrorLog(e.Message + e.StackTrace); //shubham
                    }
                }
            }
            #endregion
        }

        /// <summary>
        /// Set Interval ,Delay And Limits to the device $Y
        /// </summary>
        public void sendCommandSetInterDelLimi()
        {
            string commandText = string.Format("$Y");
            string message = commandText + Global.intervalValue + Global.delayValue + Global.tempMaxLimit + Global.tempMinLimit + Global.humiMaxLimit + Global.humiMinLimit + Global.tempSign + Global.endOfCommandAscii;
            Global.responseStringY = string.Empty;
            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }

            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(300);
                    Global.responseStringY += serialPort.ReadExisting();
                    FileLog.CommandLog("Sent command :" + "    " + message); //shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringY); //shubham
                }
                catch (Exception e)
                {
                    FileLog.ErrorLog(e.Message + e.StackTrace); //shubham
                }
            }
        }

        /// <summary>
        /// Set Display on Time $J
        /// </summary>
        public void sendCommandDisplayOnTime()
        {
            string commandText = string.Format("$J");
            string message = commandText + Global.displayOnTimeValue + Global.endOfCommandAscii;
            Global.responseStringJ = string.Empty;
            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }
            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(300);
                    Global.responseStringJ += serialPort.ReadExisting();
                    FileLog.CommandLog("Sent command :" + "    " + message);// shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringJ);// shubham
                }
                catch (Exception e)
                {
                    FileLog.ErrorLog(e.Message + e.StackTrace);// shubham
                }
            }

        }

        /// <summary>
        /// Enable Disable LED $N
        /// </summary>
        public void sendCommandEnableLED()
        {
            string commandText = string.Format("$N");
            string message = commandText + Global.enableLEDValue + Global.endOfCommandAscii;
            Global.responseStringN = string.Empty;
            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }

            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(300);
                    Global.responseStringN += serialPort.ReadExisting();
                    FileLog.CommandLog("Sent command :" + "    " + message); //shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringN);// shubham
                }
                catch (Exception e)
                {
                    FileLog.ErrorLog(e.Message + e.StackTrace);// shubham
                }
            }
        }

        /// <summary>
        /// Synchronize time of device and System $K
        /// </summary>
        public void sendCommandSyncTime()
        {
            string commandText = string.Format("$K");
            string message = commandText + Global.currDateTime + Global.endOfCommandAscii;
            Global.responseStringK = string.Empty;
            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }
            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(300);
                    Global.responseStringK += serialPort.ReadExisting();
                    FileLog.CommandLog("Sent command :" + "    " + message);// shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringK); //shubham
                }
                catch (Exception e)
                {
                    FileLog.ErrorLog(e.Message + e.StackTrace); //shubham
                }
            }

        }

        /// <summary>
        /// Device setting configuration done(Accept all Setting) $B
        /// </summary>
        public void sendCommandAllSett()
        {
            string commandText = string.Format("$B00");
            string message = commandText + Global.endOfCommandAscii;
            Global.responseStringB = string.Empty;
            if (serialPort.IsOpen == false)
            {
                serialPort.Open();
            }

            if (serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(message);
                    Thread.Sleep(300);
                    Global.responseStringB += serialPort.ReadExisting();
                    FileLog.CommandLog("Sent command :" + "    " + message); //shubham
                    FileLog.CommandLog("Received Data :" + "   " + Global.responseStringB); //shubham
                }
                catch (Exception e)
                {
                    FileLog.ErrorLog(e.Message + e.StackTrace);// shubham
                }
            }
        }

        #endregion

        /// <summary>
        /// Read the LOG Text File which contains all transaction.
        /// </summary>
        public void ReadLOGFile()
        {
            try
            {
                try
                {
                    Global.dateStrList.Clear();
                    Global.timeStrList.Clear();
                    Global.temperatureList.Clear();
                    Global.humidityList.Clear();
                    Global.dateTimeTrasactionList.Clear();

                    Global.readDataAll = File.ReadAllText(Global.filePathLOG);
                }
                catch (FileNotFoundException ex)
                {
                    MessageBox.Show("Please download the data", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                catch (ArgumentNullException ex)
                {
                    return;
                }

                Global.readDataAllSplit = Global.readDataAll.Split(new[] { "#T" }, StringSplitOptions.RemoveEmptyEntries);

                if (Global.readDataAllSplit.Length > 10)
                {
                    //Global.majorStepCount = Global.TCount1 / 10;
                    Global.majorStepCount = Global.readDataAllSplit.Length / 10;        /////Calculate value for majorstep of x axis.
                }
                else
                {
                    Global.majorStepCount = 1;
                }

                for (int i = 0; i <= Global.readDataAllSplit.Length; i++)
                {
                    if (i == 0)
                    {
                        Global.dateTimeStr = Global.readDataAllSplit[i].
                            Substring(0, Global.readDataAllSplit[i].Length - 1);

                        Global.timeStr1 = Global.dateTimeStr.Substring(0, 2);
                        Global.timeStr2 = Global.dateTimeStr.Substring(2, 2);
                        Global.dateStr1 = Global.dateTimeStr.Substring(4, 2);
                        Global.dateStr2 = Global.dateTimeStr.Substring(6, 2);

                        Global.timeStrSwap = string.Concat(Global.timeStr2, Global.timeStr1);
                        Global.dateStrSwap = string.Concat(Global.dateStr2, Global.dateStr1);

                        string timeStrBinary = Convert.ToString(Convert.ToInt64(Global.timeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                        string dateStrBinary = Convert.ToString(Convert.ToInt64(Global.dateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                        Global.timeStrDecimalHr = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 11, 5), 2);
                        Global.timeStrDecimalMin = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 6, 6), 2);
                        Global.timeStrDecimal = Global.timeStrDecimalHr.ToString("00") + " : " + Global.timeStrDecimalMin.ToString("00");
                        if (i == 0)
                        {
                            Global.timeStrList.Add(Global.timeStrDecimal);
                        }
                        Global.timeStrMinInterval = Global.timeStrDecimalMin;

                        Global.dateStrDecimalDay = Convert.ToInt32(dateStrBinary.Substring(0, 5), 2);
                        Global.dateStrDecimalMonth = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 11, 4), 2);
                        Global.dateStrDecimalYear = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 7, 7), 2);
                        Global.dateStrDecimal = Global.dateStrDecimalDay.ToString("00") + "-" + Global.dateStrDecimalMonth.ToString("00") + "-" + "20" + Global.dateStrDecimalYear;
                        Global.dateStrList.Add(Global.dateStrDecimal);

                    }
                    if ((i % 256) == 0)
                    {
                        Global.dateTimeTrasactionStr = Global.readDataAllSplit[i].Substring(0, Global.readDataAllSplit[i].Length - 1);
                        Global.dateTimeTrasactionList.Add(Global.dateTimeTrasactionStr);
                    }
                    else
                    {
                        Global.readingStr = Global.readDataAllSplit[i].Substring(0, Global.readDataAllSplit[i].Length - 1);

                        Global.tempStr1 = Global.readingStr.Substring(0, 2);
                        Global.tempStr2 = Global.readingStr.Substring(2, 2);
                        Global.humidiStr1 = Global.readingStr.Substring(4, 2);
                        Global.humidiStr2 = Global.readingStr.Substring(6, 2);

                        Global.tempStrSwap = string.Concat(Global.tempStr2, Global.tempStr1);
                        Global.humidiStrSwap = string.Concat(Global.humidiStr2, Global.humidiStr1);

                        int tempStrDecimal = Convert.ToInt32(Global.tempStrSwap, 16);
                        int humidiStrDecimal = Convert.ToInt32(Global.humidiStrSwap, 16);

                        Global.temperature = Math.Round((Global.tempFormValue1 + (Global.tempFormValue2 * (Convert.ToDouble(tempStrDecimal) / 65536))), 1);
                        Global.humidity = Math.Round((Global.humidiFormValue1 + (Global.humidiFormValue2 * (Convert.ToDouble(humidiStrDecimal) / 65536))), 1);
                        Global.temperatureList.Add(Global.temperature.ToString());
                        Global.humidity = Math.Abs(Global.humidity);
                        Global.humidityList.Add(Global.humidity.ToString());


                        Global.timeStrMinInterval = Global.timeStrMinInterval + Convert.ToInt32(Global.intervalStrDecimal);
                        Global.timeIntervalMaxCount = Global.timeStrMinInterval;

                        if (Global.timeStrMinInterval == 60)
                        {
                            Global.timeStrMinInterval = 0;
                            Global.timeStrDecimalHr = Global.timeStrDecimalHr + 1;
                        }

                        if (Global.timeStrMinInterval >= 60)
                        {
                            for (int k = 0; k <= Global.timeIntervalMaxCount; k++)
                            {
                                if (k == 60)
                                {
                                    Global.timeStrMinInterval = 0;
                                    if (Global.timeStrDecimalHr < 23)
                                    {
                                        Global.timeStrDecimalHr = Global.timeStrDecimalHr + 1;
                                    }
                                    else
                                    {
                                        Global.timeStrDecimalHr = 0;
                                        Global.dateStrDecimalDay = Global.dateStrDecimalDay + 1;
                                    }
                                    Global.timeIntervalMaxCount = Global.timeIntervalMaxCount - 60;
                                    k = 0;
                                }
                                else
                                {
                                    Global.timeStrMinInterval = k;
                                }

                            }
                        }
                        if (Global.timeStrDecimalHr == 24 && Global.timeStrMinInterval == 00)
                        {
                            Global.timeStrDecimalHr = 0;
                            Global.timeStrDecimalMin = 0;
                            Global.dateStrDecimalDay = Global.dateStrDecimalDay + 1;
                        }

                        if (Global.dateStrDecimalMonth == 1 || Global.dateStrDecimalMonth == 3 || Global.dateStrDecimalMonth == 5 || Global.dateStrDecimalMonth == 7 || Global.dateStrDecimalMonth == 8 || Global.dateStrDecimalMonth == 10)
                        {
                            if (Global.dateStrDecimalDay == 32 || Global.dateStrDecimalDay >= 32)
                            {
                                Global.dateStrDecimalDay = 1;
                                Global.dateStrDecimalMonth = Global.dateStrDecimalMonth + 1;
                            }

                        }

                        if (Global.dateStrDecimalYear == 20 || (Global.dateStrDecimalYear % 4) == 0)
                        {
                            if (Global.dateStrDecimalMonth == 2)
                            {
                                if (Global.dateStrDecimalDay == 30 || Global.dateStrDecimalDay >= 30)
                                {
                                    Global.dateStrDecimalDay = 1;
                                    Global.dateStrDecimalMonth = Global.dateStrDecimalMonth + 1;
                                }
                            }

                        }

                        else
                        {
                            if (Global.dateStrDecimalMonth == 2)
                            {
                                if (Global.dateStrDecimalDay == 29 || Global.dateStrDecimalDay >= 29)
                                {
                                    Global.dateStrDecimalDay = 1;
                                    Global.dateStrDecimalMonth = Global.dateStrDecimalMonth + 1;
                                }
                            }

                        }

                        if (Global.dateStrDecimalMonth == 12 || Global.dateStrDecimalMonth >= 12)
                        {
                            if (Global.dateStrDecimalDay == 32 || Global.dateStrDecimalDay >= 32)
                            {
                                Global.dateStrDecimalDay = 1;
                                Global.dateStrDecimalMonth = 1;
                                Global.dateStrDecimalYear = Global.dateStrDecimalYear + 1;
                            }
                        }

                        if (Global.dateStrDecimalMonth == 4 || Global.dateStrDecimalMonth == 6 || Global.dateStrDecimalMonth == 9 || Global.dateStrDecimalMonth == 11)
                        {
                            if (Global.dateStrDecimalDay == 31 || Global.dateStrDecimalDay >= 31)
                            {
                                Global.dateStrDecimalDay = 1;
                                Global.dateStrDecimalMonth = Global.dateStrDecimalMonth + 1;
                            }
                        }


                        Global.dateStrDecimal = Global.dateStrDecimalDay.ToString("00") + "-" + Global.dateStrDecimalMonth.ToString("00") + "-" + "20" + Global.dateStrDecimalYear;
                        Global.dateStrList.Add(Global.dateStrDecimal);

                        Global.timeStrDecimalInter = Global.timeStrDecimalHr.ToString("00") + " : " + Global.timeStrMinInterval.ToString("00");
                        Global.timeStrList.Add(Global.timeStrDecimalInter);
                    }

                }

            }
            catch (Exception e)
            {
                FileLog.ErrorLog(e.Message + e.StackTrace); //shubham
            }

        }

        /// <summary>
        /// Read the CFG text file which contains all configuration of device.
        /// </summary>
        public void ReadCFGFile()
        {
            try
            {
                string lineW = File.ReadLines(Global.filePathCFG).ElementAt(0);
                string lineC = File.ReadLines(Global.filePathCFG).ElementAt(1);
                string lineD = File.ReadLines(Global.filePathCFG).ElementAt(2);
                string lineP = File.ReadLines(Global.filePathCFG).ElementAt(3);
                string lineX = File.ReadLines(Global.filePathCFG).ElementAt(4);
                string lineR = File.ReadLines(Global.filePathCFG).ElementAt(5);

                Global.model = lineW;

                #region W
                try
                {
                    Global.responseDevName = lineW;
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);// shubham
                }
                #endregion

                #region C
                try
                {

                    Global.responseProgramLog = lineC;

                    #region StartDateTime

                    Global.dateTimeStr = Global.responseProgramLog.Substring(0, 8);

                    Global.timeStr1 = Global.dateTimeStr.Substring(0, 2);
                    Global.timeStr2 = Global.dateTimeStr.Substring(2, 2);
                    Global.dateStr1 = Global.dateTimeStr.Substring(4, 2);
                    Global.dateStr2 = Global.dateTimeStr.Substring(6, 2);

                    Global.timeStrSwap = string.Concat(Global.timeStr2, Global.timeStr1);
                    Global.dateStrSwap = string.Concat(Global.dateStr2, Global.dateStr1);

                    string timeStrBinary = Convert.ToString(Convert.ToInt64(Global.timeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                    string dateStrBinary = Convert.ToString(Convert.ToInt64(Global.dateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                    int timeStrDecimalHr = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 11, 5), 2);
                    int timeStrDecimalMin = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 6, 6), 2);
                    Global.timeStrDecimalC = timeStrDecimalHr.ToString("00") + ":" + timeStrDecimalMin.ToString("00");

                    int dateStrDecimalDay = Convert.ToInt32(dateStrBinary.Substring(0, 5), 2);
                    int dateStrDecimalMonth = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 11, 4), 2);
                    int dateStrDecimalYear = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 7, 7), 2);
                    Global.dateStrDecimalC = dateStrDecimalDay.ToString("00") + "-" + dateStrDecimalMonth.ToString("00") + "-" + "20" + dateStrDecimalYear.ToString("00");

                    #endregion

                    #region LogInterval

                    Global.intervalStr = Global.responseProgramLog.Substring(8, 4);

                    Global.intervalStr1 = Global.intervalStr.Substring(0, 2);
                    Global.intervalStr2 = Global.intervalStr.Substring(2, 2);
                    Global.intervalStrSwap = string.Concat(Global.intervalStr2 + Global.intervalStr1);

                    Global.intervalStrDecimal = Convert.ToString(Convert.ToInt32(Global.intervalStrSwap.Substring(0, 4), 16));

                    #endregion

                    #region Format,Event,Dummybyte
                    Global.formatStr = Global.responseProgramLog.Substring(12, 2);
                    Global.eventStr = Global.responseProgramLog.Substring(14, 2);
                    Global.dummyByteStr = Global.responseProgramLog.Substring(16, 2);
                    #endregion

                    #region StopDateTime

                    Global.StDateTimeStr = Global.responseProgramLog.Substring(Global.responseProgramLog.Length - 14, 8);

                    Global.StTimeStr1 = Global.StDateTimeStr.Substring(0, 2);
                    Global.StTimeStr2 = Global.StDateTimeStr.Substring(2, 2);
                    Global.StDateStr1 = Global.StDateTimeStr.Substring(4, 2);
                    Global.StDateStr2 = Global.StDateTimeStr.Substring(6, 2);

                    Global.StTimeStrSwap = string.Concat(Global.StTimeStr2, Global.StTimeStr1);
                    Global.StDateStrSwap = string.Concat(Global.StDateStr2, Global.StDateStr1);

                    string StTimeStrBinary = Convert.ToString(Convert.ToInt64(Global.StTimeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                    string StDateStrBinary = Convert.ToString(Convert.ToInt64(Global.StDateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                    int StTimeStrDecimalHr = Convert.ToInt32(StTimeStrBinary.Substring(StTimeStrBinary.Length - 11, 5), 2);
                    int StTimeStrDecimalMin = Convert.ToInt32(StTimeStrBinary.Substring(StTimeStrBinary.Length - 6, 6), 2);
                    Global.StTimeStrDecimalC = StTimeStrDecimalHr.ToString("00") + ":" + StTimeStrDecimalMin.ToString("00");

                    int StDateStrDecimalDay = Convert.ToInt32(StDateStrBinary.Substring(0, 5), 2);
                    int StDateStrDecimalMonth = Convert.ToInt32(StDateStrBinary.Substring(StDateStrBinary.Length - 11, 4), 2);
                    int StDateStrDecimalYear = Convert.ToInt32(StDateStrBinary.Substring(StDateStrBinary.Length - 7, 7), 2);
                    Global.StDateStrDecimalC = StDateStrDecimalDay.ToString("00") + "-" + StDateStrDecimalMonth.ToString("00") + "-" + "20" + StDateStrDecimalYear.ToString("00");

                    #endregion

                    #region TotalTransactionStoreCount

                    Global.transCountstr = Global.responseProgramLog.Substring(Global.responseProgramLog.Length - 6, 4);

                    Global.transCountstr1 = Global.transCountstr.Substring(0, 2);
                    Global.transCountstr2 = Global.transCountstr.Substring(2, 2);

                    Global.transCountstrSwap = string.Concat(Global.transCountstr2, Global.transCountstr1);
                    Global.transCountstrDecimal = Convert.ToString(Convert.ToInt32(Global.transCountstrSwap.Substring(0, 4), 16));

                    #endregion

                    #region DeviceLogStatus
                    Global.devLogStatusStr = Global.responseProgramLog.Substring(Global.responseProgramLog.Length - 2, 2);
                    #endregion

                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);//shubham
                }
                #endregion

                #region D
                try
                {
                    Global.responseProgLogger = lineD;

                    #region DisplayOnTime

                    string dispTimeStr = Global.responseProgLogger.Substring(0, 2);
                    string dispTimeStrBinary1 = Convert.ToString(Convert.ToInt64(dispTimeStr, 16), 2).PadLeft(8, '0');
                    string dispTimeStrBinary2 = dispTimeStrBinary1.Substring(dispTimeStrBinary1.Length - 8, 6).PadLeft(8, '0');
                    int dispTimeStrDecimal1 = Convert.ToInt32(dispTimeStrBinary2.Substring(0, 4), 2);
                    int dispTimeStrDecimal2 = Convert.ToInt32(dispTimeStrBinary2.Substring(dispTimeStrBinary2.Length - 4, 4), 2);
                    Global.dispTimeStrDecimal = Convert.ToString(dispTimeStrDecimal1 + dispTimeStrDecimal2);

                    #endregion

                    #region StartDateTime

                    Global.dateTimeStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 52, 8);

                    Global.timeStr1 = Global.dateTimeStr.Substring(0, 2);
                    Global.timeStr2 = Global.dateTimeStr.Substring(2, 2);
                    Global.dateStr1 = Global.dateTimeStr.Substring(4, 2);
                    Global.dateStr2 = Global.dateTimeStr.Substring(6, 2);

                    Global.timeStrSwap = string.Concat(Global.timeStr2, Global.timeStr1);
                    Global.dateStrSwap = string.Concat(Global.dateStr2, Global.dateStr1);

                    string timeStrBinary = Convert.ToString(Convert.ToInt64(Global.timeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                    string dateStrBinary = Convert.ToString(Convert.ToInt64(Global.dateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                    int timeStrDecimalHr = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 11, 5), 2);
                    int timeStrDecimalMin = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 6, 6), 2);
                    Global.timeStrDecimalD = timeStrDecimalHr.ToString("00") + ":" + timeStrDecimalMin.ToString("00");

                    int dateStrDecimalDay = Convert.ToInt32(dateStrBinary.Substring(0, 5), 2);
                    int dateStrDecimalMonth = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 11, 4), 2);
                    int dateStrDecimalYear = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 7, 7), 2);
                    Global.dateStrDecimalD = dateStrDecimalDay.ToString("00") + "-" + dateStrDecimalMonth.ToString("00") + "-" + "20" + dateStrDecimalYear.ToString("00");

                    #endregion

                    #region StopDateTime

                    Global.StDateTimeStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 44, 8);

                    Global.StTimeStr1 = Global.StDateTimeStr.Substring(0, 2);
                    Global.StTimeStr2 = Global.StDateTimeStr.Substring(2, 2);
                    Global.StDateStr1 = Global.StDateTimeStr.Substring(4, 2);
                    Global.StDateStr2 = Global.StDateTimeStr.Substring(6, 2);

                    Global.StTimeStrSwap = string.Concat(Global.StTimeStr2, Global.StTimeStr1);
                    Global.StDateStrSwap = string.Concat(Global.StDateStr2, Global.StDateStr1);

                    string StTimeStrBinary = Convert.ToString(Convert.ToInt64(Global.StTimeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                    string StDateStrBinary = Convert.ToString(Convert.ToInt64(Global.StDateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                    int StTimeStrDecimalHr = Convert.ToInt32(StTimeStrBinary.Substring(StTimeStrBinary.Length - 11, 5), 2);
                    int StTimeStrDecimalMin = Convert.ToInt32(StTimeStrBinary.Substring(StTimeStrBinary.Length - 6, 6), 2);
                    Global.StTimeStrDecimalD = StTimeStrDecimalHr.ToString("00") + ":" + StTimeStrDecimalMin.ToString("00");

                    int StDateStrDecimalDay = Convert.ToInt32(StDateStrBinary.Substring(0, 5), 2);
                    int StDateStrDecimalMonth = Convert.ToInt32(StDateStrBinary.Substring(StDateStrBinary.Length - 11, 4), 2);
                    int StDateStrDecimalYear = Convert.ToInt32(StDateStrBinary.Substring(StDateStrBinary.Length - 7, 7), 2);
                    Global.StDateStrDecimalD = StDateStrDecimalDay.ToString("00") + "-" + StDateStrDecimalMonth.ToString("00") + "-" + "20" + StDateStrDecimalYear.ToString("00");

                    #endregion

                    #region LogDelay

                    string delayStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 36, 4);
                    string delayStr1 = delayStr.Substring(0, 2);
                    string delayStr2 = delayStr.Substring(2, 2);
                    string delayStrSwap = string.Concat(delayStr2, delayStr1);
                    Global.delayStrDecimal = Convert.ToString(Convert.ToInt32(delayStrSwap, 16));

                    #endregion

                    #region tempHighAlarm

                    string tempHighStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 32, 4);
                    string tempHighStr1 = tempHighStr.Substring(0, 2);
                    string tempHighStr2 = tempHighStr.Substring(2, 2);
                    string tempHighStrSwap = string.Concat(tempHighStr2, tempHighStr1);
                    Global.tempHighStrDecimal = Convert.ToString(Convert.ToUInt32(tempHighStrSwap, 16));

                    #endregion

                    #region tempLowAlarm

                    string tempLowStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 28, 4);
                    string tempLowStr1 = tempLowStr.Substring(0, 2);
                    string tempLowStr2 = tempLowStr.Substring(2, 2);
                    string tempLowStrSwap = string.Concat(tempLowStr2, tempLowStr1);
                    Global.tempLowStrDecimal = Convert.ToString(Convert.ToUInt32(tempLowStrSwap, 16));

                    #endregion

                    #region humiHighAlarm

                    string humiHighStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 24, 2);
                    Global.humiHighStrDecimal = Convert.ToString(Convert.ToInt32(humiHighStr, 16));

                    #endregion

                    #region humiLowAlarm

                    string humiLowStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 22, 2);
                    Global.humiLowStrDecimal = Convert.ToString(Convert.ToInt32(humiLowStr, 16));

                    #endregion

                    #region tempAlarmSign

                    string tempsignStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 20, 2);
                    if (tempsignStr == "00")
                    {
                        Global.tempHighStrDecimal = Global.tempHighStrDecimal;
                        Global.tempLowStrDecimal = Global.tempLowStrDecimal;
                    }
                    if (tempsignStr == "01")
                    {
                        Global.tempHighStrDecimal = Global.tempHighStrDecimal;
                        Global.tempLowStrDecimal = "-" + Global.tempLowStrDecimal;
                    }
                    if (tempsignStr == "02")
                    {
                        Global.tempHighStrDecimal = "-" + Global.tempHighStrDecimal;
                        Global.tempLowStrDecimal = "-" + Global.tempLowStrDecimal;
                    }
                    #endregion

                    #region deviceStartMode

                    Global.deviceModeStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 18, 2);

                    #endregion

                    #region deviceName

                    Global.deviceNameStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 16, 8);

                    #endregion

                    #region serialNo

                    Global.serialNoStr = Global.responseProgLogger.Substring(Global.responseProgLogger.Length - 8, 8);

                    #endregion

                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);//shubham
                }
                #endregion

                #region P
                try
                {
                    Global.responseLoggerTime = lineP;

                    Global.dateTimeStr = Global.responseLoggerTime;
                    Global.timeStr1 = Global.dateTimeStr.Substring(0, 2);
                    Global.timeStr2 = Global.dateTimeStr.Substring(2, 2);
                    Global.dateStr1 = Global.dateTimeStr.Substring(4, 2);
                    Global.dateStr2 = Global.dateTimeStr.Substring(6, 2);

                    Global.timeStrSwap = string.Concat(Global.timeStr2, Global.timeStr1);
                    Global.dateStrSwap = string.Concat(Global.dateStr2, Global.dateStr1);

                    string timeStrBinary = Convert.ToString(Convert.ToInt64(Global.timeStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');
                    string dateStrBinary = Convert.ToString(Convert.ToInt64(Global.dateStrSwap.Substring(0, 4), 16), 2).PadLeft(16, '0');

                    Global.timeStrDecimalHr = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 11, 5), 2);
                    Global.timeStrDecimalMin = Convert.ToInt32(timeStrBinary.Substring(timeStrBinary.Length - 6, 6), 2);
                    Global.timeStrDecimalP = Global.timeStrDecimalHr.ToString("00") + " : " + Global.timeStrDecimalMin.ToString("00");

                    int dateStrDecimalDay = Convert.ToInt32(dateStrBinary.Substring(0, 5), 2);
                    int dateStrDecimalMonth = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 11, 4), 2);
                    int dateStrDecimalYear = Convert.ToInt32(dateStrBinary.Substring(dateStrBinary.Length - 7, 7), 2);
                    Global.dateStrDecimalP = dateStrDecimalDay.ToString("00") + "-" + dateStrDecimalMonth.ToString("00") + "-" + "20" + dateStrDecimalYear.ToString("00");

                    Global.loggerDateTime = Global.dateStrDecimalP + "  " + Global.timeStrDecimalP;

                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);//shubham
                }
                #endregion

                #region X
                try
                {
                    Global.responseMinMaxRead = lineX;

                    string tempMaxStr = Global.responseMinMaxRead.Substring(0, 4);
                    string tempMinStr = Global.responseMinMaxRead.Substring(Global.responseMinMaxRead.Length - 12, 4);
                    string humiMaxStr = Global.responseMinMaxRead.Substring(Global.responseMinMaxRead.Length - 8, 4);
                    string humiMinStr = Global.responseMinMaxRead.Substring(Global.responseMinMaxRead.Length - 4, 4);

                    string tempMaxStr1 = tempMaxStr.Substring(0, 2);
                    string tempMaxStr2 = tempMaxStr.Substring(tempMaxStr.Length - 2, 2);
                    string tempMinStr1 = tempMinStr.Substring(0, 2);
                    string tempMinStr2 = tempMinStr.Substring(tempMinStr.Length - 2, 2);

                    string humiMaxStr1 = humiMaxStr.Substring(0, 2);
                    string humiMaxStr2 = humiMaxStr.Substring(humiMaxStr.Length - 2, 2);
                    string humiMinStr1 = humiMinStr.Substring(0, 2);
                    string humiMinStr2 = humiMinStr.Substring(humiMinStr.Length - 2, 2);

                    string tempMaxStrSwap = string.Concat(tempMaxStr2, tempMaxStr1);
                    string tempMinStrSwap = string.Concat(tempMinStr2, tempMinStr1);
                    string humiMaxStrSwap = string.Concat(humiMaxStr2, humiMaxStr1);
                    string humiMinStrSwap = string.Concat(humiMinStr2, humiMinStr1);

                    int tempMaxStrDecimal = Convert.ToInt32(tempMaxStrSwap, 16);
                    int tempMinStrDecimal = Convert.ToInt32(tempMinStrSwap, 16);
                    int humiMaxStrDecimal = Convert.ToInt32(humiMaxStrSwap, 16);
                    int humiMinStrDecimal = Convert.ToInt32(humiMinStrSwap, 16);

                    Global.maxTemperature = Math.Round((Global.tempFormValue1 + (Global.tempFormValue2 * (Convert.ToDouble(tempMaxStrDecimal) / 65536))), 1);
                    Global.minTemperature = Math.Round((Global.tempFormValue1 + (Global.tempFormValue2 * (Convert.ToDouble(tempMinStrDecimal) / 65536))), 1);
                    Global.maxHumidity = Math.Round((Global.humidiFormValue1 + (Global.humidiFormValue2 * (Convert.ToDouble(humiMaxStrDecimal) / 65536))), 1);
                    Global.minHumidity = Math.Round((Global.humidiFormValue1 + (Global.humidiFormValue2 * (Convert.ToDouble(humiMinStrDecimal) / 65536))), 1);

                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace);// shubham
                }
                #endregion

                #region R
                try
                {
                    Global.responseLEDValue = lineR;
                }
                catch (Exception ex)
                {
                    FileLog.ErrorLog(ex.Message + ex.StackTrace); //shubham
                }
                #endregion

            }

            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace); //shubham
            }
        }

        //
        #endregion

    }
}
