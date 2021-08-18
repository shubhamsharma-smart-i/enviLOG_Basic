using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Windows.Forms;
using PDDL.Common;

namespace PDDL.Interface
{
    public class FileLog
    {

        public static void ErrorLog(string sMessage)
        {
            StreamWriter objSw = null;
            try
            {
                string sFolderName = System.IO.Path.GetTempPath() + @"enviLOG Basic\ErrorLogs\";
                if (!Directory.Exists(sFolderName))
                    Directory.CreateDirectory(sFolderName);
                string sFilePath = sFolderName + "Error " + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
                objSw = new StreamWriter(sFilePath, true);
                objSw.WriteLine(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + " " + sMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                FileLog.ErrorLog(ex.Message + ex.StackTrace);
            }
            finally
            {
                if (objSw != null)
                {
                    objSw.Flush();
                    objSw.Dispose();
                }
            }
        }

        public static void CommandLog(string sMessage)
        {
            StreamWriter objSw = null;
            try
            {
                //string sFolderName = Application.StartupPath + @"\CommandLogs\";
                string sFolderName = System.IO.Path.GetTempPath() + @"enviLOG Basic\CommandLogs\";
                // string tempPath =
                if (!Directory.Exists(sFolderName))
                    Directory.CreateDirectory(sFolderName);

                string sFilePath = sFolderName + "Command " + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
                objSw = new StreamWriter(sFilePath, true);
                objSw.WriteLine(DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "  " + sMessage + Environment.NewLine);

            }
            catch (Exception ex)
            {
                FileLog.CommandLog(ex.Message + ex.StackTrace);
            }
            finally
            {
                if (objSw != null)
                {
                    objSw.Flush();
                    objSw.Dispose();
                }
            }


        }

        public static void DataLog(string sMessage)
        {
            StreamWriter objSw = null;
            try
            {
                string sFolderName = System.IO.Path.GetTempPath() + @"enviLOG Basic\DataLogs\";
                if (!Directory.Exists(sFolderName))
                    Directory.CreateDirectory(sFolderName);

                for (int i = 0; i < Global.filePathCFGList.Count; i++)
                {
                    if (Global.filePathCFGList[i].Contains(Global.serialNoStr))
                    {
                        Global.filePathCFGList.RemoveAt(i);
                        Global.filePathLOGList.RemoveAt(i);

                        break;

                    }

                }


                Global.dFolderName = sFolderName + Global.serialNoStr + "_" + DateTime.Now.ToString("dd-MM-yyyy HH.mm");
                if (!Directory.Exists(Global.dFolderName))
                    Directory.CreateDirectory(Global.dFolderName);

                Global.filePathCFG = Global.dFolderName + @"\CFG" + ".txt";
                Global.filePathLOG = Global.dFolderName + @"\LOG" + ".txt";
            }
            catch (Exception ex)
            {
                FileLog.DataLog(ex.Message + ex.StackTrace);
            }
            finally
            {
                if (objSw != null)
                {
                    objSw.Flush();
                    objSw.Dispose();
                }
            }
        }
    }
}





