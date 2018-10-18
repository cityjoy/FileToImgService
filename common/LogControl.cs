///***********************************
///******日志操作类******************
///**********************************
using System;
using System.Configuration;
using System.IO;


namespace FileToImgService
{
    public class LogControl
    {
        private static string strLogMode = ConfigurationSettings.AppSettings["LogMode"];
        private static string strLogFilePath = ConfigurationSettings.AppSettings["LogFilePath"];
        private static bool blnLogInfo = bool.Parse(ConfigurationSettings.AppSettings["LogInfoData"].ToString());
        private static double dblMaxLogFileAge = double.Parse(ConfigurationSettings.AppSettings["MaxLogFileAge"].ToString());

        public LogControl()
        {
        }

        public static void LogInfo(string strData)
        {
            if (blnLogInfo)
            {
                switch (strLogMode.ToUpper())
                {
                    case "FILE":
                        WriteToFile(strLogFilePath, "[INFO] [" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff") + "] : " + strData);
                        break;
                    case "CONSOLE":
                        Console.WriteLine("[INFO] [" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff") + "] : " + strData);
                        break;
                    case "APPLOGGER":
                        break;
                }
            }
        }

        public static void LogException(string strData, string strSource)
        {
            switch (strLogMode.ToUpper())
            {
                case "FILE":
                    WriteToFile(strLogFilePath, "[ERROR] [" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff") + "] : " + strSource + " - " + strData);
                    break;
                case "CONSOLE":
                    Console.WriteLine("[ERROR] [" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss:ffff") + "] : " + strSource + " - " + strData);
                    break;
                case "APPLOGGER":
                    break;
            }
        }

        private static bool WriteToFile(string strFile, string strData)
        {
            if (!Directory.Exists(strFile))
                Directory.CreateDirectory(strFile);
            DateTime dtNow = DateTime.Now;
            StreamWriter sw = new StreamWriter(strFile + "\\Log_" + dtNow.ToString("yyyyMMddHH") + ".txt", true);
            sw.WriteLine(strData);
            sw.Close();
            return true;
        }

        //---------------------------------------------------
        //
        //---------------------------------------------------
        public static bool DeleteAgedLogs()
        {
            try
            {
                string[] arrLogFiles = Directory.GetFiles(strLogFilePath, "*.txt");             
                for (int i = 0; i < arrLogFiles.Length; i++)
                {
                    DateTime dtFileDate = File.GetLastWriteTime(arrLogFiles[i]);
                    if (dtFileDate < DateTime.Now.Date.AddDays(0 - dblMaxLogFileAge))
                        File.Delete(arrLogFiles[i]);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

    }
}