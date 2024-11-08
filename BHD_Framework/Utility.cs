﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace BHD_Framework
{
    public static class Utility
    {
        /// <summary>
        /// Get string in behind equal sign base on your key string
        /// </summary>
        /// <param name="KeyString">The key in front of equal sign</param>
        /// <param name="YourString">Your text</param>
        /// <returns></returns>
        public static string GetValueOf(string KeyString, string YourString) { return GetValueOf(KeyString, "=", YourString); }
        /// <summary>
        /// Get string in behind split sign base on your key string
        /// </summary>
        /// <param name="KeyString">The key string in front of split sign</param>
        /// <param name="SplitSign">The split characters</param>
        /// <param name="YourString">Your text</param>
        /// <returns></returns>
        public static string GetValueOf(string KeyString, string SplitSign, string YourString)
        {
            //your key : your.value.will.be.this
            string result = "";
            int indexSplitSign = YourString.IndexOf(SplitSign);
            if (indexSplitSign < 0) return result;
            string s1 = YourString.Substring(0, indexSplitSign).Trim();
            string s2 = YourString.Substring(indexSplitSign + SplitSign.Length, YourString.Length - indexSplitSign - SplitSign.Length).Trim();
            if (s1 != KeyString) return result;
            result = s2;
            return result;
        }

        [System.Runtime.InteropServices.DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        /// <summary>
        /// Get connect configuration of Mankichi config file
        /// </summary>
        /// <param name="FileFullPath">Path full of configuration file</param>
        /// <returns></returns>
        public static Dictionary<string, string> GetConnectionConfig(string FileFullPath)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            if (!File.Exists(FileFullPath)) return dic;
            try
            {
                StringBuilder strBuilder = new StringBuilder(255);

                int i = GetPrivateProfileString("ConfigConnectDB", "ServerName", "", strBuilder, 255, FileFullPath);
                dic.Add("Server", strBuilder.ToString().Trim());

                i = GetPrivateProfileString("ConfigConnectDB", "DatabaseName", "", strBuilder, 255, FileFullPath);
                dic.Add("Database", strBuilder.ToString().Trim());

                i = GetPrivateProfileString("ConfigConnectDB", "UserName", "", strBuilder, 255, FileFullPath);
                dic.Add("User", Cipher.ToggleMankichi(strBuilder.ToString().Trim()));

                i = GetPrivateProfileString("ConfigConnectDB", "Password", "", strBuilder, 255, FileFullPath);
                dic.Add("Password", Cipher.ToggleMankichi(strBuilder.ToString().Trim()));

                i = GetPrivateProfileString("ConfigConnectDB", "KindConnect", "", strBuilder, 255, FileFullPath);
                dic.Add("KindConnect", strBuilder.ToString().Trim());

                i = GetPrivateProfileString("ConfigConnectDB", "Url", "", strBuilder, 255, FileFullPath);
                dic.Add("Url", strBuilder.ToString().Trim() + "//Utilities.asmx");

                i = GetPrivateProfileString("ConfigConnectDB", "ServerUpdate", "", strBuilder, 255, FileFullPath);
                dic.Add("ServerUpdate", strBuilder.ToString().Trim());
            }
            catch
            {
                //Console.WriteLine("Error occur while reading file!");
            }
            return dic;
        }


        public static string CreateXml(DataTable YourData)
        {
            StringWriter sw = new StringWriter();
            YourData.WriteXml(sw, XmlWriteMode.IgnoreSchema);
            return sw.ToString();
        }
        public static string CreateXml(DataSet YourData)
        {
            StringWriter sw = new StringWriter();
            YourData.WriteXml(sw, XmlWriteMode.IgnoreSchema);
            return sw.ToString();
        }

        private static void WriteToFile(string FilePath, string YourData) { WriteToFile(FilePath, new string[] { YourData }, false); }
        private static void WriteToFile(string FilePath, string []YourData) { WriteToFile(FilePath, YourData, false); }
        private static void WriteToFile(string FilePath, string YourData, bool CreateFileIfNotExists) { WriteToFile(FilePath, new string[] { YourData }, CreateFileIfNotExists); }
        private static void WriteToFile(string FilePath, string[] YourData, bool CreateFileIfNotExists)
        {
            if (YourData == null) return;
            if (YourData.Length == 0) return;
            List<string> lst = new List<string>();
            foreach (string data in YourData)
            {
                if (data == null) continue;
                if (data.Trim() == string.Empty) continue;
                lst.Add(data);
            }
            if (lst.Count == 0) return;
            if (CreateFileIfNotExists)
            {
                try { File.Create(FilePath); } catch { Console.WriteLine("[WriteToFile] Can not create file!"); return; }
            }
            //if (!FilePath.EndsWith(".txt")) return;
            if (File.Exists(FilePath))
            {
                List<string> results = new List<string>();
                string[] lines = File.ReadAllLines(FilePath);
                foreach (string line in lines)
                {
                    results.Add(line);
                }
                foreach (string item in lst)
                {
                    results.Add(item);
                }
                try { File.WriteAllLines(FilePath, results.ToArray()); } catch { Console.WriteLine("[WriteToFile] Can not write to file!"); return; }
                return;
            }
            Console.WriteLine("[WriteToFile] File path is not exist!");
        }

    }


}
