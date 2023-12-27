using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FindExcelContent
{
    public static class Logic
    {
        private static string[] ColNameArr = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        private static Regex m_regexName = new Regex(@"^~\$");
        private static string m_dirPath;
        public static string DirPath
        {
            get
            {
                if( m_dirPath == null )
                {
                    m_dirPath = CfgFile.ReadString("", "DirPath");
                }
                return m_dirPath;
            }
            set
            {
                m_dirPath = value;
                CfgFile.WriteString("", "DirPath" , value);
            }
        }
        public static List<string> FindNameList = new List<string>();
        public static bool IsExternalCall = false;//是外部调用？
        private static FileOP m_cfgFile;
        /// <summary>
        /// 配置文件
        /// </summary>
        public static FileOP CfgFile
        {
            get
            {
                if (m_cfgFile == null)
                {
                    m_cfgFile = new FileOP();
                    m_cfgFile.CreateFile(Application.StartupPath + "\\FindExcelContent.ini");
                }
                return m_cfgFile;
            }
        }

        public static bool Check()
        {
            if (string.IsNullOrEmpty(DirPath))
            {
                MessageBox.Show("Excel文件夹路径错误");
                return false;
            }
            return true;
        }
        public static List<string> GetFilePathList()
        {
            List<string> paths = new List<string>();
            var arr = Directory.GetFiles(DirPath, "*.xlsm", SearchOption.TopDirectoryOnly);
            for (int i = 0; i < arr.Length; i++)
            {
                paths.Add(arr[i]);
            }
            List<string> fileList = new List<string>();
            for (int i = 0; i < paths.Count; i++)
            {
                if (!m_regexName.IsMatch(Path.GetFileName(paths[i])))
                {
                    fileList.Add(paths[i]);
                }
            }
            return fileList;
        }
        public static string ArrToString(string[] arr)
        {
            if (arr == null) return "null";
            string s = "";
            for (int i = 0; i < arr.Length; i++)
            {
                s += arr[i] + " ";
            }
            return s;
        }
        public static string ArrToString(List<string> arr)
        {
            if (arr == null) return "null";
            string s = "";
            for (int i = 0; i < arr.Count; i++)
            {
                s += arr[i] + " ";
            }
            return s;
        }
    }
}
