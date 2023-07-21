using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FindExcelContent
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main( string[] args )
        {
            if(args.Length > 0 )
            {
                Logic.DirPath = args[0];
                if(args.Length > 1 )
                {
                    Logic.FindNameList.AddRange(args);
                    Logic.FindNameList.RemoveAt(0);
                }
                Logic.IsExternalCall = true;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FindExcelContent());
        }
    }
}
