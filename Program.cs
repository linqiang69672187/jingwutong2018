using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace JingWuTong
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
           
        }
        //调试
        //public static string sqladb = "server=127.0.0.1;database=JingWuTong;uid=sa;pwd=sa";
        //发布
        public static string sqladb = "server=.;database=JingWuTong;uid=sa;pwd=Tzjj1a6b8c";
        public static string oradb = "";

    }
}
