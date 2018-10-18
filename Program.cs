using System;
using System.Collections.Generic; 
using System.ServiceProcess;
using System.Text; 
using System.Windows.Forms;

namespace FileToImgService
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
            { 
                new ServiceDemo() 
            };
            ServiceBase.Run(ServicesToRun);
        }
        //调试时使用下面的代码
        //[STAThread]
        //static void Main()
        //{
        //    Application.EnableVisualStyles();
        //    Application.SetCompatibleTextRenderingDefault(false);
        //    Application.Run(new Form1());
        //}
    }
}
