using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XlCodeExtractor
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                Application.Run(new CodeExtractor());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //' GC needs to be called twice in order to get the Finalizers called  
                //' - the first time in, it simply makes a list of what is to be finalized, 
                //' - the second time in, it actually is finalizing.   
                //'   Only then will the object do its automatic ReleaseComObject. 
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.GetTotalMemory(forceFullCollection: true);
            }
        }
    }
}
