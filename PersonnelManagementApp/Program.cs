using System;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // مقداردهی اولیه EPPlus License - حتماً در ابتدای برنامه!
            EPPlusLicenseInitializer.Initialize();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}