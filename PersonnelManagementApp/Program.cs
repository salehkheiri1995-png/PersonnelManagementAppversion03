using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    internal static class Program
    {
        /// <summary>
        /// نقطه ورود برنامه.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                Application.SetDefaultFont(new Font("Tahoma", 9F));
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("Program.SetDefaultFont", ex);
            }

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += OnThreadException;
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;

            // پیش‌نیاز پایگاه داده: در صورت نبود ACE، راهنمای نصب نمایش داده می‌شود.
            // برنامه در هر حال اجرا می‌شود تا کاربر بتواند تنظیمات و مسیر فایل را بررسی کند.
            DatabaseProvider.CheckProvider(null, offerInstall: true);

            Application.Run(new MainForm());
        }

        private static void OnThreadException(object sender, ThreadExceptionEventArgs e)
        {
            ShowFatalError(e.Exception);
        }

        private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            ShowFatalError(e.ExceptionObject as Exception);
        }

        private static void ShowFatalError(Exception? exception)
        {
            string message = exception?.Message ?? "خطای ناشناخته";

            try
            {
                ErrorLogger.Log("Program.UnhandledException", exception);
            }
            catch
            {
                // ثبت خطا نباید خودش باعث خطای جدید شود.
            }

            MessageBox.Show(
                "خطای پیش‌بینی‌نشده‌ای رخ داد:\n\n" + message +
                "\n\nجزئیات در پوشه Logs کنار برنامه ذخیره شد.",
                "خطای برنامه", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
