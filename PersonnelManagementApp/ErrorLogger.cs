using System;
using System.IO;
using System.Text;

namespace PersonnelManagementApp
{
    /// <summary>
    /// ثبت خطاهای برنامه در یک فایل متنی قابل بررسی.
    /// </summary>
    public static class ErrorLogger
    {
        private static readonly object SyncRoot = new object();
        private static readonly string LogDirectory = ResolveLogDirectory();

        public static string LogFolder => LogDirectory;

        public static string LogFilePath => Path.Combine(LogDirectory, "error.log");

        private static string ResolveLogDirectory()
        {
            try
            {
                string local = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PersonnelManagementApp",
                    "Logs");

                Directory.CreateDirectory(local);
                return local;
            }
            catch
            {
                try
                {
                    string fallback = Path.Combine(Path.GetTempPath(), "PersonnelManagementApp", "Logs");
                    Directory.CreateDirectory(fallback);
                    return fallback;
                }
                catch
                {
                    return Path.GetTempPath();
                }
            }
        }

        public static void Log(string context, Exception? exception)
        {
            if (exception == null)
                return;

            LogMessage(context, exception.ToString());
        }

        public static void LogMessage(string context, string message)
        {
            try
            {
                lock (SyncRoot)
                {
                    Directory.CreateDirectory(LogDirectory);
                    string path = LogFilePath;

                    // جلوگیری از بزرگ شدن بی‌رویه فایل لاگ
                    if (File.Exists(path) && new FileInfo(path).Length > 5 * 1024 * 1024)
                    {
                        string backup = Path.Combine(LogDirectory, "error.old.log");
                        try
                        {
                            File.Copy(path, backup, true);
                            File.Delete(path);
                        }
                        catch
                        {
                            // اگر چرخش لاگ ممکن نبود، نوشتن ادامه پیدا می‌کند.
                        }
                    }

                    var sb = new StringBuilder();
                    sb.AppendLine("────────────────────────────────────────");
                    sb.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}");
                    sb.AppendLine(message);
                    sb.AppendLine();

                    File.AppendAllText(path, sb.ToString(), new UTF8Encoding(false));
                }
            }
            catch
            {
                // لاگ نباید خودش باعث خطای برنامه شود.
            }
        }
    }
}
