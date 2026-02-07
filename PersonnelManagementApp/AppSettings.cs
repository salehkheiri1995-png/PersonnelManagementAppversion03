using System;
using System.IO;
using System.Configuration;

namespace PersonnelManagementApp
{
    /// <summary>
    /// مدیریت تنظیمات برنامه
    /// </summary>
    public static class AppSettings
    {
        private static readonly string ConfigFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, 
            "AppConfig.ini"
        );

        // مسیر پیش‌فرض دیتابیس
        private static readonly string DefaultDatabasePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "PersonnelDatabase.accdb"
        );

        // مسیر پیش‌فرض پوشه عکس‌ها
        private static readonly string DefaultPhotosFolder = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "PersonnelPhotos"
        );

        /// <summary>
        /// مسیر دیتابیس
        /// </summary>
        public static string DatabasePath
        {
            get
            {
                string path = ReadSetting("DatabasePath");
                if (string.IsNullOrEmpty(path) || !File.Exists(path))
                {
                    path = DefaultDatabasePath;
                    WriteSetting("DatabasePath", path);
                }
                return path;
            }
            set
            {
                WriteSetting("DatabasePath", value);
            }
        }

        /// <summary>
        /// مسیر پوشه عکس‌ها
        /// </summary>
        public static string PhotosFolder
        {
            get
            {
                string path = ReadSetting("PhotosFolder");
                if (string.IsNullOrEmpty(path))
                {
                    path = DefaultPhotosFolder;
                    WriteSetting("PhotosFolder", path);
                }

                // ایجاد پوشه اگر وجود ندارد
                if (!Directory.Exists(path))
                {
                    try
                    {
                        Directory.CreateDirectory(path);
                    }
                    catch
                    {
                        // اگر نتوانست بسازد، از مسیر پیش‌فرض استفاده کن
                        path = DefaultPhotosFolder;
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }
                    }
                }

                return path;
            }
            set
            {
                WriteSetting("PhotosFolder", value);

                // ایجاد پوشه جدید
                if (!Directory.Exists(value))
                {
                    Directory.CreateDirectory(value);
                }
            }
        }

        /// <summary>
        /// خواندن تنظیم از فایل
        /// </summary>
        private static string ReadSetting(string key)
        {
            try
            {
                if (!File.Exists(ConfigFilePath))
                    return string.Empty;

                string[] lines = File.ReadAllLines(ConfigFilePath);
                foreach (string line in lines)
                {
                    if (line.StartsWith(key + "="))
                    {
                        return line.Substring(key.Length + 1);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در خواندن تنظیمات: {ex.Message}");
            }

            return string.Empty;
        }

        /// <summary>
        /// نوشتن تنظیم در فایل
        /// </summary>
        private static void WriteSetting(string key, string value)
        {
            try
            {
                string[] lines = File.Exists(ConfigFilePath) ? File.ReadAllLines(ConfigFilePath) : new string[0];
                bool found = false;

                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].StartsWith(key + "="))
                    {
                        lines[i] = key + "=" + value;
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    // اضافه خط جدید
                    Array.Resize(ref lines, lines.Length + 1);
                    lines[lines.Length - 1] = key + "=" + value;
                }

                File.WriteAllLines(ConfigFilePath, lines);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در نوشتن تنظیمات: {ex.Message}");
            }
        }

        /// <summary>
        /// برگرداندن تنظیمات به حالت پیش‌فرض
        /// </summary>
        public static void ResetToDefaults()
        {
            DatabasePath = DefaultDatabasePath;
            PhotosFolder = DefaultPhotosFolder;
        }

        /// <summary>
        /// بررسی وجود فایل دیتابیس
        /// </summary>
        public static bool DatabaseExists()
        {
            return File.Exists(DatabasePath);
        }
    }
}