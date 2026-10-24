using System;
using System.IO;
using System.Linq;
using System.Text;

namespace PersonnelManagementApp
{
    /// <summary>
    /// مدیریت تنظیمات برنامه (مسیر پایگاه‌داده، پوشه عکس‌ها و ...).
    /// </summary>
    public static class AppSettings
    {
        private static readonly object SyncRoot = new object();
        private static readonly string LegacyConfigFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "AppConfig.ini");

        private static readonly string ConfigFilePath = ResolveConfigFilePath();

        private static string? _cachedDatabasePath;
        private static string? _cachedPhotosFolder;

        /// <summary>
        /// مسیر فایل تنظیمات فعلی.
        /// </summary>
        public static string SettingsFilePath => ConfigFilePath;

        /// <summary>
        /// مسیر پیش‌فرض دیتابیس همراه برنامه.
        /// </summary>
        public static string DefaultDatabasePath
        {
            get
            {
                string[] candidates =
                {
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MyDatabase.accdb"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PersonnelDatabase.accdb"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "MyDatabase.accdb")
                };

                foreach (string candidate in candidates)
                {
                    if (File.Exists(candidate))
                        return candidate;
                }

                return candidates[0];
            }
        }

        /// <summary>
        /// مسیر پیش‌فرض پوشه عکس‌ها.
        /// </summary>
        public static string DefaultPhotosFolder
        {
            get
            {
                string besideExe = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PersonnelPhotos");
                try
                {
                    if (Directory.Exists(besideExe) && Directory.EnumerateFileSystemEntries(besideExe).Any())
                        return besideExe;
                }
                catch
                {
                    // در صورت عدم دسترسی، مسیر کاربر استفاده می‌شود.
                }

                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PersonnelManagementApp",
                    "PersonnelPhotos");
            }
        }

        /// <summary>
        /// مسیر دیتابیس.
        /// </summary>
        public static string DatabasePath
        {
            get
            {
                lock (SyncRoot)
                {
                    if (!string.IsNullOrWhiteSpace(_cachedDatabasePath))
                        return _cachedDatabasePath!;

                    string stored = NormalizePath(ReadSetting("DatabasePath"));
                    if (!string.IsNullOrWhiteSpace(stored) && File.Exists(stored))
                    {
                        _cachedDatabasePath = stored;
                        return stored;
                    }

                    string fallback = DefaultDatabasePath;
                    _cachedDatabasePath = fallback;
                    WriteSetting("DatabasePath", fallback);
                    return fallback;
                }
            }
            set
            {
                lock (SyncRoot)
                {
                    string normalized = NormalizePath(value);
                    _cachedDatabasePath = normalized;
                    WriteSetting("DatabasePath", normalized);
                    // هیچ استخر اتصالی برای OleDb وجود ندارد؛ اتصال بعدی مسیر جدید را می‌خواند.
                }
            }
        }

        /// <summary>
        /// مسیر پوشه عکس‌ها.
        /// </summary>
        public static string PhotosFolder
        {
            get
            {
                lock (SyncRoot)
                {
                    if (!string.IsNullOrWhiteSpace(_cachedPhotosFolder))
                        return _cachedPhotosFolder!;

                    string stored = NormalizePath(ReadSetting("PhotosFolder"));
                    string path = !string.IsNullOrWhiteSpace(stored) ? stored : DefaultPhotosFolder;

                    try
                    {
                        if (!Directory.Exists(path))
                            Directory.CreateDirectory(path);
                    }
                    catch
                    {
                        path = DefaultPhotosFolder;
                        try
                        {
                            if (!Directory.Exists(path))
                                Directory.CreateDirectory(path);
                        }
                        catch
                        {
                            // اگر ساخت پوشه ممکن نبود، مسیر برگردانده می‌شود.
                        }
                    }

                    _cachedPhotosFolder = path;
                    return path;
                }
            }
            set
            {
                lock (SyncRoot)
                {
                    string normalized = NormalizePath(value);
                    _cachedPhotosFolder = normalized;
                    WriteSetting("PhotosFolder", normalized);

                    if (!string.IsNullOrWhiteSpace(normalized) && !Directory.Exists(normalized))
                    {
                        try { Directory.CreateDirectory(normalized); } catch { }
                    }
                }
            }
        }

        /// <summary>
        /// بررسی وجود فایل دیتابیس.
        /// </summary>
        public static bool DatabaseExists()
        {
            string path = DatabasePath;
            return !string.IsNullOrWhiteSpace(path) && File.Exists(path);
        }

        /// <summary>
        /// برگرداندن تنظیمات به حالت پیش‌فرض.
        /// </summary>
        public static void ResetToDefaults()
        {
            lock (SyncRoot)
            {
                _cachedDatabasePath = DefaultDatabasePath;
                _cachedPhotosFolder = DefaultPhotosFolder;

                WriteSetting("DatabasePath", _cachedDatabasePath);
                WriteSetting("PhotosFolder", _cachedPhotosFolder);
            }
        }

        private static string ResolveConfigFilePath()
        {
            string localFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PersonnelManagementApp");

            string localPath = Path.Combine(localFolder, "AppConfig.ini");

            try
            {
                Directory.CreateDirectory(localFolder);

                // بررسی امکان نوشتن در مسیر کاربر
                string probe = Path.Combine(localFolder, ".write-test");
                File.WriteAllText(probe, "ok");
                File.Delete(probe);

                return localPath;
            }
            catch
            {
                return LegacyConfigFilePath;
            }
        }

        private static string NormalizePath(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return string.Empty;

            try
            {
                return Path.GetFullPath(Environment.ExpandEnvironmentVariables(path.Trim().Trim('"')));
            }
            catch
            {
                return path!.Trim();
            }
        }

        private static string ReadSetting(string key)
        {
            lock (SyncRoot)
            {
                string value = ReadSettingFromFile(ConfigFilePath, key);
                if (!string.IsNullOrWhiteSpace(value))
                    return value;

                // مهاجرت از تنظیمات قدیمی کنار فایل اجرایی
                string legacy = ReadSettingFromFile(LegacyConfigFilePath, key);
                if (!string.IsNullOrWhiteSpace(legacy))
                {
                    WriteSettingToFile(ConfigFilePath, key, legacy);
                    return legacy;
                }

                return string.Empty;
            }
        }

        private static string ReadSettingFromFile(string filePath, string key)
        {
            try
            {
                if (!File.Exists(filePath))
                    return string.Empty;

                foreach (string line in File.ReadAllLines(filePath, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#"))
                        continue;

                    int separator = line.IndexOf('=');
                    if (separator <= 0)
                        continue;

                    string lineKey = line.Substring(0, separator).Trim();
                    if (string.Equals(lineKey, key, StringComparison.OrdinalIgnoreCase))
                        return line.Substring(separator + 1).Trim();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در خواندن تنظیمات: {ex.Message}");
                ErrorLogger.Log("AppSettings.ReadSetting", ex);
            }

            return string.Empty;
        }

        private static void WriteSetting(string key, string value)
        {
            lock (SyncRoot)
            {
                WriteSettingToFile(ConfigFilePath, key, value);
            }
        }

        private static void WriteSettingToFile(string filePath, string key, string value)
        {
            try
            {
                var lines = File.Exists(filePath)
                    ? File.ReadAllLines(filePath, Encoding.UTF8).Where(line => !string.IsNullOrWhiteSpace(line)).ToList()
                    : new System.Collections.Generic.List<string>();

                bool found = false;
                for (int i = 0; i < lines.Count; i++)
                {
                    int separator = lines[i].IndexOf('=');
                    if (separator <= 0)
                        continue;

                    if (string.Equals(lines[i].Substring(0, separator).Trim(), key, StringComparison.OrdinalIgnoreCase))
                    {
                        lines[i] = key + "=" + value;
                        found = true;
                        break;
                    }
                }

                if (!found)
                    lines.Add(key + "=" + value);

                File.WriteAllLines(filePath, lines, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در نوشتن تنظیمات: {ex.Message}");
                ErrorLogger.Log("AppSettings.WriteSetting", ex);
            }
        }
    }
}
