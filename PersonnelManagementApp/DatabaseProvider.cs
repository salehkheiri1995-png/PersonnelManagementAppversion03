using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PersonnelManagementApp
{
    /// <summary>
    /// شناسایی و راه‌اندازی موتور پایگاه‌داده Microsoft Access (ACE OLE DB).
    /// </summary>
    public static class DatabaseProvider
    {
        public const string Ace12 = "Microsoft.ACE.OLEDB.12.0";
        public const string Ace16 = "Microsoft.ACE.OLEDB.16.0";

        private const string WingetPackageId = "Microsoft.AccessDatabaseEngine.2016";
        private const string DownloadPage = "https://www.microsoft.com/en-us/download/details.aspx?id=54920";
        private const string AceX64Download = "https://download.microsoft.com/download/3/5/c/35c84c36-661a-44e6-9324-8786b8dbe231/accessdatabaseengine_X64.exe";

        private static readonly object SyncRoot = new object();
        private static bool _detected;
        private static string? _providerName;

        /// <summary>
        /// نام provider در دسترس در فرایند جاری (x86 یا x64).
        /// </summary>
        public static string? ProviderName
        {
            get
            {
                lock (SyncRoot)
                {
                    if (!_detected)
                    {
                        _providerName = DetectProvider();
                        _detected = true;
                    }

                    return _providerName;
                }
            }
        }

        public static bool IsAvailable => !string.IsNullOrEmpty(ProviderName);

        public static string Bitness => Environment.Is64BitProcess ? "64-bit" : "32-bit";

        public static void Invalidate()
        {
            lock (SyncRoot)
            {
                _detected = false;
                _providerName = null;
            }
        }

        private static string? DetectProvider()
        {
            try
            {
                var enumerator = new OleDbEnumerator();
                DataTable providers = enumerator.GetElements();

                var names = providers.Rows
                    .Cast<DataRow>()
                    .Select(row => row["SOURCES_NAME"]?.ToString())
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Select(name => name!)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                if (names.Contains(Ace16))
                    return Ace16;

                if (names.Contains(Ace12))
                    return Ace12;
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("DatabaseProvider.DetectProvider", ex);
            }

            return null;
        }

        public static string BuildConnectionString(string databasePath)
        {
            if (string.IsNullOrWhiteSpace(databasePath))
                throw new ArgumentException("مسیر پایگاه داده تعیین نشده است.", nameof(databasePath));

            string? provider = ProviderName;
            if (string.IsNullOrEmpty(provider))
                throw new InvalidOperationException("موتور پایگاه‌داده Microsoft Access (ACE OLE DB) روی این سیستم نصب نیست.");

            return $"Provider={provider};Data Source={databasePath};Persist Security Info=False;";
        }

        /// <summary>
        /// بررسی می‌کند آیا provider در فرایند جاری قابل استفاده است.
        /// </summary>
        public static bool CheckProvider(IWin32Window? owner, bool offerInstall)
        {
            if (IsAvailable)
                return true;

            if (!offerInstall)
                return false;

            return ShowMissingProviderDialog(owner);
        }

        /// <summary>
        /// دیالوگ راهنمای نصب provider. در صورت تایید کاربر، نصب خودکار انجام می‌شود.
        /// </summary>
        public static bool ShowMissingProviderDialog(IWin32Window? owner = null)
        {
            Invalidate();
            if (IsAvailable)
                return true;

            bool registered32 = IsRegistered(RegistryView.Registry32, Ace12) || IsRegistered(RegistryView.Registry32, Ace16);
            bool registered64 = IsRegistered(RegistryView.Registry64, Ace12) || IsRegistered(RegistryView.Registry64, Ace16);

            string bitnessHint;
            if (Environment.Is64BitProcess && registered32 && !registered64)
            {
                bitnessHint = "نسخه ۳۲ بیتی موتور Access روی سیستم نصب است، اما برنامه در حالت ۶۴ بیتی اجرا می‌شود.\n" +
                              "لطفاً نسخه ۶۴ بیتی (x64) را نصب کنید.";
            }
            else if (!Environment.Is64BitProcess && registered64 && !registered32)
            {
                bitnessHint = "نسخه ۶۴ بیتی موتور Access روی سیستم نصب است، اما برنامه در حالت ۳۲ بیتی اجرا می‌شود.\n" +
                              "لطفاً نسخه ۳۲ بیتی (x86) را نصب کنید.";
            }
            else
            {
                bitnessHint = "برای باز کردن فایل‌های .accdb باید «Microsoft Access Database Engine» نصب باشد.\n" +
                              $"برنامه در حالت {Bitness} اجرا می‌شود.";
            }

            string message =
                "موتور پایگاه‌داده Access (ACE OLE DB) پیدا نشد.\n\n" +
                bitnessHint + "\n\n" +
                "می‌خواهید نصب خودکار انجام شود؟\n" +
                "در صورت انتخاب «بله»، برنامه تلاش می‌کند پیش‌نیاز را نصب کند.\n" +
                "در صورت انتخاب «خیر»، صفحه دانلود رسمی Microsoft باز می‌شود.";

            DialogResult result = owner == null
                ? MessageBox.Show(message, "پیش‌نیاز پایگاه‌داده", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)
                : MessageBox.Show(owner, message, "پیش‌نیاز پایگاه‌داده", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);

            if (result == DialogResult.Cancel)
                return false;

            if (result == DialogResult.No)
            {
                OpenDownloadPage();
                return false;
            }

            return InstallProvider(owner);
        }

        public static bool InstallProvider(IWin32Window? owner = null)
        {
            Cursor? previousCursor = owner is Form form ? form.Cursor : null;
            try
            {
                if (owner is Form ownerForm)
                {
                    ownerForm.UseWaitCursor = true;
                    ownerForm.Cursor = Cursors.WaitCursor;
                    Application.DoEvents();
                }

                if (!Environment.Is64BitProcess && Environment.Is64BitOperatingSystem)
                {
                    ShowInfo(owner,
                        "برنامه در حالت ۳۲ بیتی اجرا می‌شود.\n" +
                        "لطفاً نسخه x86 موتور Access Database Engine را از صفحه رسمی Microsoft نصب کنید.");
                    OpenDownloadPage();
                    return false;
                }

                // ۱) تلاش با winget
                if (TryInstallWithWinget())
                {
                    Invalidate();
                    if (WaitForProvider())
                    {
                        ShowInfo(owner, "✅ موتور پایگاه‌داده Access با موفقیت نصب شد.");
                        return true;
                    }
                }

                // ۲) دانلود و اجرای نصب‌کننده رسمی
                string installerPath = Path.Combine(Path.GetTempPath(), "accessdatabaseengine_X64.exe");
                if (!DownloadInstaller(installerPath, owner))
                    return false;

                if (!RunInstaller(installerPath, owner))
                    return false;

                Invalidate();
                if (WaitForProvider())
                {
                    ShowInfo(owner, "✅ موتور پایگاه‌داده Access با موفقیت نصب شد.");
                    return true;
                }

                ShowInfo(owner,
                    "نصب‌کننده اجرا شد، اما هنوز provider در دسترس نیست.\n\n" +
                    "لطفاً سیستم را یک‌بار ری‌استارت کنید و دوباره برنامه را اجرا کنید.\n" +
                    "در صورت ادامه مشکل، نسخه x64 را از صفحه رسمی Microsoft نصب کنید.");
                OpenDownloadPage();
                return false;
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("DatabaseProvider.InstallProvider", ex);
                ShowError(owner, $"نصب خودکار انجام نشد:\n{ex.Message}\n\nمی‌توانید به‌صورت دستی از صفحه رسمی Microsoft نصب کنید.");
                OpenDownloadPage();
                return false;
            }
            finally
            {
                if (owner is Form ownerForm)
                {
                    ownerForm.UseWaitCursor = false;
                    ownerForm.Cursor = previousCursor ?? Cursors.Default;
                }
            }
        }

        private static bool TryInstallWithWinget()
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "winget",
                    Arguments = $"install --id {WingetPackageId} --exact --silent --accept-package-agreements --accept-source-agreements --disable-interactivity",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                using var process = Process.Start(psi);
                if (process == null)
                    return false;

                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();

                if (!process.WaitForExit(300_000))
                {
                    try { process.Kill(true); } catch { }
                    return false;
                }

                if (process.ExitCode != 0)
                {
                    ErrorLogger.LogMessage("DatabaseProvider.TryInstallWithWinget",
                        $"winget exit code: {process.ExitCode}\n{output}\n{error}");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("DatabaseProvider.TryInstallWithWinget", ex);
                return false;
            }
        }

        private static bool DownloadInstaller(string destination, IWin32Window? owner)
        {
            try
            {
                if (File.Exists(destination) && new FileInfo(destination).Length > 10 * 1024 * 1024)
                    return true;

                using var client = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };
                using var response = client.GetAsync(AceX64Download).GetAwaiter().GetResult();
                response.EnsureSuccessStatusCode();

                byte[] bytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                if (bytes.Length < 10 * 1024 * 1024)
                    throw new InvalidOperationException("فایل دانلودشده معتبر نیست.");

                File.WriteAllBytes(destination, bytes);
                return true;
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("DatabaseProvider.DownloadInstaller", ex);
                ShowError(owner, $"دانلود نصب‌کننده ممکن نشد:\n{ex.Message}");
                return false;
            }
        }

        private static bool RunInstaller(string installerPath, IWin32Window? owner)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = installerPath,
                    Arguments = "/quiet",
                    UseShellExecute = true,
                    Verb = "runas"
                };

                using var process = Process.Start(psi);
                if (process == null)
                    return false;

                if (!process.WaitForExit(600_000))
                {
                    try { process.Kill(true); } catch { }
                    ShowError(owner, "نصب‌کننده در زمان مجاز پایان پیدا نکرد.");
                    return false;
                }

                // کدهای 0 و 3010 (نیاز به راه‌اندازی مجدد) موفق محسوب می‌شوند.
                return process.ExitCode == 0 || process.ExitCode == 3010;
            }
            catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                ShowInfo(owner, "نصب توسط کاربر لغو شد.");
                return false;
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("DatabaseProvider.RunInstaller", ex);
                ShowError(owner, $"اجرای نصب‌کننده ممکن نشد:\n{ex.Message}");
                return false;
            }
        }

        private static bool WaitForProvider()
        {
            for (int i = 0; i < 20; i++)
            {
                Invalidate();
                if (IsAvailable)
                    return true;
                Thread.Sleep(500);
            }

            return false;
        }

        private static bool IsRegistered(RegistryView view, string provider)
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
                using var key = baseKey.OpenSubKey($@"SOFTWARE\Classes\{provider}\CLSID");
                if (key != null)
                    return true;

                using var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                using var userKey = currentUser.OpenSubKey($@"SOFTWARE\Classes\{provider}\CLSID");
                return userKey != null;
            }
            catch
            {
                return false;
            }
        }

        public static void OpenDownloadPage()
        {
            try
            {
                Process.Start(new ProcessStartInfo(DownloadPage) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("DatabaseProvider.OpenDownloadPage", ex);
            }
        }

        private static void ShowInfo(IWin32Window? owner, string message)
        {
            if (owner == null)
                MessageBox.Show(message, "پیش‌نیاز پایگاه‌داده", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(owner, message, "پیش‌نیاز پایگاه‌داده", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void ShowError(IWin32Window? owner, string message)
        {
            if (owner == null)
                MessageBox.Show(message, "خطا در پیش‌نیاز پایگاه‌داده", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                MessageBox.Show(owner, message, "خطا در پیش‌نیاز پایگاه‌داده", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
