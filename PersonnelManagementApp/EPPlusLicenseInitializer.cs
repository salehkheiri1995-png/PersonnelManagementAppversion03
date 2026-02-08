using System;
using OfficeOpenXml;

namespace PersonnelManagementApp
{
    /// <summary>
    /// مقداردهی اولیه License برای EPPlus
    /// باید در ابتدای برنامه صدا زده شود
    /// </summary>
    public static class EPPlusLicenseInitializer
    {
        private static bool _initialized = false;

        public static void Initialize()
        {
            if (_initialized)
                return;

            try
            {
#pragma warning disable CS0618 // Type or member is obsolete
                // روش برای EPPlus 4.5 تا 7.x
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
#pragma warning restore CS0618 // Type or member is obsolete
                _initialized = true;
            }
            catch
            {
                // اگر خطا داد، نسخه جدیدتر است - مشکلی نیست
                _initialized = true;
            }
        }
    }
}