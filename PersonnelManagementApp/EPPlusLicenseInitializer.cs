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
                // روش 1: برای EPPlus 4.5 تا 7.x
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                _initialized = true;
                return;
            }
            catch { }

            // اگر هیچ روشی کار نکرد، فرض می‌کنیم EPPlus نسخه جدیدتر است
            _initialized = true;
        }
    }
}