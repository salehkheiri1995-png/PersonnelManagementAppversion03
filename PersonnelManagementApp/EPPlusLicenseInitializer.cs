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
                // روش برای EPPlus 8 و جدیدتر
                ExcelPackage.License.SetNonCommercialPersonal("PersonnelManagementApp");
                _initialized = true;
            }
            catch
            {
                // اگر مشکلی پیش آمد ادامه بده
                _initialized = true;
            }
        }
    }
}