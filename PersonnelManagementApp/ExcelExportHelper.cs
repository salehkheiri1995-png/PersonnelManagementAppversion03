using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace PersonnelManagementApp
{
    public static class ExcelExportHelper
    {
        // نقشه فارسی فیلدها
        private static readonly Dictionary<string, string> PersianHeaders = new Dictionary<string, string>
        {
            { "PersonnelID", "شناسه" },
            { "FirstName", "نام" },
            { "LastName", "نام‌خانوادگی" },
            { "PersonnelNumber", "شماره پرسنلی" },
            { "NationalID", "کد ملی" },
            { "PostName", "پست" },
            { "DeptName", "اداره" },
            { "Province", "استان" },
            { "City", "شهر" },
            { "Affair", "امور" },
            { "District", "ناحیه" },
            { "ContractType", "نوع قرارداد" },
            { "HireDate", "تاریخ استخدام" },
            { "MobileNumber", "تلفن همراه" },
            { "Gender", "جنسیت" },
            { "Education", "تحصیلات" },
            { "JobLevel", "سطح شغلی" },
            { "Company", "شرکت" },
            { "WorkShift", "شیفت کاری" },
            { "Salary", "حقوق" },
            { "Email", "ایمیل" },
            { "BirthDate", "تاریخ تولد" },
            { "Address", "آدرس" }
        };

        // تنظیم License برای EPPlus 8+
        static ExcelExportHelper()
        {
            try
            {
                // روش صحیح برای EPPlus 8+
                ExcelPackage.License.SetLicenseInformation("NonCommercial");
            }
            catch
            {
                // اگر خطا داد، سعی می‌کنیم بدون License هم کار کنه
            }
        }

        /// <summary>
        /// Export DataGridView to Excel with selected columns
        /// </summary>
        public static void ExportToExcel(DataGridView dgv, List<string> selectedColumns, string defaultFileName = "PersonnelData")
        {
            try
            {
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "Excel Files|*.xlsx";
                    saveDialog.Title = "ذخیره فایل اکسل";
                    saveDialog.FileName = $"{defaultFileName}_{DateTime.Now:yyyy-MM-dd_HHmmss}.xlsx";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportToExcelFile(dgv, selectedColumns, saveDialog.FileName);

                        DialogResult result = MessageBox.Show(
                            "✅ فایل اکسل با موفقیت ذخیره شد!\n\nآیا می‌خواهید فایل را باز کنید؟",
                            "موفق",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information);

                        if (result == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = saveDialog.FileName,
                                UseShellExecute = true
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در ذخیره فایل اکسل:\n{ex.Message}",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Export List of PersonnelDetail to Excel
        /// </summary>
        public static void ExportToExcel(List<PersonnelDetail> personnelList, List<string> selectedColumns, string defaultFileName = "PersonnelData")
        {
            try
            {
                using (SaveFileDialog saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "Excel Files|*.xlsx";
                    saveDialog.Title = "ذخیره فایل اکسل";
                    saveDialog.FileName = $"{defaultFileName}_{DateTime.Now:yyyy-MM-dd_HHmmss}.xlsx";

                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportToExcelFile(personnelList, selectedColumns, saveDialog.FileName);

                        DialogResult result = MessageBox.Show(
                            "✅ فایل اکسل با موفقیت ذخیره شد!\n\nآیا می‌خواهید فایل را باز کنید؟",
                            "موفق",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information);

                        if (result == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = saveDialog.FileName,
                                UseShellExecute = true
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در ذخیره فایل اکسل:\n{ex.Message}",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ExportToExcelFile(DataGridView dgv, List<string> selectedColumns, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Personnel");

                // نوشتن هدرها
                int col = 1;
                foreach (var columnName in selectedColumns)
                {
                    string header = PersianHeaders.ContainsKey(columnName) ? PersianHeaders[columnName] : columnName;
                    worksheet.Cells[1, col].Value = header;

                    // استایل هدر
                    worksheet.Cells[1, col].Style.Font.Bold = true;
                    worksheet.Cells[1, col].Style.Font.Size = 12;
                    worksheet.Cells[1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 102, 204));
                    worksheet.Cells[1, col].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[1, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[1, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    col++;
                }

                // نوشتن داده‌ها
                int row = 2;
                foreach (DataGridViewRow dgvRow in dgv.Rows)
                {
                    if (dgvRow.IsNewRow) continue;

                    col = 1;
                    foreach (var columnName in selectedColumns)
                    {
                        if (dgv.Columns.Contains(columnName))
                        {
                            var cellValue = dgvRow.Cells[columnName]?.Value;
                            if (cellValue != null)
                            {
                                if (cellValue is DateTime dateValue)
                                {
                                    worksheet.Cells[row, col].Value = dateValue.ToString("yyyy/MM/dd");
                                }
                                else if (columnName == "Salary" && decimal.TryParse(cellValue.ToString(), out decimal salary))
                                {
                                    worksheet.Cells[row, col].Value = salary;
                                    worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0";
                                }
                                else
                                {
                                    worksheet.Cells[row, col].Value = cellValue.ToString();
                                }
                            }
                        }
                        col++;
                    }
                    row++;
                }

                // تنظیمات ظاهری
                worksheet.Cells.AutoFitColumns();
                worksheet.View.RightToLeft = true;

                // افزودن border
                if (selectedColumns.Count > 0 && row > 1)
                {
                    var dataRange = worksheet.Cells[1, 1, row - 1, selectedColumns.Count];
                    dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                // ذخیره فایل
                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }
        }

        private static void ExportToExcelFile(List<PersonnelDetail> personnelList, List<string> selectedColumns, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Personnel");

                // نوشتن هدرها
                int col = 1;
                foreach (var columnName in selectedColumns)
                {
                    string header = PersianHeaders.ContainsKey(columnName) ? PersianHeaders[columnName] : columnName;
                    worksheet.Cells[1, col].Value = header;

                    // استایل هدر
                    worksheet.Cells[1, col].Style.Font.Bold = true;
                    worksheet.Cells[1, col].Style.Font.Size = 12;
                    worksheet.Cells[1, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[1, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 102, 204));
                    worksheet.Cells[1, col].Style.Font.Color.SetColor(Color.White);
                    worksheet.Cells[1, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[1, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    col++;
                }

                // نوشتن داده‌ها
                int row = 2;
                foreach (var person in personnelList)
                {
                    col = 1;
                    foreach (var columnName in selectedColumns)
                    {
                        object? value = GetPropertyValue(person, columnName);
                        if (value != null)
                        {
                            if (value is DateTime dateValue)
                            {
                                worksheet.Cells[row, col].Value = dateValue.ToString("yyyy/MM/dd");
                            }
                            else if (columnName == "Salary" && decimal.TryParse(value.ToString(), out decimal salary))
                            {
                                worksheet.Cells[row, col].Value = salary;
                                worksheet.Cells[row, col].Style.Numberformat.Format = "#,##0";
                            }
                            else
                            {
                                worksheet.Cells[row, col].Value = value.ToString();
                            }
                        }
                        col++;
                    }
                    row++;
                }

                // تنظیمات ظاهری
                worksheet.Cells.AutoFitColumns();
                worksheet.View.RightToLeft = true;

                // افزودن border
                if (selectedColumns.Count > 0 && row > 1)
                {
                    var dataRange = worksheet.Cells[1, 1, row - 1, selectedColumns.Count];
                    dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                // ذخیره فایل
                FileInfo excelFile = new FileInfo(filePath);
                package.SaveAs(excelFile);
            }
        }

        private static object? GetPropertyValue(PersonnelDetail person, string propertyName)
        {
            var property = typeof(PersonnelDetail).GetProperty(propertyName);
            return property?.GetValue(person);
        }
    }
}