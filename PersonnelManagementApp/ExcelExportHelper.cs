using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;

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

        /// <summary>
        /// Export DataGridView to Excel with selected columns using ClosedXML
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
        /// Export List of PersonnelDetail to Excel using ClosedXML
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
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Personnel");
                worksheet.RightToLeft = true;

                // نوشتن هدرها
                int col = 1;
                foreach (var columnName in selectedColumns)
                {
                    string header = PersianHeaders.ContainsKey(columnName) ? PersianHeaders[columnName] : columnName;
                    worksheet.Cell(1, col).Value = header;
                    col++;
                }

                // استایل هدر
                var headerRange = worksheet.Range(1, 1, 1, selectedColumns.Count);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Font.FontSize = 12;
                headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 102, 204);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

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
                                    worksheet.Cell(row, col).Value = dateValue.ToString("yyyy/MM/dd");
                                }
                                else if (columnName == "Salary" && decimal.TryParse(cellValue.ToString(), out decimal salary))
                                {
                                    worksheet.Cell(row, col).Value = salary;
                                    worksheet.Cell(row, col).Style.NumberFormat.Format = "#,##0";
                                }
                                else
                                {
                                    worksheet.Cell(row, col).Value = cellValue.ToString();
                                }
                            }
                        }
                        col++;
                    }
                    row++;
                }

                // تنظیمات ظاهری
                worksheet.Columns().AdjustToContents();
                
                // افزودن border
                if (selectedColumns.Count > 0 && row > 1)
                {
                    var dataRange = worksheet.Range(1, 1, row - 1, selectedColumns.Count);
                    dataRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                }

                workbook.SaveAs(filePath);
            }
        }

        private static void ExportToExcelFile(List<PersonnelDetail> personnelList, List<string> selectedColumns, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Personnel");
                worksheet.RightToLeft = true;

                // نوشتن هدرها
                int col = 1;
                foreach (var columnName in selectedColumns)
                {
                    string header = PersianHeaders.ContainsKey(columnName) ? PersianHeaders[columnName] : columnName;
                    worksheet.Cell(1, col).Value = header;
                    col++;
                }

                // استایل هدر
                var headerRange = worksheet.Range(1, 1, 1, selectedColumns.Count);
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Font.FontSize = 12;
                headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 102, 204);
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

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
                                worksheet.Cell(row, col).Value = dateValue.ToString("yyyy/MM/dd");
                            }
                            else if (columnName == "Salary" && decimal.TryParse(value.ToString(), out decimal salary))
                            {
                                worksheet.Cell(row, col).Value = salary;
                                worksheet.Cell(row, col).Style.NumberFormat.Format = "#,##0";
                            }
                            else
                            {
                                worksheet.Cell(row, col).Value = value.ToString();
                            }
                        }
                        col++;
                    }
                    row++;
                }

                // تنظیمات ظاهری
                worksheet.Columns().AdjustToContents();

                // افزودن border
                if (selectedColumns.Count > 0 && row > 1)
                {
                    var dataRange = worksheet.Range(1, 1, row - 1, selectedColumns.Count);
                    dataRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    dataRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                }

                workbook.SaveAs(filePath);
            }
        }

        private static object? GetPropertyValue(PersonnelDetail person, string propertyName)
        {
            var property = typeof(PersonnelDetail).GetProperty(propertyName);
            return property?.GetValue(person);
        }
    }
}