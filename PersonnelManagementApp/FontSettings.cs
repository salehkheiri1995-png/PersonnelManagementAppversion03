using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    /// <summary>
    /// کلاس مدیریت تنظیمات فونت برنامه
    /// </summary>
    public static class FontSettings
    {
        private static readonly string ConfigFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, 
            "FontSettings.ini"
        );

        // تنظیمات پیش‌فرض
        private const string DefaultFontFamilyName = "Tahoma";
        private const float DefaultTitleSize = 16f;
        private const float DefaultLabelSize = 12f;
        private const float DefaultTextBoxSize = 11f;
        private const float DefaultButtonSize = 12f;
        private const float DefaultBodySize = 10f;
        private const float DefaultChartLabelSize = 10f;

        // Propertyهای عمومی برای دسترسی
        public static string FontFamilyName { get; set; }
        public static float TitleFontSize { get; set; }
        public static float LabelFontSize { get; set; }
        public static float TextBoxFontSize { get; set; }
        public static float ButtonFontSize { get; set; }
        public static float BodyFontSize { get; set; }
        public static float ChartLabelFontSize { get; set; }
        public static bool TitleFontBold { get; set; }
        public static bool LabelFontBold { get; set; }
        public static bool ButtonFontBold { get; set; }
        public static bool ChartLabelFontBold { get; set; }

        // فونت‌های آماده
        public static Font TitleFont { get; private set; }
        public static Font HeaderFont { get; private set; }
        public static Font SubtitleFont { get; private set; }
        public static Font BodyFont { get; private set; }
        public static Font ButtonFont { get; private set; }
        public static Font LabelFont { get; private set; }
        public static Font TextBoxFont { get; private set; }
        public static Font ChartLabelFont { get; private set; }

        // رویداد تغییر فونت
        public static event EventHandler FontChanged;

        static FontSettings()
        {
            LoadSettings();
            UpdateFonts();
        }

        /// <summary>
        /// بارگذاری تنظیمات از فایل INI
        /// </summary>
        private static void LoadSettings()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var lines = File.ReadAllLines(ConfigFilePath);
                    foreach (var line in lines)
                    {
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                            continue;

                        var parts = line.Split('=');
                        if (parts.Length != 2)
                            continue;

                        var key = parts[0].Trim();
                        var value = parts[1].Trim();

                        switch (key)
                        {
                            case "FontFamilyName":
                                FontFamilyName = value;
                                break;
                            case "TitleFontSize":
                                if (float.TryParse(value, out float titleSize))
                                    TitleFontSize = titleSize;
                                break;
                            case "LabelFontSize":
                                if (float.TryParse(value, out float labelSize))
                                    LabelFontSize = labelSize;
                                break;
                            case "TextBoxFontSize":
                                if (float.TryParse(value, out float textBoxSize))
                                    TextBoxFontSize = textBoxSize;
                                break;
                            case "ButtonFontSize":
                                if (float.TryParse(value, out float buttonSize))
                                    ButtonFontSize = buttonSize;
                                break;
                            case "BodyFontSize":
                                if (float.TryParse(value, out float bodySize))
                                    BodyFontSize = bodySize;
                                break;
                            case "ChartLabelFontSize":
                                if (float.TryParse(value, out float chartLabelSize))
                                    ChartLabelFontSize = chartLabelSize;
                                break;
                            case "TitleFontBold":
                                TitleFontBold = value.ToLower() == "true";
                                break;
                            case "LabelFontBold":
                                LabelFontBold = value.ToLower() == "true";
                                break;
                            case "ButtonFontBold":
                                ButtonFontBold = value.ToLower() == "true";
                                break;
                            case "ChartLabelFontBold":
                                ChartLabelFontBold = value.ToLower() == "true";
                                break;
                        }
                    }
                }
                else
                {
                    // اگر فایل وجود ندارد، از مقادیر پیش‌فرض استفاده کن
                    ResetToDefaults();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در بارگذاری تنظیمات فونت: {ex.Message}");
                ResetToDefaults();
            }
        }

        /// <summary>
        /// ذخیره تنظیمات در فایل INI
        /// </summary>
        public static void SaveSettings()
        {
            try
            {
                var content = $"# Font Settings\n" +
                    $"FontFamilyName={FontFamilyName}\n" +
                    $"TitleFontSize={TitleFontSize}\n" +
                    $"TitleFontBold={TitleFontBold.ToString().ToLower()}\n" +
                    $"LabelFontSize={LabelFontSize}\n" +
                    $"LabelFontBold={LabelFontBold.ToString().ToLower()}\n" +
                    $"TextBoxFontSize={TextBoxFontSize}\n" +
                    $"ButtonFontSize={ButtonFontSize}\n" +
                    $"ButtonFontBold={ButtonFontBold.ToString().ToLower()}\n" +
                    $"BodyFontSize={BodyFontSize}\n" +
                    $"ChartLabelFontSize={ChartLabelFontSize}\n" +
                    $"ChartLabelFontBold={ChartLabelFontBold.ToString().ToLower()}\n";

                File.WriteAllText(ConfigFilePath, content);
                UpdateFonts();

                // فعال‌سازی رویداد تغییر فونت
                FontChanged?.Invoke(null, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                throw new Exception($"خطا در ذخیره تنظیمات: {ex.Message}");
            }
        }

        /// <summary>
        /// بازگشت به تنظیمات پیش‌فرض
        /// </summary>
        public static void ResetToDefaults()
        {
            FontFamilyName = DefaultFontFamilyName;
            TitleFontSize = DefaultTitleSize;
            LabelFontSize = DefaultLabelSize;
            TextBoxFontSize = DefaultTextBoxSize;
            ButtonFontSize = DefaultButtonSize;
            BodyFontSize = DefaultBodySize;
            ChartLabelFontSize = DefaultChartLabelSize;
            TitleFontBold = true;
            LabelFontBold = false;
            ButtonFontBold = true;
            ChartLabelFontBold = false;

            UpdateFonts();
        }

        /// <summary>
        /// بروزرسانی تمام فونت‌ها بر اساس تنظیمات جدید
        /// </summary>
        private static void UpdateFonts()
        {
            try
            {
                // فونت عنوان بزرگ (Title)
                TitleFont = new Font(
                    FontFamilyName, 
                    TitleFontSize, 
                    TitleFontBold ? FontStyle.Bold : FontStyle.Regular
                );

                // فونت سرتیتر (Header)
                HeaderFont = new Font(
                    FontFamilyName, 
                    TitleFontSize - 2, 
                    FontStyle.Bold
                );

                // فونت زیرعنوان (Subtitle)
                SubtitleFont = new Font(
                    FontFamilyName, 
                    LabelFontSize + 2, 
                    FontStyle.Bold
                );

                // فونت متن عادی (Body)
                BodyFont = new Font(
                    FontFamilyName, 
                    BodyFontSize, 
                    FontStyle.Regular
                );

                // فونت دکمه‌ها (Button)
                ButtonFont = new Font(
                    FontFamilyName, 
                    ButtonFontSize, 
                    ButtonFontBold ? FontStyle.Bold : FontStyle.Regular
                );

                // فونت برچسب‌ها (Label)
                LabelFont = new Font(
                    FontFamilyName, 
                    LabelFontSize, 
                    LabelFontBold ? FontStyle.Bold : FontStyle.Regular
                );

                // فونت جعبه متن (TextBox)
                TextBoxFont = new Font(
                    FontFamilyName, 
                    TextBoxFontSize, 
                    FontStyle.Regular
                );

                // فونت متن داخل نمودارها (Chart Labels)
                ChartLabelFont = new Font(
                    FontFamilyName, 
                    ChartLabelFontSize, 
                    ChartLabelFontBold ? FontStyle.Bold : FontStyle.Regular
                );
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در بروزرسانی فونت‌ها: {ex.Message}");
            }
        }

        /// <summary>
        /// دریافت فونت سفارشی با اندازه دلخواه
        /// </summary>
        public static Font GetCustomFont(float sizeOffset = 0, FontStyle style = FontStyle.Regular)
        {
            try
            {
                return new Font(FontFamilyName, BodyFontSize + sizeOffset, style);
            }
            catch
            {
                return new Font("Tahoma", 10f, style);
            }
        }

        /// <summary>
        /// دریافت لیست فونت‌های فارسی رایج
        /// </summary>
        public static string[] GetPersianFonts()
        {
            return new string[]
            {
                "Tahoma",
                "B Nazanin",
                "B Titr",
                "B Lotus",
                "B Zar",
                "IRANSans",
                "Yekan",
                "Mitra",
                "Arial",
                "Vazir",
                "Samim",
                "Sahel"
            };
        }

        /// <summary>
        /// اعمال فونت به تمام کنترل‌های یک فرم
        /// </summary>
        public static void ApplyFontToForm(Form form)
        {
            if (form == null) return;

            try
            {
                // اعمال به تمام کنترل‌ها به صورت بازگشتی
                ApplyFontToControls(form.Controls);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در اعمال فونت: {ex.Message}");
            }
        }

        /// <summary>
        /// اعمال فونت به مجموعه کنترل‌ها به صورت بازگشتی
        /// </summary>
        private static void ApplyFontToControls(Control.ControlCollection controls)
        {
            foreach (Control control in controls)
            {
                try
                {
                    // تشخیص نوع کنترل و اعمال فونت مناسب
                    if (control is Button)
                    {
                        control.Font = ButtonFont;
                    }
                    else if (control is Label)
                    {
                        // بررسی برای عنوان‌های بزرگ
                        if (control.Font.Size >= TitleFontSize)
                        {
                            control.Font = TitleFont;
                        }
                        else
                        {
                            control.Font = LabelFont;
                        }
                    }
                    else if (control is TextBox || control is ComboBox || control is NumericUpDown)
                    {
                        control.Font = TextBoxFont;
                    }
                    else if (control is DataGridView)
                    {
                        DataGridView dgv = control as DataGridView;
                        dgv.DefaultCellStyle.Font = BodyFont;
                        dgv.ColumnHeadersDefaultCellStyle.Font = new Font(
                            FontFamilyName, 
                            LabelFontSize, 
                            FontStyle.Bold
                        );
                    }
                    else if (control is GroupBox)
                    {
                        control.Font = HeaderFont;
                    }
                    else
                    {
                        control.Font = BodyFont;
                    }

                    // اعمال به کنترل‌های درونی (بازگشتی)
                    if (control.HasChildren)
                    {
                        ApplyFontToControls(control.Controls);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"خطا در اعمال فونت به {control.Name}: {ex.Message}");
                }
            }
        }
    }
}