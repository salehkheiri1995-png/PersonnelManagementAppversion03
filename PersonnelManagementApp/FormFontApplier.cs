using System;
using System.Drawing;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    /// <summary>
    /// کلاس کمکی برای اعمال فونت به فرم‌ها به صورت خودکار
    /// </summary>
    public static class FormFontApplier
    {
        /// <summary>
        /// اعمال فونت به فرم در زمان Load
        /// </summary>
        public static void ApplyOnLoad(Form form)
        {
            if (form == null) return;

            // ثبت رویداد Load فرم
            form.Load += (sender, e) =>
            {
                try
                {
                    FontSettings.ApplyFontToForm(form);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"خطا در اعمال فونت در {form.Name}: {ex.Message}");
                }
            };
        }

        /// <summary>
        /// اعمال فونت به فرم به صورت مستقیم
        /// </summary>
        public static void ApplyNow(Form form)
        {
            if (form == null) return;

            try
            {
                FontSettings.ApplyFontToForm(form);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در اعمال فونت در {form.Name}: {ex.Message}");
            }
        }

        /// <summary>
        /// اعمال فونت به کنترل خاص
        /// </summary>
        public static void ApplyToControl(Control control, FontStyle style = FontStyle.Regular, float sizeOffset = 0)
        {
            if (control == null) return;

            try
            {
                control.Font = FontSettings.GetCustomFont(sizeOffset, style);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در اعمال فونت به کنترل {control.Name}: {ex.Message}");
            }
        }

        /// <summary>
        /// دریافت فونت بر اساس نوع کنترل
        /// </summary>
        public static Font GetFontForControl(Control control)
        {
            if (control is Button)
                return FontSettings.ButtonFont;
            else if (control is Label)
            {
                // بررسی برای عنوان‌های بزرگ
                if (control.Font != null && control.Font.Size >= FontSettings.TitleFontSize)
                    return FontSettings.TitleFont;
                else
                    return FontSettings.LabelFont;
            }
            else if (control is TextBox || control is ComboBox || control is NumericUpDown)
                return FontSettings.TextBoxFont;
            else if (control is GroupBox)
                return FontSettings.HeaderFont;
            else
                return FontSettings.BodyFont;
        }
    }
}