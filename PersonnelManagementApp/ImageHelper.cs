using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    /// <summary>
    /// کلاس کمکی برای مدیریت تصاویر پرسنل
    /// </summary>
    public static class ImageHelper
    {
        /// <summary>
        /// مسیر پوشه تصاویر از AppSettings
        /// </summary>
        public static string ImagesFolderPath => AppSettings.PhotosFolder;

        /// <summary>
        /// دریافت نام فایل عکس بر اساس کد ملی
        /// </summary>
        /// <param name="nationalID">کد ملی پرسنل</param>
        /// <returns>نام فایل با پسوند jpg</returns>
        public static string GetImageFileName(string nationalID)
        {
            if (string.IsNullOrWhiteSpace(nationalID))
                return string.Empty;
            
            return $"{nationalID.Trim()}.jpg";
        }

        /// <summary>
        /// دریافت مسیر کامل فایل عکس
        /// </summary>
        /// <param name="nationalID">کد ملی پرسنل</param>
        /// <returns>مسیر کامل فایل</returns>
        public static string GetImageFilePath(string nationalID)
        {
            if (string.IsNullOrWhiteSpace(nationalID))
                return string.Empty;
            
            string fileName = GetImageFileName(nationalID);
            return Path.Combine(ImagesFolderPath, fileName);
        }

        /// <summary>
        /// بررسی وجود فایل عکس برای پرسنل
        /// </summary>
        /// <param name="nationalID">کد ملی پرسنل</param>
        /// <returns>true اگر عکس وجود دارد</returns>
        public static bool ImageExists(string nationalID)
        {
            if (string.IsNullOrWhiteSpace(nationalID))
                return false;
            
            string filePath = GetImageFilePath(nationalID);
            return File.Exists(filePath);
        }

        /// <summary>
        /// ذخیره عکس با نام کد ملی
        /// </summary>
        /// <param name="sourceFilePath">مسیر فایل اصلی</param>
        /// <param name="nationalID">کد ملی پرسنل</param>
        /// <returns>true اگر موفقیت‌آمیز بود</returns>
        public static bool SaveImage(string sourceFilePath, string nationalID)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(sourceFilePath) || !File.Exists(sourceFilePath))
                {
                    MessageBox.Show("فایل عکس معتبر نیست.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                if (string.IsNullOrWhiteSpace(nationalID))
                {
                    MessageBox.Show("کد ملی معتبر نیست.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                string destinationPath = GetImageFilePath(nationalID);

                // اگر عکس قبلی وجود دارد، حذف شود
                if (File.Exists(destinationPath))
                {
                    File.Delete(destinationPath);
                }

                // بارگذاری تصویر و ذخیره با فرمت استاندارد
                using (Image img = Image.FromFile(sourceFilePath))
                {
                    // ذخیره با کیفیت بالا
                    ImageCodecInfo jpgEncoder = GetEncoder(ImageFormat.Jpeg);
                    System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;
                    EncoderParameters myEncoderParameters = new EncoderParameters(1);
                    EncoderParameter myEncoderParameter = new EncoderParameter(myEncoder, 90L);
                    myEncoderParameters.Param[0] = myEncoderParameter;
                    
                    img.Save(destinationPath, jpgEncoder, myEncoderParameters);
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در ذخیره عکس: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// حذف عکس پرسنل
        /// </summary>
        /// <param name="nationalID">کد ملی پرسنل</param>
        /// <returns>true اگر موفقیت‌آمیز بود</returns>
        public static bool DeleteImage(string nationalID)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(nationalID))
                    return false;

                string filePath = GetImageFilePath(nationalID);
                
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در حذف عکس: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// بارگذاری عکس پرسنل
        /// </summary>
        /// <param name="nationalID">کد ملی پرسنل</param>
        /// <returns>تصویر یا null اگر وجود نداشت</returns>
        public static Image LoadImage(string nationalID)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(nationalID))
                    return null;

                string filePath = GetImageFilePath(nationalID);
                
                if (File.Exists(filePath))
                {
                    // بارگذاری تصویر بدون قفل کردن فایل
                    using (var bmpTemp = new Bitmap(filePath))
                    {
                        return new Bitmap(bmpTemp);
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری عکس: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// تغییر نام فایل عکس هنگام تغییر کد ملی
        /// </summary>
        /// <param name="oldNationalID">کد ملی قبلی</param>
        /// <param name="newNationalID">کد ملی جدید</param>
        /// <returns>true اگر موفقیت‌آمیز بود</returns>
        public static bool RenameImage(string oldNationalID, string newNationalID)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(oldNationalID) || string.IsNullOrWhiteSpace(newNationalID))
                    return false;

                if (oldNationalID.Trim() == newNationalID.Trim())
                    return true; // تغییری نیست

                string oldPath = GetImageFilePath(oldNationalID);
                string newPath = GetImageFilePath(newNationalID);

                if (File.Exists(oldPath))
                {
                    // اگر فایل جدید وجود دارد، حذف شود
                    if (File.Exists(newPath))
                    {
                        File.Delete(newPath);
                    }

                    File.Move(oldPath, newPath);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در تغییر نام عکس: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// دریافت Encoder برای فرمت تصویر
        /// </summary>
        private static ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        /// <summary>
        /// باز کردن دیالوگ انتخاب عکس
        /// </summary>
        /// <returns>مسیر فایل انتخاب شده یا خالی</returns>
        public static string OpenImageDialog()
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "انتخاب عکس پرسنل";
                    ofd.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif";
                    ofd.FilterIndex = 1;

                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        return ofd.FileName;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در انتخاب عکس: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return string.Empty;
        }

        /// <summary>
        /// رسم عکس در PictureBox با حفظ نسبت ابعاد
        /// </summary>
        /// <param name="pictureBox">PictureBox مقصد</param>
        /// <param name="image">تصویر</param>
        public static void DrawImageInPictureBox(PictureBox pictureBox, Image image)
        {
            if (pictureBox == null)
                return;

            if (image == null)
            {
                pictureBox.Image = null;
                return;
            }

            pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox.Image = image;
        }

        /// <summary>
        /// ایجاد تصویر پیش‌فرض برای PictureBox
        /// </summary>
        /// <param name="width">عرض</param>
        /// <param name="height">ارتفاع</param>
        /// <returns>تصویر پیش‌فرض</returns>
        public static Image CreateDefaultImage(int width, int height)
        {
            Bitmap bmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightGray);
                
                // رسم متن "بدون عکس"
                string text = "بدون عکس";
                Font font = new Font("Tahoma", 12, FontStyle.Bold);
                SizeF textSize = g.MeasureString(text, font);
                PointF textPoint = new PointF((width - textSize.Width) / 2, (height - textSize.Height) / 2);
                
                g.DrawString(text, font, Brushes.DarkGray, textPoint);
            }
            return bmp;
        }
    }
}