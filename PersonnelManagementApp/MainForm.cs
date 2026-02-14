using System;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
//using PostDatabaseManager;

namespace PersonnelManagementApp
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            // اعمال فونت‌های تنظیم‌شده
            FontSettings.ApplyFontToForm(this);
            // بارگذاری آیکون برنامه
            LoadApplicationIcon();
        }

        /// <summary>
        /// بارگذاری ایکون برنامه با سایز مناسب
        /// </summary>
        private void LoadApplicationIcon()
        {
            try
            {
                string iconPath = Path.Combine(Application.StartupPath, "app_icon.ico");
                
                if (File.Exists(iconPath))
                {
                    // بارگذاری آیکون با سایز بزرگ (256x256 پیکسل)
                    using (Icon originalIcon = new Icon(iconPath))
                    {
                        // استخراج بزرگترین سایز موجود
                        this.Icon = new Icon(originalIcon, new Size(256, 256));
                    }
                }
                else
                {
                    // جستجو در مسیر BaseDirectory
                    string resourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app_icon.ico");
                    if (File.Exists(resourcePath))
                    {
                        using (Icon originalIcon = new Icon(resourcePath))
                        {
                            this.Icon = new Icon(originalIcon, new Size(256, 256));
                        }
                    }
                    else
                    {
                        // اگر فایل پیدا نشد، از آیکون پیش‌فرض استفاده می‌شود
                        this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                    }
                }
            }
            catch (Exception ex)
            {
                // در صورت بروز خطا، آیکون پیش‌فرض باقی می‌ماند
                System.Diagnostics.Debug.WriteLine($"خطا در بارگذاری آیکون: {ex.Message}");
                // برای تست و دیباگ می‌توانید این خط را فعال کنید:
                // MessageBox.Show($"خطا در بارگذاری آیکون: {ex.Message}", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void InitializeComponent()
        {
            this.Text = "مدیریت پرسنل";
            this.WindowState = FormWindowState.Maximized;
            this.RightToLeft = RightToLeft.Yes;
            this.BackColor = Color.FromArgb(240, 248, 255);

            // پسزمینه گرادیانت
            using (LinearGradientBrush brush = new LinearGradientBrush(this.ClientRectangle, Color.LightBlue, Color.White, LinearGradientMode.Vertical))
            {
                this.BackgroundImage = new Bitmap(this.Width, this.Height);
                using (Graphics g = Graphics.FromImage(this.BackgroundImage))
                {
                    g.FillRectangle(brush, this.ClientRectangle);
                }
            }

            int centerX = (this.ClientSize.Width - 300) / 2;
            int centerY = (this.ClientSize.Height - 450) / 2;

            // سرتیتر
            Label lblTitle = new Label
            {
                Text = "سیستم مدیریت پرسنل",
                Location = new Point(centerX, centerY - 80),
                Size = new Size(300, 50),
                Font = FontSettings.TitleFont,
                ForeColor = Color.Navy,
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblTitle);

            // دکمه ثبت پرسنل جدید
            Button btnAdd = new Button
            {
                Text = "ثبت پرسنل جدید",
                Location = new Point(centerX, centerY),
                Size = new Size(300, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightBlue,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnAdd, 15);
            btnAdd.Click += (s, e) => new FormPersonnelRegister().ShowDialog();
            this.Controls.Add(btnAdd);

            // دکمه ویرایش پرسنل
            Button btnEdit = new Button
            {
                Text = "ویرایش پرسنل",
                Location = new Point(centerX, centerY + 60),
                Size = new Size(300, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightGreen,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnEdit, 15);
            btnEdit.Click += (s, e) => new FormPersonnelEdit().ShowDialog();
            this.Controls.Add(btnEdit);

            // دکمه حذف پرسنل
            Button btnDelete = new Button
            {
                Text = "حذف پرسنل",
                Location = new Point(centerX, centerY + 120),
                Size = new Size(300, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightCoral,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnDelete, 15);
            btnDelete.Click += (s, e) => new FormPersonnelDelete().ShowDialog();
            this.Controls.Add(btnDelete);

            // دکمه جستجوی پرسنل
            Button btnSearch = new Button
            {
                Text = "جستجوی پرسنل",
                Location = new Point(centerX, centerY + 180),
                Size = new Size(300, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.Orange,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnSearch, 15);
            btnSearch.Click += (s, e) => new FormPersonnelSearch().ShowDialog();
            this.Controls.Add(btnSearch);

            // دکمه تحلیل داده‌های پرسنل
            Button btnAnalytics = new Button
            {
                Text = "تحلیل داده‌های پرسنل",
                Location = new Point(centerX, centerY + 240),
                Size = new Size(300, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.SteelBlue,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnAnalytics, 15);
            btnAnalytics.Click += (s, e) => new FormPersonnelAnalytics().ShowDialog();
            this.Controls.Add(btnAnalytics);

            // دکمه تنظیمات (جدید)
            Button btnSettings = new Button
            {
                Text = "تنظیمات",
                Location = new Point(centerX, centerY + 300),
                Size = new Size(300, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.MediumPurple,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnSettings, 15);
            btnSettings.Click += BtnSettings_Click;
            this.Controls.Add(btnSettings);

            // دکمه خروج
            Button btnExit = new Button
            {
                Text = "خروج",
                Location = new Point(centerX, centerY + 360),
                Size = new Size(300, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.Gray,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnExit, 15);
            btnExit.Click += (s, e) => Application.Exit();
            this.Controls.Add(btnExit);
        }

        private void ApplyRoundedCorners(Control control, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            path.AddArc(0, 0, radius, radius, 180, 90);
            path.AddArc(control.Width - radius, 0, radius, radius, 270, 90);
            path.AddArc(control.Width - radius, control.Height - radius, radius, radius, 0, 90);
            path.AddArc(0, control.Height - radius, radius, radius, 90, 90);
            path.CloseFigure();
            control.Region = new Region(path);
        }

        /// <summary>
        /// باز کردن فرم تنظیمات
        /// </summary>
        private void BtnSettings_Click(object sender, EventArgs e)
        {
            FormSettings settingsForm = new FormSettings();
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                // پس از ذخیره تنظیمات، فونت‌ها را دوباره اعمال می‌کنیم
                FontSettings.ApplyFontToForm(this);
            }
        }
    }
}