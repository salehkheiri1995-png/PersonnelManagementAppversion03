using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    public partial class FormSettings : Form
    {
        private TextBox txtDatabasePath;
        private TextBox txtPhotosFolder;
        private Label lblCurrentDatabase;
        private Label lblCurrentPhotos;

        // تنظیمات فونت
        private ComboBox cmbFontFamily;
        private NumericUpDown numTitleSize;
        private NumericUpDown numLabelSize;
        private NumericUpDown numTextBoxSize;
        private NumericUpDown numButtonSize;
        private NumericUpDown numBodySize;
        private NumericUpDown numChartLabelSize;
        private CheckBox chkBoldTitle;
        private CheckBox chkBoldLabel;
        private CheckBox chkBoldButton;
        private CheckBox chkBoldChartLabel;

        // Panels برای هر بخش
        private Panel pnlDatabaseContent;
        private Panel pnlPhotosContent;
        private Panel pnlFontContent;
        private Panel pnlMissingPhotosContent;
        private Panel pnlExcelImportContent;
        private Panel pnlLookupTablesContent;  // ✅ جدید
        private Panel pnlCurrentContent;

        // دکمه‌منوها
        private Panel btnMenuDatabase;
        private Panel btnMenuPhotos;
        private Panel btnMenuFont;
        private Panel btnMenuMissingPhotos;
        private Panel btnMenuExcelImport;
        private Panel btnMenuLookupTables;     // ✅ جدید
        private Panel selectedMenuButton;

        // رنگ‌های مدرن
        private readonly Color PrimaryColor    = Color.FromArgb(33, 150, 243);
        private readonly Color PrimaryDark     = Color.FromArgb(25, 118, 210);
        private readonly Color AccentColor     = Color.FromArgb(76, 175, 80);
        private readonly Color BackgroundColor = Color.FromArgb(250, 250, 250);
        private readonly Color SidebarColor    = Color.FromArgb(248, 249, 250);
        private readonly Color CardBackground  = Color.White;
        private readonly Color TextPrimary     = Color.FromArgb(33, 33, 33);
        private readonly Color TextSecondary   = Color.FromArgb(117, 117, 117);
        private readonly Color DangerColor     = Color.FromArgb(244, 67, 54);
        private readonly Color WarningColor    = Color.FromArgb(255, 152, 0);
        private readonly Color MenuHover       = Color.FromArgb(240, 240, 240);
        private readonly Color MenuSelected    = Color.FromArgb(33, 150, 243);
        private readonly Color ImportColor     = Color.FromArgb(0, 150, 136);
        private readonly Color LookupColor     = Color.FromArgb(156, 39, 176);  // ✅ بنفش برای مدیریت جداول

        public FormSettings()
        {
            InitializeComponent();
            FontSettings.ApplyFontToForm(this);
            LoadCurrentSettings();
            this.Shown += (s, e) => ShowContent(pnlDatabaseContent, btnMenuDatabase);
        }

        private Font GetSafeFont(string familyName, float size, FontStyle style = FontStyle.Regular)
        {
            try   { return new Font(familyName, size, style); }
            catch { return new Font("Tahoma",    size, style); }
        }

        private void InitializeComponent()
        {
            this.Text             = "\u2699\ufe0f تنظیمات برنامه";
            this.Size             = new Size(1000, 760);
            this.StartPosition    = FormStartPosition.CenterScreen;
            this.RightToLeft      = RightToLeft.Yes;
            this.FormBorderStyle  = FormBorderStyle.FixedDialog;
            this.MaximizeBox      = false;
            this.MinimizeBox      = false;
            this.BackColor        = BackgroundColor;
            this.Padding          = new Padding(15);

            // هدر
            Panel headerPanel = CreateHeaderPanel();
            this.Controls.Add(headerPanel);

            // Content Area
            Panel contentArea = new Panel
            {
                Location  = new Point(15, 95),
                Size      = new Size(720, 530),
                BackColor = BackgroundColor
            };
            this.Controls.Add(contentArea);

            // ساخت محتواها
            pnlDatabaseContent      = CreateDatabaseContent();
            pnlPhotosContent        = CreatePhotosContent();
            pnlFontContent          = CreateFontContent();
            pnlMissingPhotosContent = CreateMissingPhotosContent();
            pnlExcelImportContent   = CreateExcelImportContent();
            pnlLookupTablesContent  = CreateLookupTablesContent();  // ✅

            contentArea.Controls.Add(pnlDatabaseContent);
            contentArea.Controls.Add(pnlPhotosContent);
            contentArea.Controls.Add(pnlFontContent);
            contentArea.Controls.Add(pnlMissingPhotosContent);
            contentArea.Controls.Add(pnlExcelImportContent);
            contentArea.Controls.Add(pnlLookupTablesContent);       // ✅

            pnlDatabaseContent.Visible      = true;
            pnlPhotosContent.Visible        = true;
            pnlFontContent.Visible          = true;
            pnlMissingPhotosContent.Visible = true;
            pnlExcelImportContent.Visible   = true;
            pnlLookupTablesContent.Visible  = true;                 // ✅

            // Sidebar
            Panel sidebarPanel = CreateSidebar();
            this.Controls.Add(sidebarPanel);

            // دکمه‌های پایین
            Panel buttonPanel = CreateButtonPanel();
            this.Controls.Add(buttonPanel);
        }

        // ─────────────────────────────────────────────
        // Header
        // ─────────────────────────────────────────────
        private Panel CreateHeaderPanel()
        {
            Panel panel = new Panel
            {
                Location  = new Point(15, 15),
                Size      = new Size(950, 65),
                BackColor = PrimaryColor
            };
            ApplyRoundedCorners(panel, 12);

            panel.Controls.Add(new Label
            {
                Text      = "\u2699\ufe0f تنظیمات برنامه",
                Font      = GetSafeFont(FontSettings.TitleFont?.FontFamily.Name ?? "Tahoma", 18, FontStyle.Bold),
                ForeColor = Color.White,
                Location  = new Point(20, 12),
                Size      = new Size(400, 35),
                TextAlign = ContentAlignment.MiddleRight
            });
            panel.Controls.Add(new Label
            {
                Text      = "مدیریت تنظیمات مسیرها، فونت‌ها و سایر موارد",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = Color.FromArgb(230, 240, 255),
                Location  = new Point(20, 42),
                Size      = new Size(400, 18),
                TextAlign = ContentAlignment.TopRight
            });
            return panel;
        }

        // ─────────────────────────────────────────────
        // Sidebar  (✅ منوی جدید اضافه شد)
        // ─────────────────────────────────────────────
        private Panel CreateSidebar()
        {
            Panel sidebar = new Panel
            {
                Location  = new Point(755, 95),
                Size      = new Size(210, 530),
                BackColor = SidebarColor
            };
            ApplyRoundedCorners(sidebar, 10);

            int yPos = 20;

            sidebar.Controls.Add(new Label
            {
                Text      = "بخش‌ها",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextSecondary,
                Location  = new Point(15, yPos),
                Size      = new Size(180, 25),
                TextAlign = ContentAlignment.MiddleRight
            });
            yPos += 40;

            btnMenuDatabase = CreateMenuButton("\ud83d\udcbe تنظیمات دیتابیس", yPos, pnlDatabaseContent);
            sidebar.Controls.Add(btnMenuDatabase);
            yPos += 55;

            btnMenuPhotos = CreateMenuButton("\ud83d\uddbc\ufe0f تنظیمات عکس‌ها", yPos, pnlPhotosContent);
            sidebar.Controls.Add(btnMenuPhotos);
            yPos += 55;

            btnMenuFont = CreateMenuButton("\ud83d\udd24 تنظیمات فونت", yPos, pnlFontContent);
            sidebar.Controls.Add(btnMenuFont);
            yPos += 55;

            btnMenuMissingPhotos = CreateMenuButton("\ud83d\udcf8 پرسنل بدون عکس", yPos, pnlMissingPhotosContent);
            sidebar.Controls.Add(btnMenuMissingPhotos);
            yPos += 55;

            btnMenuExcelImport = CreateMenuButton("\ud83d\udce5 وارد کردن از اکسل", yPos, pnlExcelImportContent);
            sidebar.Controls.Add(btnMenuExcelImport);
            yPos += 55;

            // ✅ منوی جدید مدیریت جداول مرجع
            btnMenuLookupTables = CreateMenuButton("\ud83d\uddc2\ufe0f مدیریت جداول مرجع", yPos, pnlLookupTablesContent);
            sidebar.Controls.Add(btnMenuLookupTables);

            return sidebar;
        }

        private Panel CreateMenuButton(string text, int yPos, Panel targetContent)
        {
            Panel btn = new Panel
            {
                Location  = new Point(10, yPos),
                Size      = new Size(190, 45),
                BackColor = Color.Transparent,
                Cursor    = Cursors.Hand,
                Tag       = "menu"
            };
            ApplyRoundedCorners(btn, 8);

            Label lbl = new Label
            {
                Text      = text,
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10),
                ForeColor = TextPrimary,
                Location  = new Point(10, 0),
                Size      = new Size(170, 45),
                TextAlign = ContentAlignment.MiddleRight,
                Cursor    = Cursors.Hand
            };
            btn.Controls.Add(lbl);

            EventHandler clickHandler = (s, e) => ShowContent(targetContent, btn);
            btn.Click += clickHandler;
            lbl.Click += clickHandler;

            btn.MouseEnter += (s, e) => { if (selectedMenuButton != btn) btn.BackColor = MenuHover; };
            btn.MouseLeave += (s, e) => { if (selectedMenuButton != btn) btn.BackColor = Color.Transparent; };
            lbl.MouseEnter += (s, e) => btn.BackColor = selectedMenuButton == btn ? MenuSelected : MenuHover;
            lbl.MouseLeave += (s, e) => btn.BackColor = selectedMenuButton == btn ? MenuSelected : Color.Transparent;

            return btn;
        }

        private void ShowContent(Panel contentPanel, Panel menuButton)
        {
            // مخفی کردن همه
            pnlDatabaseContent?.let(p      => p.Visible = false);
            pnlPhotosContent?.let(p        => p.Visible = false);
            pnlFontContent?.let(p          => p.Visible = false);
            pnlMissingPhotosContent?.let(p => p.Visible = false);
            pnlExcelImportContent?.let(p   => p.Visible = false);
            pnlLookupTablesContent?.let(p  => p.Visible = false);   // ✅

            if (contentPanel != null)
            {
                contentPanel.Visible = true;
                contentPanel.BringToFront();
                contentPanel.Invalidate(true);
                this.Refresh();
                pnlCurrentContent = contentPanel;
            }

            // برداشتن هایلایت
            if (selectedMenuButton != null)
            {
                selectedMenuButton.BackColor = Color.Transparent;
                foreach (Control c in selectedMenuButton.Controls)
                    if (c is Label l) l.ForeColor = TextPrimary;
            }

            // هایلایت دکمه انتخابی
            if (menuButton != null)
            {
                menuButton.BackColor = MenuSelected;
                foreach (Control c in menuButton.Controls)
                    if (c is Label l) l.ForeColor = Color.White;
                selectedMenuButton = menuButton;
            }
        }

        // ─────────────────────────────────────────────
        // ✅ محتوای بخش مدیریت جداول مرجع
        // ─────────────────────────────────────────────
        private Panel CreateLookupTablesContent()
        {
            Panel content = new Panel
            {
                Location   = new Point(0, 0),
                Size       = new Size(720, 530),
                BackColor  = Color.Transparent,
                AutoScroll = true
            };

            Panel card = new Panel
            {
                Location  = new Point(10, 10),
                Size      = new Size(690, 750),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(card, 10);
            ApplyCardShadow(card);

            // عنوان
            card.Controls.Add(new Label
            {
                Text      = "🗂️ مدیریت جداول مرجع",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 14, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(420, 20),
                Size      = new Size(250, 35),
                TextAlign = ContentAlignment.MiddleRight
            });

            card.Controls.Add(new Label
            {
                Text      = "اضافه کردن، ویرایش و حذف رکوردهای جداول پایه سیستم",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                Location  = new Point(380, 55),
                Size      = new Size(290, 20),
                TextAlign = ContentAlignment.TopRight
            });

            int yPos = 95;
            int btnWidth = 200;
            int btnHeight = 40;
            int spacing = 15;
            int col1X = 470;
            int col2X = 240;
            int col3X = 10;

            // ردیف اول
            AddLookupTableButton(card, "🏢 استان‌ها", col1X, yPos, btnWidth, btnHeight, "Provinces", "ProvinceID", "ProvinceName", "استان‌ها");
            AddLookupTableButton(card, "🏙️ شهرها", col2X, yPos, btnWidth, btnHeight, "Cities", "CityID", "CityName", "شهرها");
            AddLookupTableButton(card, "📋 امور انتقال", col3X, yPos, btnWidth, btnHeight, "TransferAffairs", "AffairID", "AffairName", "امور انتقال");
            yPos += btnHeight + spacing;

            // ردیف دوم
            AddLookupTableButton(card, "🏛️ ادارات بهره‌برداری", col1X, yPos, btnWidth, btnHeight, "OperationDepartments", "DeptID", "DeptName", "ادارات بهره‌برداری");
            AddLookupTableButton(card, "📍 نواحی", col2X, yPos, btnWidth, btnHeight, "Districts", "DistrictID", "DistrictName", "نواحی");
            AddLookupTableButton(card, "🏭 نام پست‌ها", col3X, yPos, btnWidth, btnHeight, "PostsNames", "PostNameID", "PostName", "نام پست‌ها");
            yPos += btnHeight + spacing;

            // ردیف سوم
            AddLookupTableButton(card, "⚡ سطح ولتاژ", col1X, yPos, btnWidth, btnHeight, "VoltageLevels", "VoltageID", "VoltageName", "سطح ولتاژ");
            AddLookupTableButton(card, "📊 استانداردهای پست", col2X, yPos, btnWidth, btnHeight, "PostStandards", "StandardID", "StandardName", "استانداردهای پست");
            AddLookupTableButton(card, "🏗️ نوع پست", col3X, yPos, btnWidth, btnHeight, "PostTypes", "TypeID", "TypeName", "نوع پست");
            yPos += btnHeight + spacing;

            // ردیف چهارم
            AddLookupTableButton(card, "🔌 اتصالات توزیع شده", col1X, yPos, btnWidth, btnHeight, "DistributedConnections", "ConnID", "ConnName", "اتصالات توزیع شده");
            AddLookupTableButton(card, "🛡️ انواع عایق", col2X, yPos, btnWidth, btnHeight, "InsulationTypes", "InsID", "InsName", "انواع عایق");
            AddLookupTableButton(card, "🔧 نوع پست دو", col3X, yPos, btnWidth, btnHeight, "PostTypeTwos", "PT2ID", "PT2Name", "نوع پست دو");
            yPos += btnHeight + spacing;

            // ردیف پنجم
            AddLookupTableButton(card, "📱 ثابت/سیار", col1X, yPos, btnWidth, btnHeight, "FixedMobiles", "FMID", "FMName", "ثابت/سیار");
            AddLookupTableButton(card, "🔗 وضعیت مدار", col2X, yPos, btnWidth, btnHeight, "CircuitStatuses", "CircuitID", "CircuitName", "وضعیت مدار");
            AddLookupTableButton(card, "⚙️ دیزل ژنراتورها", col3X, yPos, btnWidth, btnHeight, "DieselGenerators", "DieselID", "DieselName", "دیزل ژنراتورها");
            yPos += btnHeight + spacing;

            // ردیف ششم
            AddLookupTableButton(card, "🔋 فیدرهای توزیع", col1X, yPos, btnWidth, btnHeight, "DistributionFeeds", "FeedID", "FeedName", "فیدرهای توزیع");
            AddLookupTableButton(card, "💧 وضعیت آب", col2X, yPos, btnWidth, btnHeight, "WaterStatuses", "WaterID", "WaterName", "وضعیت آب");
            AddLookupTableButton(card, "🏠 مهمان‌خانه‌ها", col3X, yPos, btnWidth, btnHeight, "GuestHouses", "GuestID", "GuestName", "مهمان‌خانه‌ها");
            yPos += btnHeight + spacing;

            // ردیف هفتم - جداول پرسنلی
            AddLookupTableButton(card, "🕐 شیفت کاری", col1X, yPos, btnWidth, btnHeight, "WorkShift", "WorkShiftID", "WorkShiftName", "شیفت کاری");
            AddLookupTableButton(card, "👤 جنسیت", col2X, yPos, btnWidth, btnHeight, "Gender", "GenderID", "GenderName", "جنسیت");
            AddLookupTableButton(card, "📝 نوع قرارداد", col3X, yPos, btnWidth, btnHeight, "ContractType", "ContractTypeID", "ContractTypeName", "نوع قرارداد");
            yPos += btnHeight + spacing;

            // ردیف هشتم
            AddLookupTableButton(card, "📊 سطح شغلی", col1X, yPos, btnWidth, btnHeight, "JobLevel", "JobLevelID", "JobLevelName", "سطح شغلی");
            AddLookupTableButton(card, "🏢 شرکت‌ها", col2X, yPos, btnWidth, btnHeight, "Company", "CompanyID", "CompanyName", "شرکت‌ها");
            AddLookupTableButton(card, "🎓 مدرک تحصیلی", col3X, yPos, btnWidth, btnHeight, "Degree", "DegreeID", "DegreeName", "مدرک تحصیلی");
            yPos += btnHeight + spacing;

            // ردیف نهم
            AddLookupTableButton(card, "📚 رشته تحصیلی", col1X, yPos, btnWidth, btnHeight, "DegreeField", "DegreeFieldID", "DegreeFieldName", "رشته تحصیلی");
            AddLookupTableButton(card, "✅ وضعیت حضور", col2X, yPos, btnWidth, btnHeight, "StatusPresence", "StatusID", "StatusName", "وضعیت حضور");
            AddLookupTableButton(card, "📋 چارت امور", col3X, yPos, btnWidth, btnHeight, "ChartAffairs", "ChartID", "ChartName", "چارت امور");

            content.Controls.Add(card);
            return content;
        }

        private void AddLookupTableButton(Panel parent, string text, int x, int y, int width, int height,
            string tableName, string idColumn, string nameColumn, string displayName)
        {
            Button btn = CreateModernButton(text, LookupColor, width, height);
            btn.Location = new Point(x, y);
            btn.Font = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 9, FontStyle.Bold);
            btn.Click += (s, e) =>
            {
                try
                {
                    FormLookupTableManager lookupForm = new FormLookupTableManager(tableName, idColumn, nameColumn, displayName);
                    lookupForm.ShowDialog(this);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"❌ خطا در باز کردن فرم:\n\n{ex.Message}\n\n{ex.StackTrace}",
                        "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };
            parent.Controls.Add(btn);
        }

        // ─────────────────────────────────────────────
        // ✅ محتوای بخش وارد کردن از اکسل
        // ─────────────────────────────────────────────
        private Panel CreateExcelImportContent()
        {
            Panel content = new Panel
            {
                Location   = new Point(0, 0),
                Size       = new Size(720, 530),
                BackColor  = Color.Transparent,
                AutoScroll = false
            };

            Panel card = new Panel
            {
                Location  = new Point(10, 10),
                Size      = new Size(690, 340),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(card, 10);
            ApplyCardShadow(card);

            // عنوان
            card.Controls.Add(new Label
            {
                Text      = "\ud83d\udce5 وارد کردن اطلاعات از اکسل",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 14, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(270, 20),
                Size      = new Size(400, 35),
                TextAlign = ContentAlignment.MiddleRight
            });

            card.Controls.Add(new Label
            {
                Text      = "اطلاعات پرسنلی را از فایل اکسل به دیتابیس وارد کنید",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                Location  = new Point(230, 55),
                Size      = new Size(440, 20),
                TextAlign = ContentAlignment.TopRight
            });

            // جدول ستون‌های مورد انتظار
            Panel infoBox = new Panel
            {
                Location  = new Point(15, 88),
                Size      = new Size(658, 135),
                BackColor = Color.FromArgb(232, 248, 232)
            };
            ApplyRoundedCorners(infoBox, 8);

            infoBox.Controls.Add(new Label
            {
                Text      = "\u2139\ufe0f  ستون‌های مورد انتظار در فایل اکسل (ردیف اول = سرستون):",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9, FontStyle.Bold),
                ForeColor = Color.FromArgb(27, 94, 32),
                Location  = new Point(10, 8),
                Size      = new Size(638, 20),
                TextAlign = ContentAlignment.MiddleRight
            });

            infoBox.Controls.Add(new Label
            {
                Text =
                    "استان | شهر | امور انتقال | اداره | ناحیه | نام پست | سطح ولتاژ | شیفت | جنسیت\n" +
                    "نام | نام خانوادگی | نام پدر | ش.پرسنلی | کدملی | موبایل\n" +
                    "تاریخ تولد | تاریخ استخدام | تاریخ شروع بکار | نوع قرارداد | سطح شغل\n" +
                    "شرکت | مدرک تحصیلی | رشته تحصیلی | عنوان شغلی اصلی | فعالیت فعلی | وضعیت",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8.5f),
                ForeColor = Color.FromArgb(46, 125, 50),
                Location  = new Point(10, 32),
                Size      = new Size(638, 96),
                TextAlign = ContentAlignment.TopRight
            });

            card.Controls.Add(infoBox);

            // قوانین پیش‌فرض
            Panel rulesBox = new Panel
            {
                Location  = new Point(15, 233),
                Size      = new Size(658, 70),
                BackColor = Color.FromArgb(255, 243, 224)
            };
            ApplyRoundedCorners(rulesBox, 8);

            rulesBox.Controls.Add(new Label
            {
                Text =
                    "\u26a0\ufe0f  قوانین پیش‌فرض هنگام وارد کردن:\n" +
                    "• تاریخ خالی ← 1300/01/01     • عنوان شغل خالی ← غیرمرتبط     • سایر فیلدهای خالی ← داده‌ای وجود ندارد\n" +
                    "• کدملی تکراری: ردیف نادیده گرفته می‌شود     • تاریخ میلادی: به شمسی تبدیل می‌شود",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8.5f),
                ForeColor = Color.FromArgb(230, 81, 0),
                Location  = new Point(8, 6),
                Size      = new Size(642, 60),
                TextAlign = ContentAlignment.TopRight
            });

            card.Controls.Add(rulesBox);

            // دکمه باز کردن فرم import
            Button btnOpen = CreateModernButton("\ud83d\udce5  باز کردن پنجره وارد کردن از اکسل", ImportColor, 380, 46);
            btnOpen.Location = new Point(155, 288);
            btnOpen.Font     = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 11, FontStyle.Bold);
            btnOpen.Click   += BtnOpenExcelImport_Click;
            card.Controls.Add(btnOpen);

            content.Controls.Add(card);
            return content;
        }

        private void BtnOpenExcelImport_Click(object sender, EventArgs e)
        {
            try
            {
                using var frm = new FormExcelImport();
                frm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"\u274c خطا در باز کردن فرم وارد کردن:\n\n{ex.Message}\n\n{ex.StackTrace}",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ─────────────────────────────────────────────
        // بخش پرسنل بدون عکس
        // ─────────────────────────────────────────────
        private Panel CreateMissingPhotosContent()
        {
            Panel content = new Panel
            {
                Location   = new Point(0, 0),
                Size       = new Size(720, 530),
                BackColor  = Color.Transparent,
                AutoScroll = false
            };

            Panel card = new Panel
            {
                Location  = new Point(10, 10),
                Size      = new Size(690, 280),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(card, 10);
            ApplyCardShadow(card);

            card.Controls.Add(new Label
            {
                Text      = "\ud83d\udcf8 بررسی پرسنل بدون عکس",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 14, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(420, 20),
                Size      = new Size(250, 35),
                TextAlign = ContentAlignment.MiddleRight
            });

            card.Controls.Add(new Label
            {
                Text      = "لیست پرسنلی که عکس پرسنلی ندارند را مشاهده کنید\nو اقدامات لازم را انجام دهید.",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                Location  = new Point(380, 55),
                Size      = new Size(290, 40),
                TextAlign = ContentAlignment.TopRight
            });

            card.Controls.Add(new Label
            {
                Text =
                    "\u2705 مشاهده لیست کامل پرسنل بدون عکس\n" +
                    "\u2705 خروجی اکسل برای گزارش‌گیری\n" +
                    "\u2705 امکان ویرایش و حذف مستقیم\n" +
                    "\u2705 بروزرسانی لحظه‌ای اطلاعات",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9.5f),
                ForeColor = TextSecondary,
                Location  = new Point(380, 105),
                Size      = new Size(290, 90),
                TextAlign = ContentAlignment.TopRight
            });

            Button btnOpenMissingPhotos = CreateModernButton("\ud83d\udccb مشاهده لیست پرسنل بدون عکس", AccentColor, 320, 50);
            btnOpenMissingPhotos.Location = new Point(185, 210);
            btnOpenMissingPhotos.Font     = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 11, FontStyle.Bold);
            btnOpenMissingPhotos.Click   += BtnOpenMissingPhotos_Click;
            card.Controls.Add(btnOpenMissingPhotos);

            content.Controls.Add(card);
            return content;
        }

        private void BtnOpenMissingPhotos_Click(object sender, EventArgs e)
        {
            try
            {
                FormMissingPhotos missingPhotosForm = new FormMissingPhotos();
                missingPhotosForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"\u274c خطا در باز کردن فرم:\n\n{ex.Message}\n\n{ex.StackTrace}",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ─────────────────────────────────────────────
        // بخش دیتابیس
        // ─────────────────────────────────────────────
        private Panel CreateDatabaseContent()
        {
            Panel content = new Panel
            {
                Location   = new Point(0, 0),
                Size       = new Size(720, 530),
                BackColor  = Color.Transparent,
                AutoScroll = false
            };

            Panel card = new Panel
            {
                Location  = new Point(10, 10),
                Size      = new Size(690, 180),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(card, 10);
            ApplyCardShadow(card);

            card.Controls.Add(new Label
            {
                Text      = "\ud83d\udcbe تنظیمات دیتابیس",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 14, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(480, 20),
                Size      = new Size(190, 35),
                TextAlign = ContentAlignment.MiddleRight
            });

            card.Controls.Add(new Label
            {
                Text      = "مسیر فایل دیتابیس Access را انتخاب کنید",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                Location  = new Point(480, 50),
                Size      = new Size(190, 20),
                TextAlign = ContentAlignment.TopRight
            });

            card.Controls.Add(new Label
            {
                Text      = "مسیر فایل:",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(600, 90),
                Size      = new Size(70, 25),
                TextAlign = ContentAlignment.MiddleRight
            });

            txtDatabasePath = new TextBox
            {
                Location    = new Point(122, 92),
                Size        = new Size(470, 28),
                Font        = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 9),
                ReadOnly    = true,
                BackColor   = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };
            card.Controls.Add(txtDatabasePath);

            Button btnBrowse = CreateModernButton("\ud83d\udd0d جستجو", PrimaryColor, 100, 28);
            btnBrowse.Location = new Point(15, 92);
            btnBrowse.Click   += BtnBrowseDatabase_Click;
            card.Controls.Add(btnBrowse);

            lblCurrentDatabase = new Label
            {
                Location  = new Point(122, 125),
                Size      = new Size(470, 20),
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 7.5f),
                ForeColor = TextSecondary
            };
            card.Controls.Add(lblCurrentDatabase);

            content.Controls.Add(card);
            return content;
        }

        // ─────────────────────────────────────────────
        // بخش عکس‌ها
        // ─────────────────────────────────────────────
        private Panel CreatePhotosContent()
        {
            Panel content = new Panel
            {
                Location   = new Point(0, 0),
                Size       = new Size(720, 530),
                BackColor  = Color.Transparent,
                AutoScroll = false
            };

            Panel card = new Panel
            {
                Location  = new Point(10, 10),
                Size      = new Size(690, 180),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(card, 10);
            ApplyCardShadow(card);

            card.Controls.Add(new Label
            {
                Text      = "\ud83d\uddbc\ufe0f تنظیمات عکس‌ها",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 14, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(480, 20),
                Size      = new Size(190, 35),
                TextAlign = ContentAlignment.MiddleRight
            });

            card.Controls.Add(new Label
            {
                Text      = "پوشه ذخیره عکس پرسنل را مشخص کنید",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                Location  = new Point(480, 50),
                Size      = new Size(190, 20),
                TextAlign = ContentAlignment.TopRight
            });

            card.Controls.Add(new Label
            {
                Text      = "مسیر پوشه:",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(600, 90),
                Size      = new Size(70, 25),
                TextAlign = ContentAlignment.MiddleRight
            });

            txtPhotosFolder = new TextBox
            {
                Location    = new Point(122, 92),
                Size        = new Size(470, 28),
                Font        = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 9),
                ReadOnly    = true,
                BackColor   = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };
            card.Controls.Add(txtPhotosFolder);

            Button btnBrowse = CreateModernButton("\ud83d\udd0d جستجو", PrimaryColor, 100, 28);
            btnBrowse.Location = new Point(15, 92);
            btnBrowse.Click   += BtnBrowsePhotos_Click;
            card.Controls.Add(btnBrowse);

            lblCurrentPhotos = new Label
            {
                Location  = new Point(122, 125),
                Size      = new Size(470, 20),
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 7.5f),
                ForeColor = TextSecondary
            };
            card.Controls.Add(lblCurrentPhotos);

            content.Controls.Add(card);
            return content;
        }

        // ─────────────────────────────────────────────
        // بخش فونت
        // ─────────────────────────────────────────────
        private Panel CreateFontContent()
        {
            Panel content = new Panel
            {
                Location   = new Point(0, 0),
                Size       = new Size(720, 530),
                BackColor  = Color.Transparent,
                AutoScroll = false
            };

            Panel card = new Panel
            {
                Location  = new Point(10, 10),
                Size      = new Size(690, 450),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(card, 10);
            ApplyCardShadow(card);

            card.Controls.Add(new Label
            {
                Text      = "\ud83d\udd24 تنظیمات فونت",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 14, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(480, 20),
                Size      = new Size(190, 35),
                TextAlign = ContentAlignment.MiddleRight
            });

            card.Controls.Add(new Label
            {
                Text      = "نوع و اندازه فونت‌های برنامه را تنظیم کنید",
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                Location  = new Point(430, 50),
                Size      = new Size(240, 20),
                TextAlign = ContentAlignment.TopRight
            });

            int yPos = 85;

            card.Controls.Add(new Label
            {
                Text      = "نوع فونت:",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(600, yPos),
                Size      = new Size(70, 25),
                TextAlign = ContentAlignment.MiddleRight
            });

            cmbFontFamily = new ComboBox
            {
                Location      = new Point(350, yPos),
                Size          = new Size(240, 28),
                Font          = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 9),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle     = FlatStyle.Flat
            };
            cmbFontFamily.Items.AddRange(new string[] {
                "Tahoma", "Arial", "Segoe UI", "Calibri", "Times New Roman",
                "B Nazanin", "B Mitra", "B Lotus", "B Titr", "IRANSans", "Vazir"
            });
            card.Controls.Add(cmbFontFamily);
            yPos += 50;

            Panel divider = new Panel { Location = new Point(30, yPos), Size = new Size(630, 1), BackColor = Color.FromArgb(230, 230, 230) };
            card.Controls.Add(divider);
            yPos += 20;

            card.Controls.Add(new Label
            {
                Text      = "اندازه فونت‌ها:",
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location  = new Point(555, yPos),
                Size      = new Size(115, 25),
                TextAlign = ContentAlignment.MiddleRight
            });
            yPos += 35;

            int col1X = 460, col2X = 250, col3X = 40, labelW = 90, numW = 60, checkW = 60;

            AddFontSizeRowCompact(card, "سرتیتر:",   col1X, yPos, labelW, out numTitleSize,    out chkBoldTitle,      numW, checkW, 16);
            AddFontSizeRowCompact(card, "برچسب:",    col2X, yPos, labelW, out numLabelSize,    out chkBoldLabel,      numW, checkW, 12);
            AddFontSizeRowCompact(card, "دکمه:",     col3X, yPos, labelW, out numButtonSize,   out chkBoldButton,     numW, checkW, 12);
            yPos += 40;

            AddFontSizeRowCompactNoCheckbox(card, "متن:",      col1X, yPos, labelW, out numTextBoxSize, numW, checkW, 11);
            AddFontSizeRowCompactNoCheckbox(card, "متن عادی:", col2X, yPos, labelW, out numBodySize,    numW, checkW, 10);
            yPos += 40;

            AddFontSizeRowCompact(card, "نمودارها:", col1X, yPos, labelW, out numChartLabelSize, out chkBoldChartLabel, numW, checkW, 10);

            content.Controls.Add(card);
            return content;
        }

        private void AddFontSizeRowCompact(Panel parent, string label, int x, int y, int labelW,
            out NumericUpDown numeric, out CheckBox checkbox, int numW, int checkW, int defaultValue)
        {
            parent.Controls.Add(new Label
            {
                Text      = label,
                Location  = new Point(x + numW + checkW + 5, y + 2),
                Size      = new Size(labelW, 22),
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.MiddleRight
            });

            numeric = new NumericUpDown
            {
                Location    = new Point(x + checkW + 3, y),
                Size        = new Size(numW, 26),
                Minimum     = 8, Maximum = 72, Value = defaultValue,
                Font        = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                BorderStyle = BorderStyle.FixedSingle
            };
            parent.Controls.Add(numeric);

            checkbox = new CheckBox
            {
                Text      = "ضخیم",
                Location  = new Point(x, y + 2),
                Size      = new Size(checkW, 22),
                Font      = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8),
                ForeColor = TextSecondary
            };
            parent.Controls.Add(checkbox);
        }

        private void AddFontSizeRowCompactNoCheckbox(Panel parent, string label, int x, int y, int labelW,
            out NumericUpDown numeric, int numW, int checkW, int defaultValue)
        {
            parent.Controls.Add(new Label
            {
                Text      = label,
                Location  = new Point(x + numW + checkW + 5, y + 2),
                Size      = new Size(labelW, 22),
                Font      = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.MiddleRight
            });

            numeric = new NumericUpDown
            {
                Location    = new Point(x + checkW + 3, y),
                Size        = new Size(numW, 26),
                Minimum     = 8, Maximum = 72, Value = defaultValue,
                Font        = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                BorderStyle = BorderStyle.FixedSingle
            };
            parent.Controls.Add(numeric);
        }

        // ─────────────────────────────────────────────
        // دکمه‌های پایین
        // ─────────────────────────────────────────────
        private Panel CreateButtonPanel()
        {
            Panel panel = new Panel
            {
                Location  = new Point(15, 635),
                Size      = new Size(950, 60),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(panel, 10);
            ApplyCardShadow(panel);

            int centerX     = panel.Width / 2;
            int buttonWidth = 130;
            int buttonHeight= 38;
            int spacing     = 12;

            Button btnSave = CreateModernButton("\ud83d\udcbe ذخیره", AccentColor, buttonWidth, buttonHeight);
            btnSave.Location = new Point(centerX - buttonWidth / 2, 11);
            btnSave.Font     = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            btnSave.Click   += BtnSave_Click;
            panel.Controls.Add(btnSave);

            Button btnReset = CreateModernButton("\ud83d\udd04 بازنشانی", WarningColor, buttonWidth, buttonHeight);
            btnReset.Location = new Point(centerX + buttonWidth / 2 + spacing, 11);
            btnReset.Font     = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            btnReset.Click   += BtnReset_Click;
            panel.Controls.Add(btnReset);

            Button btnCancel = CreateModernButton("\u274c لغو", DangerColor, buttonWidth, buttonHeight);
            btnCancel.Location = new Point(centerX - buttonWidth / 2 - buttonWidth - spacing, 11);
            btnCancel.Font     = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            btnCancel.Click   += (s, e) => this.Close();
            panel.Controls.Add(btnCancel);

            return panel;
        }

        private Button CreateModernButton(string text, Color backColor, int width, int height)
        {
            Button btn = new Button
            {
                Text      = text,
                Size      = new Size(width, height),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor    = Cursors.Hand,
                Font      = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10)
            };
            btn.FlatAppearance.BorderSize = 0;
            ApplyRoundedCorners(btn, 8);

            Color orig = backColor;
            btn.MouseEnter += (s, e) => btn.BackColor = ControlPaint.Light(orig, 0.1f);
            btn.MouseLeave += (s, e) => btn.BackColor = orig;
            return btn;
        }

        // ─────────────────────────────────────────────
        // بارگذاری تنظیمات
        // ─────────────────────────────────────────────
        private void LoadCurrentSettings()
        {
            txtDatabasePath.Text  = AppSettings.DatabasePath;
            txtPhotosFolder.Text  = AppSettings.PhotosFolder;
            lblCurrentDatabase.Text = $"\ud83d\udcc2 {AppSettings.DatabasePath}";
            lblCurrentPhotos.Text   = $"\ud83d\udcc2 {AppSettings.PhotosFolder}";

            cmbFontFamily.Text      = FontSettings.FontFamilyName;
            numTitleSize.Value      = (decimal)FontSettings.TitleFontSize;
            numLabelSize.Value      = (decimal)FontSettings.LabelFontSize;
            numTextBoxSize.Value    = (decimal)FontSettings.TextBoxFontSize;
            numButtonSize.Value     = (decimal)FontSettings.ButtonFontSize;
            numBodySize.Value       = (decimal)FontSettings.BodyFontSize;
            numChartLabelSize.Value = (decimal)FontSettings.ChartLabelFontSize;
            chkBoldTitle.Checked        = FontSettings.TitleFontBold;
            chkBoldLabel.Checked        = FontSettings.LabelFontBold;
            chkBoldButton.Checked       = FontSettings.ButtonFontBold;
            chkBoldChartLabel.Checked   = FontSettings.ChartLabelFontBold;
        }

        private void BtnBrowseDatabase_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter           = "Access Database (*.accdb)|*.accdb|All Files (*.*)|*.*";
                ofd.Title            = "انتخاب فایل دیتابیس";
                ofd.InitialDirectory = Path.GetDirectoryName(AppSettings.DatabasePath);
                if (ofd.ShowDialog() == DialogResult.OK)
                    txtDatabasePath.Text = ofd.FileName;
            }
        }

        private void BtnBrowsePhotos_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description     = "انتخاب پوشه عکس‌ها";
                fbd.SelectedPath    = AppSettings.PhotosFolder;
                fbd.ShowNewFolderButton = true;
                if (fbd.ShowDialog() == DialogResult.OK)
                    txtPhotosFolder.Text = fbd.SelectedPath;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (!File.Exists(txtDatabasePath.Text))
                {
                    if (MessageBox.Show(
                        "\u26a0\ufe0f فایل دیتابیس در مسیر انتخابی وجود ندارد.\n\nآیا می‌خواهید ادامه دهید؟",
                        "هشدار", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                        return;
                }

                if (!Directory.Exists(txtPhotosFolder.Text))
                {
                    if (MessageBox.Show(
                        "\ud83d\udcc1 پوشه عکس‌ها وجود ندارد.\n\nآیا می‌خواهید آن را ایجاد کنید؟",
                        "پرسش", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        Directory.CreateDirectory(txtPhotosFolder.Text);
                    else
                        return;
                }

                AppSettings.DatabasePath = txtDatabasePath.Text;
                AppSettings.PhotosFolder = txtPhotosFolder.Text;

                FontSettings.FontFamilyName     = cmbFontFamily.Text;
                FontSettings.TitleFontSize      = (float)numTitleSize.Value;
                FontSettings.LabelFontSize      = (float)numLabelSize.Value;
                FontSettings.TextBoxFontSize    = (float)numTextBoxSize.Value;
                FontSettings.ButtonFontSize     = (float)numButtonSize.Value;
                FontSettings.BodyFontSize       = (float)numBodySize.Value;
                FontSettings.ChartLabelFontSize = (float)numChartLabelSize.Value;
                FontSettings.TitleFontBold      = chkBoldTitle.Checked;
                FontSettings.LabelFontBold      = chkBoldLabel.Checked;
                FontSettings.ButtonFontBold     = chkBoldButton.Checked;
                FontSettings.ChartLabelFontBold = chkBoldChartLabel.Checked;

                FontSettings.SaveSettings();

                MessageBox.Show(
                    "\u2705 تنظیمات با موفقیت ذخیره شد!\n\n\ud83d\udd04 لطفاً برنامه را مجدداٌ راه‌اندازی کنید.",
                    "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"\u274c خطا در ذخیره تنظیمات:\n\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(
                "\u26a0\ufe0f آیا مطمئن هستید که می‌خواهید تنظیمات را به حالت پیش‌فرض برگردانید?\n\nتمامی تغییرات از بین خواهد رفت!",
                "تایید بازنشانی", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                AppSettings.ResetToDefaults();
                FontSettings.ResetToDefaults();
                LoadCurrentSettings();

                MessageBox.Show(
                    "\u2705 تنظیمات به حالت پیش‌فرض برگشت!\n\n\ud83d\udd04 برای اعمال تغییرات، برنامه را مجدداً راه‌اندازی کنید.",
                    "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // ─────────────────────────────────────────────
        // ابزارهای کمکی UI
        // ─────────────────────────────────────────────
        private void ApplyRoundedCorners(Control control, int radius)
        {
            try
            {
                GraphicsPath path = new GraphicsPath();
                path.AddArc(0, 0, radius, radius, 180, 90);
                path.AddArc(control.Width  - radius, 0,                    radius, radius, 270, 90);
                path.AddArc(control.Width  - radius, control.Height - radius, radius, radius, 0,   90);
                path.AddArc(0,                        control.Height - radius, radius, radius, 90,  90);
                path.CloseFigure();
                control.Region = new Region(path);
            }
            catch { }
        }

        private void ApplyCardShadow(Panel panel)
        {
            panel.Paint += (s, e) =>
            {
                using (SolidBrush shadowBrush = new SolidBrush(Color.FromArgb(10, 0, 0, 0)))
                    e.Graphics.FillRectangle(shadowBrush, new Rectangle(3, 3, panel.Width - 3, panel.Height - 3));
            };
        }
    }

    // ✅ Extension method کوچک برای null-safe invoke
    internal static class PanelExtensions
    {
        internal static void let(this Panel p, Action<Panel> action)
        {
            if (p != null) action(p);
        }
    }
}