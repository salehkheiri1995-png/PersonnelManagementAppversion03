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
        private CheckBox chkBoldTitle;
        private CheckBox chkBoldLabel;
        private CheckBox chkBoldButton;

        // Panels برای هر بخش
        private Panel pnlDatabaseContent;
        private Panel pnlPhotosContent;
        private Panel pnlFontContent;
        private Panel pnlCurrentContent;

        // دکمه‌منوها
        private Panel btnMenuDatabase;
        private Panel btnMenuPhotos;
        private Panel btnMenuFont;
        private Panel selectedMenuButton;

        // رنگ‌های مدرن
        private readonly Color PrimaryColor = Color.FromArgb(33, 150, 243);
        private readonly Color PrimaryDark = Color.FromArgb(25, 118, 210);
        private readonly Color AccentColor = Color.FromArgb(76, 175, 80);
        private readonly Color BackgroundColor = Color.FromArgb(250, 250, 250);
        private readonly Color SidebarColor = Color.FromArgb(248, 249, 250);
        private readonly Color CardBackground = Color.White;
        private readonly Color TextPrimary = Color.FromArgb(33, 33, 33);
        private readonly Color TextSecondary = Color.FromArgb(117, 117, 117);
        private readonly Color DangerColor = Color.FromArgb(244, 67, 54);
        private readonly Color WarningColor = Color.FromArgb(255, 152, 0);
        private readonly Color MenuHover = Color.FromArgb(240, 240, 240);
        private readonly Color MenuSelected = Color.FromArgb(33, 150, 243);

        public FormSettings()
        {
            InitializeComponent();
            FontSettings.ApplyFontToForm(this);
            LoadCurrentSettings();
            // نمایش پیش‌فرض بخش دیتابیس بعد از نمایش فرم تا رندر صحیح انجام شود
            this.Shown += (s, e) => ShowContent(pnlDatabaseContent, btnMenuDatabase);
        }

        // متد کمکی برای گرفتن فونت با فالبک
        private Font GetSafeFont(string familyName, float size, FontStyle style = FontStyle.Regular)
        {
            try
            {
                return new Font(familyName, size, style);
            }
            catch
            {
                return new Font("Tahoma", size, style);
            }
        }

        private void InitializeComponent()
        {
            this.Text = "⚙️ تنظیمات برنامه";
            this.Size = new Size(1000, 720);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.RightToLeft = RightToLeft.Yes;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = BackgroundColor;
            this.Padding = new Padding(15);

            // ========== هدر ==========
            Panel headerPanel = CreateHeaderPanel();
            this.Controls.Add(headerPanel);

            // ========== Content Area (سمت چپ) ==========
            Panel contentArea = new Panel
            {
                Location = new Point(15, 95),
                Size = new Size(720, 490),
                BackColor = BackgroundColor
            };
            this.Controls.Add(contentArea);

            // ساخت محتواها
            pnlDatabaseContent = CreateDatabaseContent();
            pnlPhotosContent = CreatePhotosContent();
            pnlFontContent = CreateFontContent();

            // اضافه کردن به contentArea بدون Dock
            contentArea.Controls.Add(pnlDatabaseContent);
            contentArea.Controls.Add(pnlPhotosContent);
            contentArea.Controls.Add(pnlFontContent);

            // همه بخش‌ها را نمایان نگه‌دار (برای جلوگیری از مشکلات رندر هنگام نمایش مجدد)
            pnlDatabaseContent.Visible = true;
            pnlPhotosContent.Visible = true;
            pnlFontContent.Visible = true;

            // ========== Sidebar (منوی سمت راست) ==========
            Panel sidebarPanel = CreateSidebar();
            this.Controls.Add(sidebarPanel);

            // ========== دکمه‌های پایین ==========
            Panel buttonPanel = CreateButtonPanel();
            this.Controls.Add(buttonPanel);
        }

        private Panel CreateHeaderPanel()
        {
            Panel panel = new Panel
            {
                Location = new Point(15, 15),
                Size = new Size(950, 65),
                BackColor = PrimaryColor
            };
            ApplyRoundedCorners(panel, 12);

            Label lblTitle = new Label
            {
                Text = "⚙️ تنظیمات برنامه",
                Font = GetSafeFont(FontSettings.TitleFont?.FontFamily.Name ?? "Tahoma", 18, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(20, 12),
                Size = new Size(400, 35),
                TextAlign = ContentAlignment.MiddleRight
            };
            panel.Controls.Add(lblTitle);

            Label lblSubtitle = new Label
            {
                Text = "مدیریت تنظیمات مسیرها، فونت‌ها و سایر موارد",
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = Color.FromArgb(230, 240, 255),
                Location = new Point(20, 42),
                Size = new Size(400, 18),
                TextAlign = ContentAlignment.TopRight
            };
            panel.Controls.Add(lblSubtitle);

            return panel;
        }

        private Panel CreateSidebar()
        {
            Panel sidebar = new Panel
            {
                Location = new Point(755, 95),
                Size = new Size(210, 490),
                BackColor = SidebarColor
            };
            ApplyRoundedCorners(sidebar, 10);

            int yPos = 20;

            // عنوان منو
            Label lblMenuTitle = new Label
            {
                Text = "بخش‌ها",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextSecondary,
                Location = new Point(15, yPos),
                Size = new Size(180, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            sidebar.Controls.Add(lblMenuTitle);
            yPos += 40;

            // دکمه دیتابیس
            btnMenuDatabase = CreateMenuButton("💾 تنظیمات دیتابیس", yPos, pnlDatabaseContent);
            sidebar.Controls.Add(btnMenuDatabase);
            yPos += 55;

            // دکمه عکس‌ها
            btnMenuPhotos = CreateMenuButton("🖼️ تنظیمات عکس‌ها", yPos, pnlPhotosContent);
            sidebar.Controls.Add(btnMenuPhotos);
            yPos += 55;

            // دکمه فونت
            btnMenuFont = CreateMenuButton("🔤 تنظیمات فونت", yPos, pnlFontContent);
            sidebar.Controls.Add(btnMenuFont);

            return sidebar;
        }

        private Panel CreateMenuButton(string text, int yPos, Panel targetContent)
        {
            Panel btn = new Panel
            {
                Location = new Point(10, yPos),
                Size = new Size(190, 45),
                BackColor = Color.Transparent,
                Cursor = Cursors.Hand,
                Tag = "menu"
            };
            ApplyRoundedCorners(btn, 8);

            Label lbl = new Label
            {
                Text = text,
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10),
                ForeColor = TextPrimary,
                Location = new Point(10, 0),
                Size = new Size(170, 45),
                TextAlign = ContentAlignment.MiddleRight,
                Cursor = Cursors.Hand
            };
            btn.Controls.Add(lbl);

            // رویداد Click برای Panel
            EventHandler clickHandler = (s, e) => ShowContent(targetContent, btn);
            btn.Click += clickHandler;

            // رویدادهای Hover
            btn.MouseEnter += (s, e) => {
                if (selectedMenuButton != btn)
                {
                    btn.BackColor = MenuHover;
                }
            };
            btn.MouseLeave += (s, e) => {
                if (selectedMenuButton != btn)
                {
                    btn.BackColor = Color.Transparent;
                }
            };

            lbl.MouseEnter += (s, e) => btn.BackColor = selectedMenuButton == btn ? MenuSelected : MenuHover;
            lbl.MouseLeave += (s, e) => btn.BackColor = selectedMenuButton == btn ? MenuSelected : Color.Transparent;

            // وقتی روی label کلیک می‌شه، همون handler رو صدا می‌زنیم
            lbl.Click += clickHandler;

            return btn;
        }

        private void ShowContent(Panel contentPanel, Panel menuButton)
        {
            // مخفی کردن همه محتواها
            if (pnlDatabaseContent != null)
            {
                pnlDatabaseContent.Visible = false;
            }
            if (pnlPhotosContent != null)
            {
                pnlPhotosContent.Visible = false;
            }
            if (pnlFontContent != null)
            {
                pnlFontContent.Visible = false;
            }

            // نمایش محتوای انتخابی
            if (contentPanel != null)
            {
                contentPanel.Visible = true;
                contentPanel.BringToFront();
                contentPanel.Invalidate(true);
                this.Refresh();
                pnlCurrentContent = contentPanel;
            }

            // برداشتن هایلایت از همه دکمه‌ها
            if (selectedMenuButton != null)
            {
                selectedMenuButton.BackColor = Color.Transparent;
                foreach (Control c in selectedMenuButton.Controls)
                {
                    if (c is Label lbl)
                        lbl.ForeColor = TextPrimary;
                }
            }

            // هایلایت دکمه انتخابی
            if (menuButton != null)
            {
                menuButton.BackColor = MenuSelected;
                foreach (Control c in menuButton.Controls)
                {
                    if (c is Label lbl)
                        lbl.ForeColor = Color.White;
                }
                selectedMenuButton = menuButton;
            }
        }

        private Panel CreateDatabaseContent()
        {
            Panel content = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(720, 490),
                BackColor = Color.Transparent,
                AutoScroll = false
            };

            Panel card = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(690, 180),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(card, 10);
            ApplyCardShadow(card);

            // عنوان
            Label lblTitle = new Label
            {
                Text = "💾 تنظیمات دیتابیس",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 14, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(480, 20),
                Size = new Size(190, 35),
                TextAlign = ContentAlignment.MiddleRight
            };
            card.Controls.Add(lblTitle);

            Label lblDesc = new Label
            {
                Text = "مسیر فایل دیتابیس Access را انتخاب کنید",
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                Location = new Point(480, 50),
                Size = new Size(190, 20),
                TextAlign = ContentAlignment.TopRight
            };
            card.Controls.Add(lblDesc);

            // لیبل
            Label lblPath = new Label
            {
                Text = "مسیر فایل:",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(600, 90),
                Size = new Size(70, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            card.Controls.Add(lblPath);

            txtDatabasePath = new TextBox
            {
                Location = new Point(122, 92),
                Size = new Size(470, 28),
                Font = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 9),
                ReadOnly = true,
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };
            card.Controls.Add(txtDatabasePath);

            Button btnBrowse = CreateModernButton("🔍 جستجو", PrimaryColor, 100, 28);
            btnBrowse.Location = new Point(15, 92);
            btnBrowse.Click += BtnBrowseDatabase_Click;
            card.Controls.Add(btnBrowse);

            lblCurrentDatabase = new Label
            {
                Location = new Point(122, 125),
                Size = new Size(470, 20),
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 7.5f),
                ForeColor = TextSecondary,
                Text = ""
            };
            card.Controls.Add(lblCurrentDatabase);

            content.Controls.Add(card);
            return content;
        }

        private Panel CreatePhotosContent()
        {
            Panel content = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(720, 490),
                BackColor = Color.Transparent,
                AutoScroll = false
            };

            Panel card = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(690, 180),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(card, 10);
            ApplyCardShadow(card);

            Label lblTitle = new Label
            {
                Text = "🖼️ تنظیمات عکس‌ها",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 14, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(480, 20),
                Size = new Size(190, 35),
                TextAlign = ContentAlignment.MiddleRight
            };
            card.Controls.Add(lblTitle);

            Label lblDesc = new Label
            {
                Text = "پوشه ذخیره عکس پرسنل را مشخص کنید",
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                Location = new Point(480, 50),
                Size = new Size(190, 20),
                TextAlign = ContentAlignment.TopRight
            };
            card.Controls.Add(lblDesc);

            Label lblPath = new Label
            {
                Text = "مسیر پوشه:",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(600, 90),
                Size = new Size(70, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            card.Controls.Add(lblPath);

            txtPhotosFolder = new TextBox
            {
                Location = new Point(122, 92),
                Size = new Size(470, 28),
                Font = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 9),
                ReadOnly = true,
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };
            card.Controls.Add(txtPhotosFolder);

            Button btnBrowse = CreateModernButton("🔍 جستجو", PrimaryColor, 100, 28);
            btnBrowse.Location = new Point(15, 92);
            btnBrowse.Click += BtnBrowsePhotos_Click;
            card.Controls.Add(btnBrowse);

            lblCurrentPhotos = new Label
            {
                Location = new Point(122, 125),
                Size = new Size(470, 20),
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 7.5f),
                ForeColor = TextSecondary,
                Text = ""
            };
            card.Controls.Add(lblCurrentPhotos);

            content.Controls.Add(card);
            return content;
        }

        private Panel CreateFontContent()
        {
            Panel content = new Panel
            {
                Location = new Point(0, 0),
                Size = new Size(720, 490),
                BackColor = Color.Transparent,
                AutoScroll = false
            };

            Panel card = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(690, 400),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(card, 10);
            ApplyCardShadow(card);

            Label lblTitle = new Label
            {
                Text = "🔤 تنظیمات فونت",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 14, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(480, 20),
                Size = new Size(190, 35),
                TextAlign = ContentAlignment.MiddleRight
            };
            card.Controls.Add(lblTitle);

            Label lblDesc = new Label
            {
                Text = "نوع و اندازه فونت‌های برنامه را تنظیم کنید",
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                Location = new Point(430, 50),
                Size = new Size(240, 20),
                TextAlign = ContentAlignment.TopRight
            };
            card.Controls.Add(lblDesc);

            int yPos = 85;

            // نوع فونت
            Label lblFontFamily = new Label
            {
                Text = "نوع فونت:",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(600, yPos),
                Size = new Size(70, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            card.Controls.Add(lblFontFamily);

            cmbFontFamily = new ComboBox
            {
                Location = new Point(350, yPos),
                Size = new Size(240, 28),
                Font = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 9),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat
            };
            cmbFontFamily.Items.AddRange(new string[] {
                "Tahoma", "Arial", "Segoe UI", "Calibri", "Times New Roman",
                "B Nazanin", "B Mitra", "B Lotus", "B Titr", "IRANSans", "Vazir"
            });
            card.Controls.Add(cmbFontFamily);
            yPos += 50;

            // خط جداکننده
            Panel divider = new Panel
            {
                Location = new Point(30, yPos),
                Size = new Size(630, 1),
                BackColor = Color.FromArgb(230, 230, 230)
            };
            card.Controls.Add(divider);
            yPos += 20;

            // عنوان اندازه‌ها
            Label lblSizesTitle = new Label
            {
                Text = "اندازه فونت‌ها:",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(555, yPos),
                Size = new Size(115, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            card.Controls.Add(lblSizesTitle);
            yPos += 35;

            // Grid فونت‌ها - 3 ستونی
            int col1X = 460;
            int col2X = 250;
            int col3X = 40;
            int labelW = 90;
            int numW = 60;
            int checkW = 60;

            // ردیف 1
            AddFontSizeRowCompact(card, "سرتیتر:", col1X, yPos, labelW, out numTitleSize, out chkBoldTitle, numW, checkW, 16);
            AddFontSizeRowCompact(card, "برچسب:", col2X, yPos, labelW, out numLabelSize, out chkBoldLabel, numW, checkW, 12);
            AddFontSizeRowCompact(card, "دکمه:", col3X, yPos, labelW, out numButtonSize, out chkBoldButton, numW, checkW, 12);
            yPos += 40;

            // ردیف 2
            AddFontSizeRowCompactNoCheckbox(card, "متن:", col1X, yPos, labelW, out numTextBoxSize, numW, checkW, 11);
            AddFontSizeRowCompactNoCheckbox(card, "متن عادی:", col2X, yPos, labelW, out numBodySize, numW, checkW, 10);

            content.Controls.Add(card);
            return content;
        }

        private void AddFontSizeRowCompact(Panel parent, string label, int x, int y, int labelW,
            out NumericUpDown numeric, out CheckBox checkbox, int numW, int checkW, int defaultValue)
        {
            Label lbl = new Label
            {
                Text = label,
                Location = new Point(x + numW + checkW + 5, y + 2),
                Size = new Size(labelW, 22),
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.MiddleRight
            };
            parent.Controls.Add(lbl);

            numeric = new NumericUpDown
            {
                Location = new Point(x + checkW + 3, y),
                Size = new Size(numW, 26),
                Minimum = 8,
                Maximum = 72,
                Value = defaultValue,
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                BorderStyle = BorderStyle.FixedSingle
            };
            parent.Controls.Add(numeric);

            checkbox = new CheckBox
            {
                Text = "ضخیم",
                Location = new Point(x, y + 2),
                Size = new Size(checkW, 22),
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8),
                ForeColor = TextSecondary
            };
            parent.Controls.Add(checkbox);
        }

        private void AddFontSizeRowCompactNoCheckbox(Panel parent, string label, int x, int y, int labelW,
            out NumericUpDown numeric, int numW, int checkW, int defaultValue)
        {
            Label lbl = new Label
            {
                Text = label,
                Location = new Point(x + numW + checkW + 5, y + 2),
                Size = new Size(labelW, 22),
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.MiddleRight
            };
            parent.Controls.Add(lbl);

            numeric = new NumericUpDown
            {
                Location = new Point(x + checkW + 3, y),
                Size = new Size(numW, 26),
                Minimum = 8,
                Maximum = 72,
                Value = defaultValue,
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                BorderStyle = BorderStyle.FixedSingle
            };
            parent.Controls.Add(numeric);
        }

        private Panel CreateButtonPanel()
        {
            Panel panel = new Panel
            {
                Location = new Point(15, 595),
                Size = new Size(950, 60),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(panel, 10);
            ApplyCardShadow(panel);

            int centerX = panel.Width / 2;
            int buttonWidth = 130;
            int buttonHeight = 38;
            int spacing = 12;

            // دکمه ذخیره (وسط)
            Button btnSave = CreateModernButton("💾 ذخیره", AccentColor, buttonWidth, buttonHeight);
            btnSave.Location = new Point(centerX - buttonWidth / 2, 11);
            btnSave.Font = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            btnSave.Click += BtnSave_Click;
            panel.Controls.Add(btnSave);

            // دکمه بازنشانی (راست)
            Button btnReset = CreateModernButton("🔄 بازنشانی", WarningColor, buttonWidth, buttonHeight);
            btnReset.Location = new Point(centerX + buttonWidth / 2 + spacing, 11);
            btnReset.Font = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            btnReset.Click += BtnReset_Click;
            panel.Controls.Add(btnReset);

            // دکمه لغو (چپ)
            Button btnCancel = CreateModernButton("❌ لغو", DangerColor, buttonWidth, buttonHeight);
            btnCancel.Location = new Point(centerX - buttonWidth / 2 - buttonWidth - spacing, 11);
            btnCancel.Font = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            btnCancel.Click += (s, e) => this.Close();
            panel.Controls.Add(btnCancel);

            return panel;
        }

        private Button CreateModernButton(string text, Color backColor, int width, int height)
        {
            Button btn = new Button
            {
                Text = text,
                Size = new Size(width, height),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10)
            };
            btn.FlatAppearance.BorderSize = 0;
            ApplyRoundedCorners(btn, 8);

            Color originalColor = backColor;
            btn.MouseEnter += (s, e) => btn.BackColor = ControlPaint.Light(originalColor, 0.1f);
            btn.MouseLeave += (s, e) => btn.BackColor = originalColor;

            return btn;
        }

        private void LoadCurrentSettings()
        {
            txtDatabasePath.Text = AppSettings.DatabasePath;
            txtPhotosFolder.Text = AppSettings.PhotosFolder;
            lblCurrentDatabase.Text = $"📂 {AppSettings.DatabasePath}";
            lblCurrentPhotos.Text = $"📂 {AppSettings.PhotosFolder}";

            cmbFontFamily.Text = FontSettings.FontFamilyName;
            numTitleSize.Value = (decimal)FontSettings.TitleFontSize;
            numLabelSize.Value = (decimal)FontSettings.LabelFontSize;
            numTextBoxSize.Value = (decimal)FontSettings.TextBoxFontSize;
            numButtonSize.Value = (decimal)FontSettings.ButtonFontSize;
            numBodySize.Value = (decimal)FontSettings.BodyFontSize;
            chkBoldTitle.Checked = FontSettings.TitleFontBold;
            chkBoldLabel.Checked = FontSettings.LabelFontBold;
            chkBoldButton.Checked = FontSettings.ButtonFontBold;
        }

        private void BtnBrowseDatabase_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Access Database (*.accdb)|*.accdb|All Files (*.*)|*.*";
                ofd.Title = "انتخاب فایل دیتابیس";
                ofd.InitialDirectory = Path.GetDirectoryName(AppSettings.DatabasePath);

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtDatabasePath.Text = ofd.FileName;
                }
            }
        }

        private void BtnBrowsePhotos_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "انتخاب پوشه عکس‌ها";
                fbd.SelectedPath = AppSettings.PhotosFolder;
                fbd.ShowNewFolderButton = true;

                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtPhotosFolder.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (!File.Exists(txtDatabasePath.Text))
                {
                    DialogResult result = MessageBox.Show(
                        "⚠️ فایل دیتابیس در مسیر انتخابی وجود ندارد.\n\nآیا می‌خواهید ادامه دهید؟",
                        "هشدار",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );

                    if (result == DialogResult.No)
                        return;
                }

                if (!Directory.Exists(txtPhotosFolder.Text))
                {
                    DialogResult result = MessageBox.Show(
                        "📁 پوشه عکس‌ها وجود ندارد.\n\nآیا می‌خواهید آن را ایجاد کنید؟",
                        "پرسش",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );

                    if (result == DialogResult.Yes)
                    {
                        Directory.CreateDirectory(txtPhotosFolder.Text);
                    }
                    else
                    {
                        return;
                    }
                }

                AppSettings.DatabasePath = txtDatabasePath.Text;
                AppSettings.PhotosFolder = txtPhotosFolder.Text;

                FontSettings.FontFamilyName = cmbFontFamily.Text;
                FontSettings.TitleFontSize = (float)numTitleSize.Value;
                FontSettings.LabelFontSize = (float)numLabelSize.Value;
                FontSettings.TextBoxFontSize = (float)numTextBoxSize.Value;
                FontSettings.ButtonFontSize = (float)numButtonSize.Value;
                FontSettings.BodyFontSize = (float)numBodySize.Value;
                FontSettings.TitleFontBold = chkBoldTitle.Checked;
                FontSettings.LabelFontBold = chkBoldLabel.Checked;
                FontSettings.ButtonFontBold = chkBoldButton.Checked;

                FontSettings.SaveSettings();

                MessageBox.Show(
                    "✅ تنظیمات با موفقیت ذخیره شد!\n\n🔄 لطفاً برنامه را مجدداٌ راه‌اندازی کنید.",
                    "موفقیت",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"❌ خطا در ذخیره تنظیمات:\n\n{ex.Message}",
                    "خطا",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "⚠️ آیا مطمئن هستید که می‌خواهید تنظیمات را به حالت پیش‌فرض برگردانید؟\n\nتمامی تغییرات از بین خواهد رفت!",
                "تایید بازنشانی",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result == DialogResult.Yes)
            {
                AppSettings.ResetToDefaults();
                FontSettings.ResetToDefaults();
                LoadCurrentSettings();

                MessageBox.Show(
                    "✅ تنظیمات به حالت پیش‌فرض برگشت!\n\n🔄 برای اعمال تغییرات، برنامه را مجدداً راه‌اندازی کنید.",
                    "موفقیت",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
        }

        private void ApplyRoundedCorners(Control control, int radius)
        {
            try
            {
                GraphicsPath path = new GraphicsPath();
                path.AddArc(0, 0, radius, radius, 180, 90);
                path.AddArc(control.Width - radius, 0, radius, radius, 270, 90);
                path.AddArc(control.Width - radius, control.Height - radius, radius, radius, 0, 90);
                path.AddArc(0, control.Height - radius, radius, radius, 90, 90);
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
                {
                    e.Graphics.FillRectangle(shadowBrush, new Rectangle(3, 3, panel.Width - 3, panel.Height - 3));
                }
            };
        }
    }
}