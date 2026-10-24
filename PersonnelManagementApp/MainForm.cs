using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    /// <summary>
    /// داشبورد اصلی — طراحی مدرن سازمانی (RTL، گریادیانت، آیکون‌های خطی، سایه نرم).
    /// </summary>
    public partial class MainForm : Form
    {
        // ── پالت رنگی ──
        private static readonly Color CanvasColor   = Color.FromArgb(244, 246, 250);
        private static readonly Color CardColor     = Color.White;
        private static readonly Color BorderColor   = Color.FromArgb(226, 232, 240);
        private static readonly Color TextColor     = Color.FromArgb(15, 30, 55);
        private static readonly Color MutedColor    = Color.FromArgb(110, 122, 140);
        private static readonly Color HeaderTop     = Color.FromArgb(11, 30, 63);
        private static readonly Color HeaderBottom  = Color.FromArgb(30, 99, 199);
        private static readonly Color Accent        = Color.FromArgb(56, 189, 248);
        private static readonly Color PrimaryColor  = Color.FromArgb(0, 122, 204);
        private static readonly Color SuccessColor  = Color.FromArgb(16, 185, 129);
        private static readonly Color WarningColor  = Color.FromArgb(245, 158, 11);
        private static readonly Color DangerColor   = Color.FromArgb(239, 68, 68);
        private static readonly Color VioletColor   = Color.FromArgb(139, 92, 246);

        private Label lblKpiPersonnel = null!;
        private Label lblKpiPosts = null!;
        private Label lblKpiProvinces = null!;
        private Label lblKpiStatus = null!;
        private Label lblFooterStatus = null!;
        private Label lblFooterPath = null!;

        public MainForm()
        {
            BuildUi();
            Shown += async (_, _) => await RefreshDatabaseStatusAsync();
        }

        // ─────────────────────────────────────────────
        // فونت وزیرمتن (با جایگزین امن)
        // ─────────────────────────────────────────────
        private static string ResolveFontFamily()
        {
            try
            {
                using var installed = new System.Drawing.Text.InstalledFontCollection();
                foreach (var family in installed.Families)
                {
                    if (family.Name.Contains("Vazir", StringComparison.OrdinalIgnoreCase))
                        return family.Name;
                }
            }
            catch { }

            return string.IsNullOrWhiteSpace(FontSettings.FontFamilyName) ? "Tahoma" : FontSettings.FontFamilyName;
        }

        private static Font UiFont(float size, FontStyle style = FontStyle.Regular)
        {
            try { return new Font(ResolveFontFamily(), size, style); }
            catch { return new Font("Tahoma", size, style); }
        }

        // ─────────────────────────────────────────────
        // ساختار کلی
        // ─────────────────────────────────────────────
        private void BuildUi()
        {
            SuspendLayout();

            Text = "سامانه مدیریت پرسنل و پست‌ها";
            WindowState = FormWindowState.Maximized;
            MinimumSize = new Size(1100, 700);
            StartPosition = FormStartPosition.CenterScreen;
            RightToLeft = RightToLeft.Yes;
            BackColor = CanvasColor;
            Font = UiFont(9F);
            AutoScaleMode = AutoScaleMode.Dpi;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                BackColor = CanvasColor,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 104F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 96F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));
            Controls.Add(root);

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildStatsStrip(), 0, 1);
            root.Controls.Add(BuildNavArea(), 0, 2);
            root.Controls.Add(BuildFooter(), 0, 3);

            ResumeLayout(true);
        }

        // ─────────────────────────────────────────────
        // ۱) هدر
        // ─────────────────────────────────────────────
        private Control BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            header.Paint += (_, e) =>
            {
                using var brush = new LinearGradientBrush(header.ClientRectangle, HeaderTop, HeaderBottom, 90f);
                e.Graphics.FillRectangle(brush, header.ClientRectangle);
                using var glow = new SolidBrush(Color.FromArgb(25, 255, 255, 255));
                e.Graphics.FillEllipse(glow, header.Width - 340, -120, 440, 330);
                e.Graphics.FillEllipse(glow, header.Width - 80, 52, 190, 190);
            };

            var content = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.Transparent,
                Padding = new Padding(28, 12, 28, 12),
                Margin = new Padding(0)
            };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80F));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            var iconBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                Margin = new Padding(0, 6, 12, 6),
                BackColor = Color.Transparent
            };
            TryLoadHeaderIcon(iconBox);

            var titleStack = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };
            titleStack.Controls.Add(new Label
            {
                Text = "مدیریت یکپارچه پرسنل، پست‌های انتقال و گزارش‌های تحلیلی",
                Dock = DockStyle.Fill,
                Font = UiFont(9.5F),
                ForeColor = Color.FromArgb(205, 218, 235),
                TextAlign = ContentAlignment.MiddleRight,
                AutoEllipsis = true
            });
            titleStack.Controls.Add(new Label
            {
                Text = "سامانه مدیریت پرسنل و پست‌ها",
                Dock = DockStyle.Top,
                Height = 42,
                Font = UiFont(19F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleRight,
                AutoEllipsis = true
            });

            var dateStack = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };
            dateStack.Controls.Add(new Label
            {
                Text = "نسخه ۳.۰ • Windows Forms",
                Dock = DockStyle.Fill,
                Font = UiFont(8.5F),
                ForeColor = Color.FromArgb(150, 168, 190),
                TextAlign = ContentAlignment.MiddleLeft
            });
            dateStack.Controls.Add(new Label
            {
                Text = GetPersianDateText(),
                Dock = DockStyle.Top,
                Height = 30,
                Font = UiFont(11F, FontStyle.Bold),
                ForeColor = Accent,
                TextAlign = ContentAlignment.MiddleLeft
            });

            content.Controls.Add(iconBox, 0, 0);
            content.Controls.Add(titleStack, 1, 0);
            content.Controls.Add(dateStack, 2, 0);
            header.Controls.Add(content);

            header.Controls.Add(new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 3,
                BackColor = Accent
            });

            return header;
        }

        private void TryLoadHeaderIcon(PictureBox box)
        {
            try
            {
                string iconPath = Path.Combine(Application.StartupPath, "app_icon.ico");
                if (File.Exists(iconPath))
                {
                    using var icon = new Icon(iconPath);
                    box.Image = icon.ToBitmap();
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("MainForm.TryLoadHeaderIcon", ex);
            }
        }

        private static string GetPersianDateText()
        {
            var calendar = new PersianCalendar();
            DateTime now = DateTime.Now;
            string[] weekDays =
            {
                "یکشنبه", "دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه"
            };
            return $"{weekDays[(int)now.DayOfWeek]} {calendar.GetDayOfMonth(now):00}/{calendar.GetMonth(now):00}/{calendar.GetYear(now)}";
        }

        // ─────────────────────────────────────────────
        // ۲) نوار آمار
        // ─────────────────────────────────────────────
        private Control BuildStatsStrip()
        {
            var strip = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                BackColor = CanvasColor,
                Padding = new Padding(18, 12, 18, 8),
                Margin = new Padding(0)
            };
            for (int i = 0; i < 4; i++)
                strip.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
            strip.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            strip.Controls.Add(CreateStatCard(FlatIconKind.Personnel, "پرسنل ثبت‌شده", "…", PrimaryColor, out lblKpiPersonnel), 0, 0);
            strip.Controls.Add(CreateStatCard(FlatIconKind.Posts, "پست‌های ثبت‌شده", "…", SuccessColor, out lblKpiPosts), 1, 0);
            strip.Controls.Add(CreateStatCard(FlatIconKind.Provinces, "استان‌های تحت پوشش", "…", WarningColor, out lblKpiProvinces), 2, 0);

            var statusCard = CreateStatCard(FlatIconKind.Database, "وضعیت پایگاه داده", "در حال بررسی…", MutedColor, out lblKpiStatus);
            statusCard.Cursor = Cursors.Hand;
            statusCard.Click += (_, _) => BtnSettings_Click(this, EventArgs.Empty);
            strip.Controls.Add(statusCard, 3, 0);

            return strip;
        }

        private Panel CreateStatCard(FlatIconKind icon, string caption, string initialValue, Color accent, out Label valueLabel)
        {
            var card = new RoundedPanel(14)
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Margin = new Padding(5),
                Padding = new Padding(14, 10, 14, 8)
            };
            ApplySoftShadow(card);
            card.Paint += (_, e) =>
            {
                using var top = new SolidBrush(accent);
                e.Graphics.FillRectangle(top, 0, 0, card.Width, 4);
            };

            var iconChip = new RoundedPanel(11)
            {
                Dock = DockStyle.Right,
                Width = 50,
                BackColor = BlendWithWhite(accent, 0.86f),
                Margin = new Padding(12, 4, 0, 4)
            };
            iconChip.Paint += (_, e) =>
            {
                using var pen = new Pen(accent, 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                FlatIcons.Draw(icon, e.Graphics, iconChip.ClientRectangle, pen);
            };

            valueLabel = new Label
            {
                Text = initialValue,
                Dock = DockStyle.Fill,
                Font = UiFont(17F, FontStyle.Bold),
                ForeColor = accent,
                TextAlign = ContentAlignment.MiddleRight,
                AutoEllipsis = true,
                BackColor = Color.Transparent
            };

            var captionLabel = new Label
            {
                Text = caption,
                Dock = DockStyle.Top,
                Height = 22,
                Font = UiFont(9F),
                ForeColor = MutedColor,
                TextAlign = ContentAlignment.MiddleRight,
                BackColor = Color.Transparent
            };

            card.Controls.Add(valueLabel);
            card.Controls.Add(iconChip);
            card.Controls.Add(captionLabel);
            return card;
        }

        // ─────────────────────────────────────────────
        // ۳) منطقه مرکزی
        // ─────────────────────────────────────────────
        private Control BuildNavArea()
        {
            var nav = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CanvasColor,
                Padding = new Padding(0)
            };

            // محتوای پایدار (بدون اسکرول) — در هر اندازه پنجره جا می‌شود
            var content = new Panel
            {
                Dock = DockStyle.Top,
                BackColor = CanvasColor,
                Padding = new Padding(0, 10, 0, 10)
            };

            var heading = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = CanvasColor,
                Margin = new Padding(0)
            };
            heading.Controls.Add(new Label
            {
                Text = "بخش‌های اصلی",
                Dock = DockStyle.Top,
                Height = 24,
                Font = UiFont(14F, FontStyle.Bold),
                ForeColor = TextColor,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 8, 0)
            });
            heading.Controls.Add(new Label
            {
                Text = "برای دسترسی سریع، روی یکی از دکمه‌ها کلیک کنید",
                Dock = DockStyle.Bottom,
                Height = 16,
                Font = UiFont(8F),
                ForeColor = MutedColor,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 8, 0)
            });

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                BackColor = CanvasColor,
                ColumnCount = 2,
                RowCount = 2,
                RightToLeft = RightToLeft.Yes,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 192F));
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 192F));

            grid.Controls.Add(CreateSectionCard(
                FlatIconKind.Personnel, "پرسنل",
                "ثبت، ویرایش، حذف، جستجو و مدیریت اطلاعات کارکنان", PrimaryColor,
                ("ثبت پرسنل", () => OpenForm(() => new FormPersonnelRegister())),
                ("ویرایش پرسنل", () => OpenForm(() => new FormPersonnelEdit())),
                ("حذف پرسنل", () => OpenForm(() => new FormPersonnelDelete())),
                ("جستجوی پیشرفته", () => OpenForm(() => new FormPersonnelSearch())),
                ("مدیریت کامل", () => OpenForm(() => new FormPersonnelManager()))), 0, 0);

            grid.Controls.Add(CreateSectionCard(
                FlatIconKind.Posts, "پست‌ها",
                "ثبت و مدیریت پست‌های سازمانی", SuccessColor,
                ("ثبت پست جدید", () => OpenForm(() => new FormPostsEditor())),
                ("مدیریت پست‌ها", () => OpenForm(() => new FormPostsManager()))), 1, 0);

            grid.Controls.Add(CreateSectionCard(
                FlatIconKind.Chart, "تحلیل داده‌ها",
                "نمودارها، آمار و خروجی گزارش‌های تحلیلی", VioletColor,
                ("تحلیل پرسنل", () => OpenForm(() => new FormPersonnelAnalytics())),
                ("تحلیل پست‌ها", () => OpenForm(() => new FormPostsAnalytics())),
                ("خروجی نمودار پرسنل", () => OpenForm(() => new FormExportCharts())),
                ("خروجی نمودار پست‌ها", () => OpenForm(() => new FormExportPostCharts()))), 0, 1);

            grid.Controls.Add(CreateSectionCard(
                FlatIconKind.Settings, "ابزارها و تنظیمات",
                "تنظیمات سامانه، ورود داده و ابزارهای کمکی", Color.FromArgb(71, 85, 105),
                ("تنظیمات", () => BtnSettings_Click(this, EventArgs.Empty)),
                ("ورود از اکسل", () => OpenForm(() => new FormImportExcel())),
                ("بررسی عکس‌های ناقص", () => OpenForm(() => new FormMissingPhotos())),
                ("پوشه پایگاه داده", OpenDatabaseFolder),
                ("درباره برنامه", () => ShowAboutDialog())), 1, 1);

            content.Controls.Add(grid);
            content.Controls.Add(heading);
            nav.Controls.Add(content);

            void Recenter()
            {
                int available = Math.Max(0, nav.ClientSize.Width - 48);
                int width = Math.Max(720, Math.Min(1240, available));
                content.Width = width;
                int height = 10 + heading.Height + 6 + grid.Height + 10;
                content.Height = height;
                content.Margin = new Padding(Math.Max(0, (nav.ClientSize.Width - width) / 2), 0, 0, 0);
                grid.Height = 2 * 192;
            }
            nav.Resize += (_, _) => Recenter();
            Recenter();

            return nav;
        }

        private Panel CreateSectionCard(
            FlatIconKind icon, string title, string description, Color accent,
            params (string Text, Action Click)[] actions)
        {
            var card = new RoundedPanel(16)
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Margin = new Padding(6),
                Padding = new Padding(16, 12, 16, 10)
            };
            ApplySoftShadow(card);
            card.Paint += (_, e) =>
            {
                using var pen = new Pen(BorderColor);
                e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
                using var top = new SolidBrush(accent);
                e.Graphics.FillRectangle(top, 0, 0, card.Width, 4);
            };

            // ناحیه سربرگ کارت
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 62,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };

            var iconChip = new RoundedPanel(11)
            {
                Dock = DockStyle.Right,
                Width = 52,
                BackColor = BlendWithWhite(accent, 0.88f),
                Margin = new Padding(12, 4, 0, 4)
            };
            iconChip.Paint += (_, e) =>
            {
                using var pen = new Pen(accent, 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                FlatIcons.Draw(icon, e.Graphics, iconChip.ClientRectangle, pen);
            };

            header.Controls.Add(new Label
            {
                Text = description,
                Dock = DockStyle.Fill,
                Font = UiFont(8.5F),
                ForeColor = MutedColor,
                TextAlign = ContentAlignment.TopRight,
                Padding = new Padding(0, 2, 0, 0),
                BackColor = Color.Transparent
            });
            header.Controls.Add(new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                Height = 28,
                Font = UiFont(13F, FontStyle.Bold),
                ForeColor = TextColor,
                TextAlign = ContentAlignment.MiddleRight,
                BackColor = Color.Transparent
            });
            header.Controls.Add(iconChip);

            // جداکننده زیر سربرگ
            var divider = new Panel
            {
                Dock = DockStyle.Top,
                Height = 1,
                BackColor = BorderColor,
                Margin = new Padding(0, 4, 0, 6)
            };

            // دکمه‌های عملیات بخش
            var actionsFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                WrapContents = true,
                FlowDirection = FlowDirection.RightToLeft,
                BackColor = Color.Transparent,
                Padding = new Padding(0, 8, 0, 0),
                Margin = new Padding(0)
            };
            foreach (var (text, click) in actions)
                actionsFlow.Controls.Add(CreateActionButton(text, accent, click));

            card.Controls.Add(actionsFlow);
            card.Controls.Add(divider);
            card.Controls.Add(header);

            return card;
        }

        private Button CreateActionButton(string text, Color accent, Action onClick)
        {
            var font = UiFont(9F, FontStyle.Bold);
            int textWidth = TextRenderer.MeasureText(text, font).Width;

            var button = new Button
            {
                Text = text,
                Width = Math.Max(104, textWidth + 34),
                Height = 34,
                Margin = new Padding(0, 0, 8, 8),
                FlatStyle = FlatStyle.Flat,
                BackColor = BlendWithWhite(accent, 0.93f),
                ForeColor = accent,
                Font = font,
                Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter,
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderColor = BlendWithWhite(accent, 0.72f);
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.MouseOverBackColor = BlendWithWhite(accent, 0.85f);
            button.FlatAppearance.MouseDownBackColor = BlendWithWhite(accent, 0.78f);
            button.Click += (_, _) => onClick();
            return button;
        }

        // ─────────────────────────────────────────────
        // ۴) فوتر
        // ─────────────────────────────────────────────
        private Control BuildFooter()
        {
            var footer = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Padding = new Padding(20, 0, 20, 0),
                Margin = new Padding(0)
            };
            footer.Paint += (_, e) =>
            {
                using var pen = new Pen(BorderColor);
                e.Graphics.DrawLine(pen, 0, 0, footer.Width, 0);
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = CardColor,
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 380F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260F));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            lblFooterStatus = new Label
            {
                Text = "○ در حال بررسی پایگاه داده…",
                Dock = DockStyle.Fill,
                Font = UiFont(9F),
                ForeColor = MutedColor,
                TextAlign = ContentAlignment.MiddleRight,
                AutoEllipsis = true
            };

            lblFooterPath = new Label
            {
                Text = AppSettings.DatabasePath,
                Dock = DockStyle.Fill,
                Font = UiFont(8.5F),
                ForeColor = MutedColor,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };

            // دکمه‌های کوچک سمت چپ: درباره و خروج
            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = CardColor,
                Padding = new Padding(0)
            };

            var btnAbout = CreateFooterButton("درباره", Color.FromArgb(99, 102, 241));
            btnAbout.Click += (_, _) => ShowAboutDialog();

            var btnExit = CreateFooterButton("خروج", DangerColor);
            btnExit.Click += (_, _) => Close();

            actions.Controls.Add(btnAbout);
            actions.Controls.Add(btnExit);

            layout.Controls.Add(lblFooterStatus, 0, 0);
            layout.Controls.Add(lblFooterPath, 1, 0);
            layout.Controls.Add(actions, 2, 0);
            footer.Controls.Add(layout);
            return footer;
        }

        private static Button CreateFooterButton(string text, Color accent)
        {
            var button = new Button
            {
                Text = text,
                FlatStyle = FlatStyle.Flat,
                BackColor = BlendWithWhite(accent, 0.90f),
                ForeColor = accent,
                Font = UiFont(8.5F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Width = 112,
                Height = 30,
                Margin = new Padding(0, 0, 8, 0),
                TextAlign = ContentAlignment.MiddleCenter,
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = BlendWithWhite(accent, 0.70f);
            button.FlatAppearance.MouseOverBackColor = BlendWithWhite(accent, 0.80f);
            button.FlatAppearance.MouseDownBackColor = BlendWithWhite(accent, 0.72f);
            return button;
        }

        // ─────────────────────────────────────────────
        // عملیات
        // ─────────────────────────────────────────────
        private void OpenForm(Func<Form> factory)
        {
            Cursor previous = Cursor;
            Cursor = Cursors.WaitCursor;
            try
            {
                using Form form = factory();
                form.ShowDialog(this);
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("MainForm.OpenForm", ex);
                MessageBox.Show(this,
                    $"خطا در باز کردن فرم:\n\n{ex.Message}",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = previous;
            }
        }

        private void OpenDatabaseFolder()
        {
            try
            {
                string path = AppSettings.DatabasePath;
                if (!File.Exists(path))
                {
                    MessageBox.Show(this,
                        $"فایل پایگاه داده پیدا نشد:\n{path}\n\nاز بخش تنظیمات مسیر فایل را بررسی کنید.",
                        "پایگاه داده", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{path}\"",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("MainForm.OpenDatabaseFolder", ex);
                MessageBox.Show(this, $"خطا در باز کردن پوشه:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowAboutDialog()
        {
            string provider = DatabaseProvider.ProviderName ?? "نصب نشده";
            MessageBox.Show(this,
                "سامانه مدیریت پرسنل و پست‌ها\n" +
                "نسخه ۳.۰\n\n" +
                $"موتور پایگاه داده: {provider}\n" +
                $"معماری برنامه: {DatabaseProvider.Bitness}\n" +
                $"نسخه .NET: {Environment.Version}\n" +
                $"مسیر پایگاه داده:\n{AppSettings.DatabasePath}\n\n" +
                "ساخته‌شده با Windows Forms و Access OLE DB",
                "درباره برنامه", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnSettings_Click(object? sender, EventArgs e)
        {
            using var settingsForm = new FormSettings();
            if (settingsForm.ShowDialog(this) == DialogResult.OK)
            {
                _ = RefreshDatabaseStatusAsync();
            }
        }

        private async Task RefreshDatabaseStatusAsync()
        {
            SetStatus("○ در حال بررسی پایگاه داده…", MutedColor);
            lblKpiStatus.ForeColor = MutedColor;
            lblKpiStatus.Text = "در حال بررسی…";

            DashboardStats stats = await Task.Run(LoadDashboardStats);
            if (IsDisposed)
                return;

            if (stats.Connected)
            {
                lblKpiPersonnel.Text = stats.PersonnelCount.ToString("N0");
                lblKpiPosts.Text = stats.PostsCount.ToString("N0");
                lblKpiProvinces.Text = stats.ProvincesCount.ToString("N0");
                lblKpiStatus.Text = "متصل";
                lblKpiStatus.ForeColor = SuccessColor;
                SetStatus($"● پایگاه داده متصل است • {stats.Provider} • {stats.Bitness}", SuccessColor);
            }
            else
            {
                lblKpiPersonnel.Text = "—";
                lblKpiPosts.Text = "—";
                lblKpiProvinces.Text = "—";
                lblKpiStatus.Text = "قطع";
                lblKpiStatus.ForeColor = DangerColor;
                SetStatus($"● اتصال برقرار نیست: {stats.Error}", DangerColor);
            }

            lblFooterPath.Text = AppSettings.DatabasePath;
        }

        private void SetStatus(string text, Color color)
        {
            if (lblFooterStatus == null)
                return;

            lblFooterStatus.Text = text;
            lblFooterStatus.ForeColor = color;
            lblFooterStatus.Refresh();
        }

        private static DashboardStats LoadDashboardStats()
        {
            try
            {
                var db = new DbHelper();
                if (!db.TestConnectionSilent())
                {
                    return DashboardStats.Failed(
                        "اتصال به پایگاه داده برقرار نشد. پیش‌نیاز Access را بررسی کنید.");
                }

                object? personnel = db.ExecuteScalar("SELECT COUNT(*) FROM Personnel", null, showError: false);
                object? posts = db.ExecuteScalar("SELECT COUNT(*) FROM Posts", null, showError: false);
                object? provinces = db.ExecuteScalar("SELECT COUNT(*) FROM Provinces", null, showError: false);

                return new DashboardStats(
                    connected: true,
                    personnel: Convert.ToInt32(personnel ?? 0),
                    posts: Convert.ToInt32(posts ?? 0),
                    provinces: Convert.ToInt32(provinces ?? 0),
                    provider: DatabaseProvider.ProviderName ?? "—",
                    bitness: DatabaseProvider.Bitness);
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("MainForm.LoadDashboardStats", ex);
                return DashboardStats.Failed(ex.Message);
            }
        }

        // ─────────────────────────────────────────────
        // ابزارهای ظاهری
        // ─────────────────────────────────────────────
        private static Color BlendWithWhite(Color color, float amount)
        {
            int r = (int)(color.R + (255 - color.R) * amount);
            int g = (int)(color.G + (255 - color.G) * amount);
            int b = (int)(color.B + (255 - color.B) * amount);
            return Color.FromArgb(r, g, b);
        }

        private void ApplyRoundedCorners(Control control, int radius)
        {
            try
            {
                var path = new GraphicsPath();
                path.AddArc(0, 0, radius, radius, 180, 90);
                path.AddArc(control.Width - radius, 0, radius, radius, 270, 90);
                path.AddArc(control.Width - radius, control.Height - radius, radius, radius, 0, 90);
                path.AddArc(0, control.Height - radius, radius, radius, 90, 90);
                path.CloseFigure();
                control.Region = new Region(path);
            }
            catch { }
        }

        private void ApplySoftShadow(Panel panel)
        {
            panel.Paint += (_, e) =>
            {
                using var pen = new Pen(Color.FromArgb(12, 15, 23, 40), 1f);
                e.Graphics.DrawRectangle(pen, 1, 1, panel.Width - 3, panel.Height - 3);
            };
        }

        private static GraphicsPath RoundedPath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius, radius, 180, 90);
            path.AddArc(rect.Right - radius, rect.Y, radius, radius, 270, 90);
            path.AddArc(rect.Right - radius, rect.Bottom - radius, radius, radius, 0, 90);
            path.AddArc(rect.X, rect.Bottom - radius, radius, radius, 90, 90);
            path.CloseFigure();
            return path;
        }

        private sealed class DashboardStats
        {
            public DashboardStats(bool connected, int personnel, int posts, int provinces, string provider, string bitness)
            {
                Connected = connected;
                PersonnelCount = personnel;
                PostsCount = posts;
                ProvincesCount = provinces;
                Provider = provider;
                Bitness = bitness;
            }

            public bool Connected { get; }
            public int PersonnelCount { get; }
            public int PostsCount { get; }
            public int ProvincesCount { get; }
            public string Provider { get; }
            public string Bitness { get; }
            public string Error { get; set; } = string.Empty;

            public static DashboardStats Failed(string error)
            {
                var stats = new DashboardStats(false, 0, 0, 0, "—", "—");
                stats.Error = error;
                return stats;
            }
        }
    }

    /// <summary>
    /// پنل با گوشه‌های گرد که Region خود را پس از هر تغییر اندازه به‌روزرسانی می‌کند.
    /// </summary>
    internal sealed class RoundedPanel : Panel
    {
        private readonly int _radius;

        public RoundedPanel(int radius)
        {
            _radius = radius;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            try
            {
                if (_radius > 0 && Width > 0 && Height > 0)
                {
                    using var path = new GraphicsPath();
                    int r = _radius;
                    path.AddArc(0, 0, r, r, 180, 90);
                    path.AddArc(Width - r, 0, r, r, 270, 90);
                    path.AddArc(Width - r, Height - r, r, r, 0, 90);
                    path.AddArc(0, Height - r, r, r, 90, 90);
                    path.CloseFigure();
                    Region = new Region(path);
                }
                else
                {
                    Region = null;
                }
            }
            catch { }
        }
    }

    /// <summary>
    /// آیکون‌های خطی مینیمال (رسم با GDI+ بدون ایموجی).
    /// </summary>
    public enum FlatIconKind
    {
        Personnel,
        Posts,
        Provinces,
        Database,
        Chart,
        Settings,
        Folder,
        Info,
        Exit
    }

    public static class FlatIcons
    {
        public static void Draw(FlatIconKind kind, Graphics g, Rectangle bounds, Pen pen)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;

            switch (kind)
            {
                case FlatIconKind.Personnel: DrawPersonnel(g, bounds, pen); break;
                case FlatIconKind.Posts: DrawPosts(g, bounds, pen); break;
                case FlatIconKind.Provinces: DrawProvinces(g, bounds, pen); break;
                case FlatIconKind.Database: DrawDatabase(g, bounds, pen); break;
                case FlatIconKind.Chart: DrawChart(g, bounds, pen); break;
                case FlatIconKind.Settings: DrawSettings(g, bounds, pen); break;
                case FlatIconKind.Folder: DrawFolder(g, bounds, pen); break;
                case FlatIconKind.Info: DrawInfo(g, bounds, pen); break;
                case FlatIconKind.Exit: DrawExit(g, bounds, pen); break;
            }
        }

        private static void DrawPersonnel(Graphics g, Rectangle b, Pen pen)
        {
            int c = b.Width / 2;
            int cy = b.Height / 2;
            g.DrawEllipse(pen, c - 13, cy - 18, 13, 12);
            g.DrawArc(pen, c - 20, cy - 4, 18, 18, 0, 180);
            g.DrawEllipse(pen, c + 2, cy - 24, 13, 12);
            g.DrawArc(pen, c - 5, cy - 10, 18, 18, 0, 180);
        }

        private static void DrawPosts(Graphics g, Rectangle b, Pen pen)
        {
            int midX = b.Width / 2;
            int topY = 4;
            int bottomY = b.Height - 4;
            g.DrawLine(pen, midX, topY, midX, bottomY);
            g.DrawLine(pen, midX, bottomY, midX - 14, bottomY);
            g.DrawLine(pen, midX, bottomY, midX + 14, bottomY);
            g.DrawLine(pen, midX - 12, topY + 12, midX + 12, topY + 12);
            g.DrawLine(pen, midX - 10, topY + 24, midX + 10, topY + 24);
            g.DrawLine(pen, midX - 6, topY + 4, midX + 6, topY + 4);
        }

        private static void DrawProvinces(Graphics g, Rectangle b, Pen pen)
        {
            int cx = b.Width / 2;
            int cy = b.Height / 2;
            g.DrawEllipse(pen, cx - 6, cy - 4, 12, 12);
            g.DrawLine(pen, cx, cy + 6, cx, b.Height - 6);
        }

        private static void DrawDatabase(Graphics g, Rectangle b, Pen pen)
        {
            int cx = b.Width / 2;
            int w = b.Width - 16;
            int top = b.Height / 2 - 10;
            int h = b.Height - top - 8;
            g.DrawEllipse(pen, cx - w / 2, top, w, 12);
            g.DrawLine(pen, cx - w / 2, top + 6, cx - w / 2, top + h);
            g.DrawLine(pen, cx + w / 2, top + 6, cx + w / 2, top + h);
            g.DrawEllipse(pen, cx - w / 2, top + h - 6, w, 12);
        }

        private static void DrawChart(Graphics g, Rectangle b, Pen pen)
        {
            int bottom = b.Height - 4;
            int x = b.Width / 2 - 20;
            int[] heights = { 10, 20, 30 };
            for (int i = 0; i < 3; i++)
            {
                int y = bottom - heights[i];
                g.DrawLine(pen, x + i * 14, y, x + i * 14, bottom);
                g.DrawLine(pen, x + i * 14, y, x + i * 14 + 6, y);
                g.DrawLine(pen, x + i * 14 + 6, y, x + i * 14 + 6, bottom);
            }
        }

        private static void DrawSettings(Graphics g, Rectangle b, Pen pen)
        {
            int cx = b.Width / 2;
            int cy = b.Height / 2;
            for (int i = 0; i < 8; i++)
            {
                double angle = i * Math.PI / 4;
                int x1 = cx + (int)(10 * Math.Cos(angle));
                int y1 = cy + (int)(10 * Math.Sin(angle));
                int x2 = cx + (int)(16 * Math.Cos(angle));
                int y2 = cy + (int)(16 * Math.Sin(angle));
                g.DrawLine(pen, x1, y1, x2, y2);
            }
            g.DrawEllipse(pen, cx - 10, cy - 10, 20, 20);
            g.DrawEllipse(pen, cx - 4, cy - 4, 8, 8);
        }

        private static void DrawFolder(Graphics g, Rectangle b, Pen pen)
        {
            int left = b.Width / 2 - 16;
            int top = b.Height / 2 - 12;
            int w = 32;
            int h = 24;
            g.DrawLine(pen, left, top + 3, left, top + h);
            g.DrawLine(pen, left, top + h, left + w, top + h);
            g.DrawLine(pen, left + w, top + h, left + w, top + 3);
            g.DrawLine(pen, left + 4, top + 3, left + 12, top - 3);
            g.DrawLine(pen, left + 12, top - 3, left + 20, top - 3);
            g.DrawLine(pen, left + 20, top - 3, left + w, top + 3);
        }

        private static void DrawInfo(Graphics g, Rectangle b, Pen pen)
        {
            int cx = b.Width / 2;
            int cy = b.Height / 2;
            g.DrawEllipse(pen, cx - 14, cy - 14, 28, 28);
            g.DrawEllipse(pen, cx - 2, cy - 8, 4, 4);
            g.DrawLine(pen, cx, cy - 2, cx, cy + 8);
        }

        private static void DrawExit(Graphics g, Rectangle b, Pen pen)
        {
            int cx = b.Width / 2;
            int cy = b.Height / 2;
            g.DrawRectangle(pen, cx - 15, cy - 12, 20, 24);
            g.DrawLine(pen, cx - 3, cy, b.Width - 6, cy);
            g.DrawLine(pen, b.Width - 10, cy - 5, b.Width - 4, cy);
            g.DrawLine(pen, b.Width - 10, cy + 5, b.Width - 4, cy);
        }
    }
}
