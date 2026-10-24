using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    /// <summary>
    /// تنظیمات برنامه — رابط مدرن و یکپارچه با داشبورد (RTL).
    /// </summary>
    public partial class FormSettings : Form
    {
        // پالت
        private static readonly Color HeaderTop = Color.FromArgb(11, 30, 63);
        private static readonly Color HeaderBottom = Color.FromArgb(30, 99, 199);
        private static readonly Color PrimaryColor = Color.FromArgb(0, 122, 204);
        private static readonly Color SuccessColor = Color.FromArgb(16, 185, 129);
        private static readonly Color DangerColor = Color.FromArgb(239, 68, 68);
        private static readonly Color WarningColor = Color.FromArgb(245, 158, 11);
        private static readonly Color VioletColor = Color.FromArgb(139, 92, 246);
        private static readonly Color CanvasColor = Color.FromArgb(244, 246, 250);
        private static readonly Color CardColor = Color.White;
        private static readonly Color BorderColor = Color.FromArgb(226, 232, 240);
        private static readonly Color TextPrimary = Color.FromArgb(15, 23, 42);
        private static readonly Color TextSecondary = Color.FromArgb(100, 116, 139);
        private static readonly Color CardTint = Color.FromArgb(250, 251, 253);

        private static readonly Color SidebarBase = Color.FromArgb(15, 23, 42);
        private static readonly Color SidebarItem = Color.FromArgb(25, 36, 62);
        private static readonly Color SidebarHover = Color.FromArgb(42, 58, 92);
        private static readonly Color SidebarText = Color.FromArgb(203, 213, 225);
        private static readonly Color SidebarIcon = Color.FromArgb(148, 163, 184);

        private Panel pnlDatabaseContent = null!;
        private Panel pnlPhotosContent = null!;
        private Panel pnlFontContent = null!;
        private Panel pnlMissingPhotosContent = null!;
        private Panel pnlExcelImportContent = null!;
        private Panel pnlLookupTablesContent = null!;
        private Panel selectedMenuButton = null!;
        private readonly System.Collections.Generic.List<Panel> menuButtons = new();
        private readonly System.Collections.Generic.Dictionary<Panel, Label> menuLabels = new();
        private readonly System.Collections.Generic.Dictionary<Panel, Color> menuPenColors = new();

        private TextBox txtDatabasePath = null!;
        private TextBox txtPhotosFolder = null!;
        private Label lblCurrentDatabase = null!;
        private Label lblCurrentPhotos = null!;
        private Label lblProviderStatus = null!;
        private Label lblConnectionStatus = null!;
        private Label lblProviderHint = null!;
        private readonly ToolTip toolTip = new();

        private ComboBox cmbFontFamily = null!;
        private NumericUpDown numTitleSize = null!;
        private NumericUpDown numLabelSize = null!;
        private NumericUpDown numTextBoxSize = null!;
        private NumericUpDown numButtonSize = null!;
        private NumericUpDown numBodySize = null!;
        private NumericUpDown numChartLabelSize = null!;
        private CheckBox chkBoldTitle = null!;
        private CheckBox chkBoldLabel = null!;
        private CheckBox chkBoldButton = null!;
        private CheckBox chkBoldChartLabel = null!;

        public FormSettings()
        {
            BuildUi();
            FontSettings.ApplyFontToForm(this);
            LoadCurrentSettings();
            Shown += (_, _) =>
            {
                Panel start = pnlDatabaseContent;
                switch (Environment.GetEnvironmentVariable("SC_SETTINGS_PAGE"))
                {
                    case "photos": start = pnlPhotosContent; break;
                    case "font": start = pnlFontContent; break;
                    case "missing": start = pnlMissingPhotosContent; break;
                    case "excel": start = pnlExcelImportContent; break;
                    case "lookup": start = pnlLookupTablesContent; break;
                }
                ShowContent(start, FindMenuByContent(start));
                DumpLayout();
                if (Environment.GetEnvironmentVariable("SC_SHOT_DIR") is string shotDir)
                {
                    var timer = new System.Windows.Forms.Timer { Interval = 600 };
                    timer.Tick += (_, _) =>
                    {
                        timer.Stop();
                        try
                        {
                            string page = Environment.GetEnvironmentVariable("SC_SETTINGS_PAGE") ?? "db";
                            string file = Path.Combine(shotDir, $"selfshot_{page}.png");
                            using var bmp = new Bitmap(Width, Height);
                            DrawToBitmap(bmp, new Rectangle(0, 0, Width, Height));
                            bmp.Save(file);
                        }
                        catch (Exception ex) { ErrorLogger.Log("FormSettings.SelfShot", ex); }
                    };
                    timer.Start();
                }
            };
        }

        private void DumpLayout()
        {
            try
            {
                string file = Environment.GetEnvironmentVariable("SC_LAYOUT_DUMP") ?? string.Empty;
                if (string.IsNullOrEmpty(file)) return;
                var sb = new System.Text.StringBuilder();
                void Walk(Control c, int depth)
                {
                    sb.AppendLine($"{new string(' ', depth * 2)}{c.GetType().Name} '{c.Text}' bounds=({c.Left},{c.Top},{c.Width}x{c.Height}) visible={c.Visible} dock={c.Dock}");
                    foreach (Control child in c.Controls) Walk(child, depth + 1);
                }
                Walk(this, 0);
                System.IO.File.WriteAllText(file, sb.ToString());
            }
            catch { }
        }

        private Font SafeFont(string family, float size, FontStyle style = FontStyle.Regular)
        {
            try { return new Font(family, size, style); }
            catch { return new Font("Tahoma", size, style); }
        }

        private static Color BlendWithWhite(Color color, float amount)
        {
            int r = (int)(color.R + (255 - color.R) * amount);
            int g = (int)(color.G + (255 - color.G) * amount);
            int b = (int)(color.B + (255 - color.B) * amount);
            return Color.FromArgb(r, g, b);
        }

        // ─────────────────────────────────────────────
        // ساختار اصلی
        // ─────────────────────────────────────────────
        private void BuildUi()
        {
            Text = "تنظیمات برنامه";
            MinimumSize = new Size(1040, 700);
            Size = new Size(1200, 780);
            StartPosition = FormStartPosition.CenterScreen;
            RightToLeft = RightToLeft.Yes;
            BackColor = CanvasColor;
            Font = new Font(FontSettings.FontFamilyName, 9F);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = CanvasColor,
                Padding = new Padding(14),
                Margin = new Padding(0)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 246F));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            root.Controls.Add(BuildSidebar(), 0, 0);
            root.Controls.Add(BuildMain(), 1, 0);
            Controls.Add(root);
        }

        // ─────────────────────────────────────────────
        // هدر صفحه
        // ─────────────────────────────────────────────
        private Panel BuildPageHeader(FlatIconKind icon, string title, string subtitle)
        {
            var header = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = HeaderTop,
                Margin = new Padding(0)
            };
            header.Paint += (_, e) =>
            {
                using var brush = new LinearGradientBrush(header.ClientRectangle, HeaderTop, HeaderBottom, 90f);
                e.Graphics.FillRectangle(brush, header.ClientRectangle);
                using var glow = new SolidBrush(Color.FromArgb(22, 255, 255, 255));
                e.Graphics.FillEllipse(glow, header.Width - 260, -100, 320, 240);
                using var pen = new Pen(Color.FromArgb(210, 255, 255, 255), 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                int s = 24;
                var iconRect = new Rectangle(header.Width - 52, (header.Height - s) / 2, s, s);
                FlatIcons.Draw(icon, e.Graphics, iconRect, pen);
            };

            header.Controls.Add(new Label
            {
                Text = subtitle,
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = Color.FromArgb(185, 200, 220),
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 64, 0)
            });
            header.Controls.Add(new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                Height = 32,
                Font = SafeFont(FontSettings.TitleFont?.FontFamily.Name ?? "Tahoma", 14F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 6, 64, 0)
            });
            header.Controls.Add(new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 2,
                BackColor = Color.FromArgb(120, 255, 255, 255)
            });
            return header;
        }

        // ─────────────────────────────────────────────
        // سایدبار
        // ─────────────────────────────────────────────
        private Control BuildSidebar()
        {
            var sidebar = new RoundedPanel(14)
            {
                Dock = DockStyle.Fill,
                BackColor = SidebarBase,
                Margin = new Padding(0, 0, 10, 0),
                Padding = new Padding(12, 16, 12, 12)
            };
            sidebar.Paint += (_, e) =>
            {
                using var glow = new SolidBrush(Color.FromArgb(18, 56, 189, 248));
                e.Graphics.FillEllipse(glow, sidebar.Width - 220, -140, 300, 260);
            };

            var stack = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                BackColor = Color.Transparent,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            // برند
            var brand = new Panel
            {
                Width = 212,
                Height = 58,
                BackColor = Color.Transparent,
                Margin = new Padding(0, 0, 0, 6)
            };
            brand.Controls.Add(new Label
            {
                Text = "سامانه مدیریت پرسنل و پست‌ها",
                Dock = DockStyle.Bottom,
                Height = 18,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8F),
                ForeColor = Color.FromArgb(148, 163, 184),
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 4, 0)
            });
            brand.Controls.Add(new Label
            {
                Text = "تنظیمات",
                Dock = DockStyle.Top,
                Height = 32,
                Font = SafeFont(FontSettings.TitleFont?.FontFamily.Name ?? "Tahoma", 13F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 4, 0)
            });
            stack.Controls.Add(brand);

            stack.Controls.Add(new Panel
            {
                Width = 212,
                Height = 1,
                BackColor = Color.FromArgb(51, 65, 85),
                Margin = new Padding(0, 0, 0, 10)
            });

            stack.Controls.Add(new Label
            {
                Text = "بخش‌ها",
                Width = 212,
                Height = 24,
                Font = SafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 116, 139),
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 4, 0)
            });

            pnlDatabaseContent = CreateDatabaseContent();
            pnlPhotosContent = CreatePhotosContent();
            pnlFontContent = CreateFontContent();
            pnlMissingPhotosContent = CreateMissingPhotosContent();
            pnlExcelImportContent = CreateExcelImportContent();
            pnlLookupTablesContent = CreateLookupTablesContent();

            stack.Controls.Add(CreateMenuButton(FlatIconKind.Database, "دیتابیس", pnlDatabaseContent));
            stack.Controls.Add(CreateMenuButton(FlatIconKind.Personnel, "عکس‌ها", pnlPhotosContent));
            stack.Controls.Add(CreateMenuButton(FlatIconKind.Settings, "فونت", pnlFontContent));
            stack.Controls.Add(CreateMenuButton(FlatIconKind.Info, "پرسنل بدون عکس", pnlMissingPhotosContent));
            stack.Controls.Add(CreateMenuButton(FlatIconKind.Folder, "وارد کردن از اکسل", pnlExcelImportContent));
            stack.Controls.Add(CreateMenuButton(FlatIconKind.Chart, "جداول مرجع", pnlLookupTablesContent));

            stack.Controls.Add(new Panel
            {
                Width = 212,
                Height = 1,
                BackColor = Color.FromArgb(51, 65, 85),
                Margin = new Padding(0, 10, 0, 10)
            });

            var btnBack = CreateSidebarAction("بازگشت به داشبورد", () => Close());
            stack.Controls.Add(btnBack);

            sidebar.Controls.Add(stack);
            return sidebar;
        }

        private Panel CreateSidebarAction(string text, Action onClick)
        {
            var btn = new RoundedPanel(8)
            {
                Width = 212,
                Height = 44,
                BackColor = Color.FromArgb(30, 41, 63),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 0, 4)
            };
            var lbl = new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9F, FontStyle.Bold),
                ForeColor = SidebarText,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 12, 0),
                Cursor = Cursors.Hand
            };
            btn.Controls.Add(lbl);
            btn.Click += (_, _) => onClick();
            lbl.Click += (_, _) => onClick();
            btn.MouseEnter += (_, _) => btn.BackColor = SidebarHover;
            btn.MouseLeave += (_, _) => btn.BackColor = Color.FromArgb(30, 41, 63);
            return btn;
        }

        private Panel CreateMenuButton(FlatIconKind icon, string text, Panel target)
        {
            var btn = new RoundedPanel(8)
            {
                Width = 212,
                Height = 46,
                BackColor = SidebarItem,
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 0, 4),
                Tag = target
            };

            var lbl = new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9.5F),
                ForeColor = SidebarText,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 48, 0),
                Cursor = Cursors.Hand
            };
            btn.Controls.Add(lbl);
            menuLabels[btn] = lbl;
            menuPenColors[btn] = SidebarIcon;

            btn.Paint += (_, e) =>
            {
                using var pen = new Pen(menuPenColors[btn], 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                int s = 22;
                var iconRect = new Rectangle(btn.Width - 40, (btn.Height - s) / 2, s, s);
                FlatIcons.Draw(icon, e.Graphics, iconRect, pen);
            };

            void Activate(object? s, EventArgs e) => ShowContent(target, btn);
            btn.Click += Activate;
            lbl.Click += Activate;

            btn.MouseEnter += (_, _) => { if (selectedMenuButton != btn) btn.BackColor = SidebarHover; };
            btn.MouseLeave += (_, _) => { if (selectedMenuButton != btn) btn.BackColor = SidebarItem; };

            menuButtons.Add(btn);
            return btn;
        }

        // ─────────────────────────────────────────────
        // ناحیه اصلی
        // ─────────────────────────────────────────────
        private Control BuildMain()
        {
            var main = new RoundedPanel(16)
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };

            main.Controls.Add(pnlDatabaseContent);
            main.Controls.Add(pnlPhotosContent);
            main.Controls.Add(pnlFontContent);
            main.Controls.Add(pnlMissingPhotosContent);
            main.Controls.Add(pnlExcelImportContent);
            main.Controls.Add(pnlLookupTablesContent);

            pnlDatabaseContent.Visible = true;
            pnlPhotosContent.Visible = true;
            pnlFontContent.Visible = true;
            pnlMissingPhotosContent.Visible = true;
            pnlExcelImportContent.Visible = true;
            pnlLookupTablesContent.Visible = true;
            return main;
        }

        private Panel? FindMenuByContent(Panel content)
        {
            return menuButtons.FirstOrDefault(btn => ReferenceEquals(btn.Tag, content));
        }

        private void ShowContent(Panel contentPanel, Panel? menuButton)
        {
            foreach (Panel p in new[] { pnlDatabaseContent, pnlPhotosContent, pnlFontContent,
                                         pnlMissingPhotosContent, pnlExcelImportContent, pnlLookupTablesContent })
            {
                p.Visible = false;
            }

            if (contentPanel != null)
            {
                contentPanel.Visible = true;
                contentPanel.BringToFront();
            }

            selectedMenuButton = menuButton!;
            foreach (var pair in menuLabels)
            {
                bool selected = ReferenceEquals(pair.Key, menuButton);
                pair.Key.BackColor = selected ? PrimaryColor : SidebarItem;
                pair.Value.ForeColor = selected ? Color.White : SidebarText;
                menuPenColors[pair.Key] = selected ? Color.White : SidebarIcon;
                pair.Key.Invalidate();
            }
        }

        // ─────────────────────────────────────────────
        // قالب صفحه (هدر + بدنه اسکرول + فوتر)
        // ─────────────────────────────────────────────
        private Panel CreatePageShell(Panel header, Panel body, Panel footer)
        {
            var shell = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Margin = new Padding(0)
            };

            var tlp = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = CardColor,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 74F));
            tlp.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tlp.RowStyles.Add(new RowStyle(SizeType.Absolute, 64F));

            tlp.Controls.Add(header, 0, 0);
            tlp.Controls.Add(body, 0, 1);
            tlp.Controls.Add(footer, 0, 2);

            shell.Controls.Add(tlp);
            return shell;
        }

        private Panel CreateScrollBody()
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                AutoScroll = true,
                Padding = new Padding(14, 10, 14, 6),
                Margin = new Padding(0)
            };
        }

        private void AddCards(Panel body, params Control[] cards)
        {
            for (int i = cards.Length - 1; i >= 0; i--)
                body.Controls.Add(cards[i]);
        }

        private Panel CreateCard()
        {
            var card = new RoundedPanel(12)
            {
                Dock = DockStyle.Top,
                BackColor = CardTint,
                Margin = new Padding(0, 8, 0, 8),
                Padding = new Padding(22, 18, 22, 18)
            };
            card.Paint += (_, e) =>
            {
                using var pen = new Pen(BorderColor);
                e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
            };
            return card;
        }

        private Label CreateCardTitle(string text, FlatIconKind? icon = null)
        {
            var lbl = new Label
            {
                Text = text,
                Dock = DockStyle.Top,
                Height = 30,
                Font = SafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 11.5F, FontStyle.Bold),
                ForeColor = TextPrimary,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 4, 0)
            };
            if (icon.HasValue)
            {
                lbl.Paint += (_, e) =>
                {
                    using var pen = new Pen(TextSecondary, 1.6f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                    int s = 18;
                    var iconRect = new Rectangle(lbl.Width - 30, (lbl.Height - s) / 2, s, s);
                    FlatIcons.Draw(icon.Value, e.Graphics, iconRect, pen);
                };
                lbl.Padding = new Padding(0, 0, 36, 0);
            }
            return lbl;
        }

        private Panel CreateFooterRow()
        {
            var row = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Margin = new Padding(0),
                Padding = new Padding(14, 8, 14, 8)
            };
            row.Paint += (_, e) =>
            {
                using var pen = new Pen(BorderColor);
                e.Graphics.DrawLine(pen, 0, 0, row.Width, 0);
            };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = CardColor,
                Padding = new Padding(0)
            };

            var btnSave = CreateButton("ذخیره تنظیمات", SuccessColor, 180, () => SaveSettings());
            var btnReset = CreateButton("بازنشانی", WarningColor, 150, () => ResetSettings());
            var btnCancel = CreateButton("لغو", Color.FromArgb(100, 116, 139), 120, () => Close());

            toolTip.SetToolTip(btnSave, "ذخیره همه تغییرات این صفحه");
            toolTip.SetToolTip(btnReset, "بازگرداندن تنظیمات به حالت پیش‌فرض");
            toolTip.SetToolTip(btnCancel, "بستن پنجره بدون ذخیره");

            flow.Controls.Add(btnSave);
            flow.Controls.Add(btnReset);
            flow.Controls.Add(btnCancel);

            row.Controls.Add(flow);
            return row;
        }

        // ─────────────────────────────────────────────
        // دکمه‌ها
        // ─────────────────────────────────────────────
        private PillButton CreateButton(string text, Color backColor, int width, Action onClick, FlatIconKind? icon = null)
        {
            return new PillButton(text, backColor, icon)
            {
                Width = width,
                Height = 42,
                Margin = new Padding(0, 0, 8, 0),
                Clicked = onClick
            };
        }

        /// <summary>
        /// دکمه قرصی شکل با گوشه‌های گرد، آیکون خطی و حالت‌های hover — برای وضوح کامل در همه بخش‌ها.
        /// </summary>
        private sealed class PillButton : Control
        {
            private readonly Color _accent;
            private readonly FlatIconKind? _icon;
            private Color _back;

            public Action? Clicked;

            public PillButton(string text, Color accent, FlatIconKind? icon = null)
            {
                Text = text;
                _accent = accent;
                _icon = icon;
                _back = accent;
                Height = 42;
                Cursor = Cursors.Hand;
                Font = new Font("Tahoma", 9.5F, FontStyle.Bold);
                ForeColor = Color.White;
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            }

            protected override void OnMouseEnter(EventArgs e)
            {
                base.OnMouseEnter(e);
                _back = BlendWithWhite(_accent, 0.12f);
                Invalidate();
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                base.OnMouseLeave(e);
                _back = _accent;
                Invalidate();
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                base.OnMouseDown(e);
                _back = BlendWithWhite(_accent, 0.22f);
                Invalidate();
            }

            protected override void OnMouseUp(MouseEventArgs e)
            {
                base.OnMouseUp(e);
                _back = _accent;
                Invalidate();
            }

            protected override void OnClick(EventArgs e)
            {
                base.OnClick(e);
                Clicked?.Invoke();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                var path = new GraphicsPath();
                int r = 10;
                path.AddArc(0, 0, r, r, 180, 90);
                path.AddArc(Width - r, 0, r, r, 270, 90);
                path.AddArc(Width - r, Height - r, r, r, 0, 90);
                path.AddArc(0, Height - r, r, r, 90, 90);
                path.CloseFigure();
                using (var brush = new SolidBrush(_back))
                    g.FillPath(brush, path);
                using (var pen = new Pen(BlendWithWhite(_accent, 0.05f)))
                    g.DrawPath(pen, path);

                var textArea = new Rectangle(12, 0, Width - 24, Height);
                if (_icon.HasValue)
                {
                    int s = 20;
                    var iconRect = new Rectangle(Width - 42, (Height - s) / 2, s, s);
                    using var pen = new Pen(Color.White, 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                    FlatIcons.Draw(_icon.Value, g, iconRect, pen);
                    textArea = new Rectangle(12, 0, Width - 56, Height);
                }

                using var sf = new StringFormat(StringFormatFlags.DirectionRightToLeft | StringFormatFlags.NoWrap)
                {
                    Alignment = StringAlignment.Near,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter
                };
                using var brushText = new SolidBrush(ForeColor);
                g.DrawString(Text, Font, brushText, textArea, sf);
            }
        }

        private Label CreateFieldLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9.5F, FontStyle.Bold),
                ForeColor = TextPrimary,
                TextAlign = ContentAlignment.MiddleRight
            };
        }

        private Label CreateIconFieldLabel(string text, FlatIconKind icon)
        {
            var lbl = CreateFieldLabel(text);
            lbl.Padding = new Padding(0, 0, 30, 0);
            lbl.Paint += (_, e) =>
            {
                using var pen = new Pen(TextSecondary, 1.6f) { StartCap = LineCap.Round, EndCap = LineCap.Round };
                int s = 16;
                var iconRect = new Rectangle(lbl.Width - 25, (lbl.Height - s) / 2, s, s);
                FlatIcons.Draw(icon, e.Graphics, iconRect, pen);
            };
            return lbl;
        }

        // ─────────────────────────────────────────────
        // صفحه دیتابیس
        // ─────────────────────────────────────────────
        private Panel CreateDatabaseContent()
        {
            var header = BuildPageHeader(FlatIconKind.Database, "دیتابیس", "مسیر فایل Access و وضعیت موتور پایگاه داده");
            var body = CreateScrollBody();

            // کارت ۱: فایل پایگاه داده
            var fileCard = CreateCard();
            fileCard.Height = 220;
            toolTip.SetToolTip(fileCard, "مسیر فایل Access مورد استفاده برنامه");

            var fileLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 4,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0, 8, 0, 0)
            };
            fileLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            fileLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            fileLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));
            fileLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 8F));
            fileLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48F));
            fileLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));

            fileLayout.Controls.Add(CreateIconFieldLabel("مسیر فایل:", FlatIconKind.Folder), 0, 0);

            txtDatabasePath = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 9),
                ReadOnly = true,
                BackColor = Color.FromArgb(248, 250, 252),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 4, 0, 4)
            };
            fileLayout.Controls.Add(txtDatabasePath, 1, 0);
            toolTip.SetToolTip(txtDatabasePath, "فایل Access فعلی برنامه");

            var btnBrowse = CreateButton("انتخاب فایل دیتابیس", PrimaryColor, 230, () => BrowseDatabase(), FlatIconKind.Folder);
            btnBrowse.Dock = DockStyle.Right;
            btnBrowse.Margin = new Padding(0, 4, 0, 4);
            fileLayout.Controls.Add(btnBrowse, 1, 2);
            toolTip.SetToolTip(btnBrowse, "انتخاب فایل جدید پایگاه داده");

            lblCurrentDatabase = new Label
            {
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8.5F),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.MiddleRight,
                AutoEllipsis = true
            };
            fileLayout.Controls.Add(lblCurrentDatabase, 0, 3);
            fileLayout.SetColumnSpan(lblCurrentDatabase, 2);

            fileCard.Controls.Add(fileLayout);
            fileCard.Controls.Add(CreateCardTitle("فایل پایگاه داده", FlatIconKind.Folder));

            // کارت ۲: موتور پایگاه داده
            var providerCard = CreateCard();
            providerCard.Height = 316;

            var pLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 5,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0, 6, 0, 0)
            };
            pLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            pLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            pLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 10F));
            pLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 32F));
            pLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28F));
            pLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));
            pLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30F));
            pLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 96F));

            pLayout.Controls.Add(CreateCardTitle("موتور پایگاه داده", FlatIconKind.Database), 0, 0);
            pLayout.SetColumnSpan(pLayout.GetControlFromPosition(0, 0)!, 3);

            lblProviderStatus = new Label
            {
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9.5F, FontStyle.Bold),
                ForeColor = TextPrimary,
                TextAlign = ContentAlignment.MiddleRight,
                AutoEllipsis = true
            };
            pLayout.Controls.Add(lblProviderStatus, 0, 1);
            pLayout.SetColumnSpan(lblProviderStatus, 3);

            lblConnectionStatus = new Label
            {
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.MiddleRight,
                AutoEllipsis = true
            };
            pLayout.Controls.Add(lblConnectionStatus, 0, 2);
            pLayout.SetColumnSpan(lblConnectionStatus, 3);

            var actionRow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent,
                Padding = new Padding(0)
            };

            var btnTest = CreateButton("تست اتصال", PrimaryColor, 220, () => TestProvider(), FlatIconKind.Database);
            actionRow.Controls.Add(btnTest);
            toolTip.SetToolTip(btnTest, "بررسی وضعیت موتور و اتصال به فایل");

            var btnInstall = CreateButton("نصب / تعمیر موتور ACE", SuccessColor, 300, () => InstallProvider(), FlatIconKind.Settings);
            actionRow.Controls.Add(btnInstall);
            toolTip.SetToolTip(btnInstall, "نصب خودکار یا تعمیر موتور Access Database Engine");

            pLayout.Controls.Add(actionRow, 0, 3);
            pLayout.SetColumnSpan(actionRow, 3);

            lblProviderHint = new Label
            {
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8.5F),
                ForeColor = Color.FromArgb(30, 64, 110),
                TextAlign = ContentAlignment.TopRight,
                Padding = new Padding(12, 8, 12, 8),
                BackColor = Color.FromArgb(235, 243, 252),
                Text = "اگر هنگام اجرا خطای «Microsoft.ACE.OLEDB.12.0 provider is not registered» دیدید، موتور Access Database Engine " +
                       "نصب نیست یا معماری آن با برنامه هم‌خوان نیست. با دکمه «نصب / تعمیر موتور ACE» می‌توانید آن را خودکار نصب کنید."
            };
            pLayout.Controls.Add(lblProviderHint, 0, 4);
            pLayout.SetColumnSpan(lblProviderHint, 3);

            providerCard.Controls.Add(pLayout);

            AddCards(body, fileCard, providerCard);
            return CreatePageShell(header, body, CreateFooterRow());
        }

        // ─────────────────────────────────────────────
        // صفحه عکس‌ها
        // ─────────────────────────────────────────────
        private Panel CreatePhotosContent()
        {
            var header = BuildPageHeader(FlatIconKind.Personnel, "عکس‌ها", "پوشه ذخیره عکس پرسنل");
            var body = CreateScrollBody();

            var card = CreateCard();
            card.Height = 220;
            toolTip.SetToolTip(card, "پوشه‌ای که عکس پرسنل در آن ذخیره می‌شود");

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 4,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0, 8, 0, 0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 8F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));

            layout.Controls.Add(CreateIconFieldLabel("مسیر پوشه:", FlatIconKind.Folder), 0, 0);

            txtPhotosFolder = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 9),
                ReadOnly = true,
                BackColor = Color.FromArgb(248, 250, 252),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 4, 0, 4)
            };
            layout.Controls.Add(txtPhotosFolder, 1, 0);
            toolTip.SetToolTip(txtPhotosFolder, "پوشه عکس‌های فعلی برنامه");

            var btnBrowse = CreateButton("انتخاب پوشه عکس‌ها", PrimaryColor, 230, () => BrowsePhotos(), FlatIconKind.Folder);
            btnBrowse.Dock = DockStyle.Right;
            btnBrowse.Margin = new Padding(0, 4, 0, 4);
            layout.Controls.Add(btnBrowse, 1, 2);
            toolTip.SetToolTip(btnBrowse, "انتخاب پوشه جدید برای ذخیره عکس‌ها");

            lblCurrentPhotos = new Label
            {
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8.5F),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.MiddleRight,
                AutoEllipsis = true
            };
            layout.Controls.Add(lblCurrentPhotos, 0, 3);
            layout.SetColumnSpan(lblCurrentPhotos, 2);

            card.Controls.Add(layout);
            card.Controls.Add(CreateCardTitle("پوشه عکس‌ها", FlatIconKind.Folder));
            AddCards(body, card);
            return CreatePageShell(header, body, CreateFooterRow());
        }

        // ─────────────────────────────────────────────
        // صفحه فونت
        // ─────────────────────────────────────────────
        private Panel CreateFontContent()
        {
            var header = BuildPageHeader(FlatIconKind.Settings, "فونت", "نوع و اندازه فونت‌های برنامه");
            var body = CreateScrollBody();

            var card = CreateCard();
            card.Height = 500;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 9,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0, 6, 0, 0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            for (int i = 0; i < 9; i++)
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));

            var familyRow = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };
            var familyLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };
            familyLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            familyLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            familyLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            familyLayout.Controls.Add(CreateFieldLabel("نوع فونت:"), 0, 0);
            cmbFontFamily = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = SafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 9.5F),
                Margin = new Padding(0, 4, 0, 6)
            };
            cmbFontFamily.Items.AddRange(new string[]
            {
                "Tahoma", "Arial", "Segoe UI", "Calibri", "Times New Roman",
                "B Nazanin", "B Mitra", "B Lotus", "B Titr", "IRANSans", "Vazir"
            });
            familyLayout.Controls.Add(cmbFontFamily, 1, 0);
            familyRow.Controls.Add(familyLayout);
            layout.Controls.Add(familyRow, 0, 0);
            layout.Controls.Add(new Panel { Dock = DockStyle.Fill, BackColor = BorderColor }, 0, 1);

            layout.Controls.Add(CreateFieldLabel("اندازه فونت‌ها:"), 0, 2);

            AddFontRow(layout, 3, "سرتیتر:", out numTitleSize, out chkBoldTitle, 16);
            AddFontRow(layout, 4, "برچسب:", out numLabelSize, out chkBoldLabel, 12);
            AddFontRow(layout, 5, "دکمه:", out numButtonSize, out chkBoldButton, 12);
            AddFontRow(layout, 6, "متن ورودی:", out numTextBoxSize, out _, 11);
            AddFontRow(layout, 7, "متن عادی:", out numBodySize, out _, 10);
            AddFontRow(layout, 8, "نمودارها:", out numChartLabelSize, out chkBoldChartLabel, 10);

            card.Controls.Add(layout);
            AddCards(body, card);
            return CreatePageShell(header, body, CreateFooterRow());
        }

        private void AddFontRow(TableLayoutPanel layout, int row, string caption,
            out NumericUpDown numeric, out CheckBox bold, int defaultValue)
        {
            var rowPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };

            var rowLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };
            rowLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180F));
            rowLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            rowLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130F));
            rowLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            rowLayout.Controls.Add(CreateFieldLabel(caption), 0, 0);

            numeric = new NumericUpDown
            {
                Dock = DockStyle.Fill,
                Minimum = 8,
                Maximum = 72,
                Value = defaultValue,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 4, 12, 4)
            };
            rowLayout.Controls.Add(numeric, 1, 0);

            bold = new CheckBox
            {
                Text = "ضخیم",
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8.5F),
                ForeColor = TextSecondary,
                Margin = new Padding(0, 8, 0, 0)
            };
            rowLayout.Controls.Add(bold, 2, 0);

            rowPanel.Controls.Add(rowLayout);
            layout.Controls.Add(rowPanel, 0, row);
        }

        // ─────────────────────────────────────────────
        // صفحه پرسنل بدون عکس
        // ─────────────────────────────────────────────
        private Panel CreateMissingPhotosContent()
        {
            var header = BuildPageHeader(FlatIconKind.Info, "پرسنل بدون عکس", "بررسی و مدیریت پرسنل بدون عکس");
            var body = CreateScrollBody();

            var card = CreateCard();
            card.Height = 268;
            card.Controls.Add(CreateCardTitle("بررسی پرسنل بدون عکس", FlatIconKind.Personnel));

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0, 8, 0, 0)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 120F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 14F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52F));

            layout.Controls.Add(new Label
            {
                Text = "✓ مشاهده لیست کامل پرسنل بدون عکس\n" +
                       "✓ خروجی اکسل برای گزارش‌گیری\n" +
                       "✓ امکان ویرایش و حذف مستقیم\n" +
                       "✓ بروزرسانی لحظه‌ای اطلاعات",
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9.5F),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.TopRight
            }, 0, 0);

            var btnOpen = CreateButton("مشاهده لیست پرسنل بدون عکس", SuccessColor, 320, () =>
            {
                try
                {
                    using var form = new FormMissingPhotos();
                    form.ShowDialog(this);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"خطا در باز کردن فرم:\n{ex.Message}", "خطا",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }, FlatIconKind.Personnel);
            btnOpen.Dock = DockStyle.Right;
            btnOpen.Width = 360;
            toolTip.SetToolTip(btnOpen, "باز کردن لیست پرسنل بدون عکس");
            layout.Controls.Add(btnOpen, 0, 2);

            card.Controls.Add(layout);
            AddCards(body, card);
            return CreatePageShell(header, body, CreateFooterRow());
        }

        // ─────────────────────────────────────────────
        // صفحه وارد کردن از اکسل
        // ─────────────────────────────────────────────
        private Panel CreateExcelImportContent()
        {
            var header = BuildPageHeader(FlatIconKind.Folder, "وارد کردن از اکسل", "وارد کردن اطلاعات پرسنلی از فایل اکسل");
            var body = CreateScrollBody();

            var card = CreateCard();
            card.Height = 366;
            card.Controls.Add(CreateCardTitle("راهنمای وارد کردن", FlatIconKind.Chart));

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0, 8, 0, 0)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 128F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 92F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 14F));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52F));

            layout.Controls.Add(new Label
            {
                Text = "ستون‌های مورد انتظار در فایل اکسل (ردیف اول = سرستون):\n" +
                       "استان | شهر | امور انتقال | اداره | ناحیه | نام پست | سطح ولتاژ | شیفت | جنسیت\n" +
                       "نام | نام خانوادگی | نام پدر | ش.پرسنلی | کدملی | موبایل\n" +
                       "تاریخ تولد | تاریخ استخدام | تاریخ شروع بکار | نوع قرارداد | سطح شغل\n" +
                       "شرکت | مدرک تحصیلی | رشته تحصیلی | عنوان شغلی اصلی | فعالیت فعلی | وضعیت",
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8.5F),
                ForeColor = Color.FromArgb(27, 94, 32),
                TextAlign = ContentAlignment.TopRight,
                BackColor = Color.FromArgb(232, 248, 232),
                Padding = new Padding(12, 8, 12, 8)
            }, 0, 0);

            layout.Controls.Add(new Label
            {
                Text = "قوانین پیش‌فرض هنگام وارد کردن:\n" +
                       "• تاریخ خالی ← 1300/01/01     • عنوان شغل خالی ← غیرمرتبط     • سایر فیلدهای خالی ← داده‌ای وجود ندارد\n" +
                       "• کدملی تکراری: ردیف نادیده گرفته می‌شود     • تاریخ میلادی: به شمسی تبدیل می‌شود",
                Dock = DockStyle.Fill,
                Font = SafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 8.5F),
                ForeColor = Color.FromArgb(230, 81, 0),
                TextAlign = ContentAlignment.TopRight,
                BackColor = Color.FromArgb(255, 243, 224),
                Padding = new Padding(12, 8, 12, 8)
            }, 0, 1);

            var btnOpen = CreateButton("باز کردن پنجره وارد کردن از اکسل", Color.FromArgb(0, 150, 136), 340, () =>
            {
                try
                {
                    using var form = new FormImportExcel();
                    form.ShowDialog(this);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"خطا در باز کردن فرم:\n{ex.Message}", "خطا",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }, FlatIconKind.Folder);
            btnOpen.Dock = DockStyle.Right;
            btnOpen.Width = 400;
            toolTip.SetToolTip(btnOpen, "باز کردن پنجره وارد کردن اطلاعات از فایل اکسل");
            layout.Controls.Add(btnOpen, 0, 3);

            card.Controls.Add(layout);
            AddCards(body, card);
            return CreatePageShell(header, body, CreateFooterRow());
        }

        // ─────────────────────────────────────────────
        // صفحه جداول مرجع
        // ─────────────────────────────────────────────
        private Panel CreateLookupTablesContent()
        {
            var header = BuildPageHeader(FlatIconKind.Chart, "جداول مرجع", "اضافه کردن، ویرایش و حذف رکوردهای جداول پایه");
            var body = CreateScrollBody();

            var card = CreateCard();
            card.Controls.Add(CreateCardTitle("جداول مرجع", FlatIconKind.Chart));

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Padding = new Padding(0, 8, 0, 0)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));

            var rows = new[]
            {
                new[] { ("استان‌ها", "Provinces", "ProvinceID", "ProvinceName", "استان‌ها", "", "", "", ""),
                         ("شهرها", "Cities", "CityID", "CityName", "شهرها", "Provinces", "ProvinceID", "ProvinceName", "ProvinceID"),
                         ("امور انتقال", "TransferAffairs", "AffairID", "AffairName", "امور انتقال", "Provinces", "ProvinceID", "ProvinceName", "ProvinceID") },
                new[] { ("ادارات بهره‌برداری", "OperationDepartments", "DeptID", "DeptName", "ادارات بهره‌برداری", "TransferAffairs", "AffairID", "AffairName", "AffairID"),
                         ("نواحی", "Districts", "DistrictID", "DistrictName", "نواحی", "OperationDepartments", "DeptID", "DeptName", "DeptID"),
                         ("نام پست‌ها", "PostsNames", "PostNameID", "PostName", "نام پست‌ها", "Districts", "DistrictID", "DistrictName", "DistrictID") },
                new[] { ("سطح ولتاژ", "VoltageLevels", "VoltageID", "VoltageName", "سطح ولتاژ", "", "", "", ""),
                         ("استانداردهای پست", "PostStandards", "StandardID", "StandardName", "استانداردهای پست", "", "", "", ""),
                         ("نوع پست", "PostTypes", "TypeID", "TypeName", "نوع پست", "", "", "", "") },
                new[] { ("اتصالات توزیع", "DistributedConnections", "ConnID", "ConnName", "اتصالات توزیع", "", "", "", ""),
                         ("انواع عایق", "InsulationTypes", "InsID", "InsName", "انواع عایق", "", "", "", ""),
                         ("نوع پست دو", "PostTypeTwos", "PT2ID", "PT2Name", "نوع پست دو", "", "", "", "") },
                new[] { ("ثابت/سیار", "FixedMobiles", "FMID", "FMName", "ثابت/سیار", "", "", "", ""),
                         ("وضعیت مدار", "CircuitStatuses", "CircuitID", "CircuitName", "وضعیت مدار", "", "", "", ""),
                         ("دیزل ژنراتورها", "DieselGenerators", "DieselID", "DieselName", "دیزل ژنراتورها", "", "", "", "") },
                new[] { ("فیدرهای توزیع", "DistributionFeeds", "FeedID", "FeedName", "فیدرهای توزیع", "", "", "", ""),
                         ("وضعیت آب", "WaterStatuses", "WaterID", "WaterName", "وضعیت آب", "", "", "", ""),
                         ("مهمان‌خانه‌ها", "GuestHouses", "GuestID", "GuestName", "مهمان‌خانه‌ها", "", "", "", "") },
                new[] { ("شیفت کاری", "WorkShift", "WorkShiftID", "WorkShiftName", "شیفت کاری", "", "", "", ""),
                         ("جنسیت", "Gender", "GenderID", "GenderName", "جنسیت", "", "", "", ""),
                         ("نوع قرارداد", "ContractType", "ContractTypeID", "ContractTypeName", "نوع قرارداد", "", "", "", "") },
                new[] { ("سطح شغلی", "JobLevel", "JobLevelID", "JobLevelName", "سطح شغلی", "", "", "", ""),
                         ("شرکت‌ها", "Company", "CompanyID", "CompanyName", "شرکت‌ها", "", "", "", ""),
                         ("مدرک تحصیلی", "Degree", "DegreeID", "DegreeName", "مدرک تحصیلی", "", "", "", "") },
                new[] { ("رشته تحصیلی", "DegreeField", "DegreeFieldID", "DegreeFieldName", "رشته تحصیلی", "", "", "", ""),
                         ("وضعیت حضور", "StatusPresence", "StatusID", "StatusName", "وضعیت حضور", "", "", "", ""),
                         ("چارت امور", "ChartAffairs", "ChartID", "ChartName", "چارت امور", "", "", "", "") }
            };

            for (int r = 0; r < rows.Length; r++)
            {
                grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 52F));
                for (int c = 0; c < 3; c++)
                {
                    (string text, string table, string id, string name, string display, string pt, string pid, string pname, string fk) = rows[r][c];
                    var btn = CreateButton(text, VioletColor, 0, () => { });
                    btn.Dock = DockStyle.Fill;
                    btn.Margin = new Padding(4);
                    btn.Text = text;
                    btn.Click += (_, _) =>
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(pt))
                            {
                                using var form = new FormLookupTableManagerWithDependency(
                                    table, id, name, display, pt, pid, pname, display, fk);
                                form.ShowDialog(this);
                            }
                            else
                            {
                                using var form = new FormLookupTableManager(table, id, name, display);
                                form.ShowDialog(this);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(this, $"خطا در باز کردن فرم:\n{ex.Message}", "خطا",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    };
                    grid.Controls.Add(btn, c, r);
                }
            }

            card.Controls.Add(grid);
            card.Height = rows.Length * 52 + 96;
            AddCards(body, card);
            return CreatePageShell(header, body, CreateFooterRow());
        }

        // ─────────────────────────────────────────────
        // منطق
        // ─────────────────────────────────────────────
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
            numChartLabelSize.Value = (decimal)FontSettings.ChartLabelFontSize;
            chkBoldTitle.Checked = FontSettings.TitleFontBold;
            chkBoldLabel.Checked = FontSettings.LabelFontBold;
            chkBoldButton.Checked = FontSettings.ButtonFontBold;
            chkBoldChartLabel.Checked = FontSettings.ChartLabelFontBold;

            UpdateProviderStatus();
        }

        private void BrowseDatabase()
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Access Database (*.accdb)|*.accdb|All Files (*.*)|*.*",
                Title = "انتخاب فایل دیتابیس",
                InitialDirectory = Path.GetDirectoryName(AppSettings.DatabasePath)
            };
            if (ofd.ShowDialog(this) == DialogResult.OK)
            {
                txtDatabasePath.Text = ofd.FileName;
                UpdateProviderStatus();
            }
        }

        private void BrowsePhotos()
        {
            using var fbd = new FolderBrowserDialog
            {
                Description = "انتخاب پوشه عکس‌ها",
                SelectedPath = AppSettings.PhotosFolder,
                ShowNewFolderButton = true
            };
            if (fbd.ShowDialog(this) == DialogResult.OK)
                txtPhotosFolder.Text = fbd.SelectedPath;
        }

        private void UpdateProviderStatus()
        {
            DatabaseProvider.Invalidate();
            string? provider = DatabaseProvider.ProviderName;

            if (!string.IsNullOrWhiteSpace(provider))
            {
                lblProviderStatus.Text = $"✓ موتور پایگاه داده: {provider}  ({DatabaseProvider.Bitness})";
                lblProviderStatus.ForeColor = SuccessColor;
            }
            else
            {
                lblProviderStatus.Text = $"✗ موتور پایگاه داده Access نصب نیست  ({DatabaseProvider.Bitness})";
                lblProviderStatus.ForeColor = DangerColor;
            }

            string path = txtDatabasePath?.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                lblConnectionStatus.Text = "⚠ فایل پایگاه داده در مسیر انتخاب‌شده پیدا نشد.";
                lblConnectionStatus.ForeColor = WarningColor;
                return;
            }

            if (DbHelper.TestDatabasePath(path, out string error))
            {
                lblConnectionStatus.Text = "✓ اتصال به فایل Access با موفقیت تست شد.";
                lblConnectionStatus.ForeColor = SuccessColor;
            }
            else
            {
                lblConnectionStatus.Text = "✗ تست اتصال ناموفق: " + error;
                lblConnectionStatus.ForeColor = DangerColor;
            }
        }

        private void TestProvider()
        {
            Cursor previous = Cursor;
            Cursor = Cursors.WaitCursor;
            try
            {
                UpdateProviderStatus();
                MessageBox.Show(this, lblProviderStatus.Text + "\n\n" + lblConnectionStatus.Text, "تست پایگاه داده",
                    MessageBoxButtons.OK, lblConnectionStatus.ForeColor == DangerColor ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            finally
            {
                Cursor = previous;
            }
        }

        private void InstallProvider()
        {
            Cursor previous = Cursor;
            Cursor = Cursors.WaitCursor;
            try
            {
                bool installed = DatabaseProvider.InstallProvider(this);
                DatabaseProvider.Invalidate();
                UpdateProviderStatus();

                if (installed)
                {
                    MessageBox.Show(this,
                        "✓ موتور پایگاه داده با موفقیت نصب/تعمیر شد.\n\nبرای اعمال کامل تغییرات، برنامه را ببندید و دوباره اجرا کنید.",
                        "نصب پیش‌نیاز", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormSettings.InstallProvider", ex);
                MessageBox.Show(this, $"خطا در نصب موتور پایگاه داده:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = previous;
            }
        }

        private void SaveSettings()
        {
            try
            {
                if (!File.Exists(txtDatabasePath.Text))
                {
                    if (MessageBox.Show(this,
                            "⚠ فایل دیتابیس در مسیر انتخابی وجود ندارد.\n\nآیا می‌خواهید ادامه دهید؟",
                            "هشدار", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                        return;
                }

                if (!Directory.Exists(txtPhotosFolder.Text))
                {
                    if (MessageBox.Show(this,
                            "⚠ پوشه عکس‌ها وجود ندارد.\n\nآیا می‌خواهید آن را ایجاد کنید؟",
                            "پرسش", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        Directory.CreateDirectory(txtPhotosFolder.Text);
                    else
                        return;
                }

                AppSettings.DatabasePath = txtDatabasePath.Text;
                AppSettings.PhotosFolder = txtPhotosFolder.Text;

                FontSettings.FontFamilyName = cmbFontFamily.Text;
                FontSettings.TitleFontSize = (float)numTitleSize.Value;
                FontSettings.LabelFontSize = (float)numLabelSize.Value;
                FontSettings.TextBoxFontSize = (float)numTextBoxSize.Value;
                FontSettings.ButtonFontSize = (float)numButtonSize.Value;
                FontSettings.BodyFontSize = (float)numBodySize.Value;
                FontSettings.ChartLabelFontSize = (float)numChartLabelSize.Value;
                FontSettings.TitleFontBold = chkBoldTitle.Checked;
                FontSettings.LabelFontBold = chkBoldLabel.Checked;
                FontSettings.ButtonFontBold = chkBoldButton.Checked;
                FontSettings.ChartLabelFontBold = chkBoldChartLabel.Checked;
                FontSettings.SaveSettings();

                MessageBox.Show(this, "✓ تنظیمات با موفقیت ذخیره شد!\n\n🔄 لطفاً برنامه را مجدداً راه‌اندازی کنید.",
                    "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"خطا در ذخیره تنظیمات:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ResetSettings()
        {
            if (MessageBox.Show(this,
                    "⚠ آیا مطمئن هستید که می‌خواهید تنظیمات را به حالت پیش‌فرض برگردانید؟\n\nتمامی تغییرات از بین خواهد رفت!",
                    "تأیید بازنشانی", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                AppSettings.ResetToDefaults();
                FontSettings.ResetToDefaults();
                LoadCurrentSettings();
                MessageBox.Show(this, "✓ تنظیمات به حالت پیش‌فرض برگشت!\n\n🔄 برای اعمال تغییرات، برنامه را مجدداً راه‌اندازی کنید.",
                    "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
