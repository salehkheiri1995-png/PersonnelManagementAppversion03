using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace PersonnelManagementApp
{
    /// <summary>
    /// پنجره خروجی حرفه‌ای نمودارهای پست‌ها: پیش‌نمایش، ذخیره تصویر و چاپ.
    /// </summary>
    public partial class FormExportPostCharts : Form
    {
        private static readonly Color HeaderColor = Color.FromArgb(15, 23, 42);
        private static readonly Color HeaderAccent = Color.FromArgb(56, 189, 248);
        private static readonly Color CardColor = Color.White;
        private static readonly Color CanvasColor = Color.FromArgb(241, 245, 249);
        private static readonly Color BorderColor = Color.FromArgb(226, 232, 240);
        private static readonly Color PrimaryColor = Color.FromArgb(0, 122, 204);
        private static readonly Color SuccessColor = Color.FromArgb(16, 185, 129);
        private static readonly Color DangerColor = Color.FromArgb(239, 68, 68);
        private static readonly Color TextColor = Color.FromArgb(30, 41, 59);
        private static readonly Color MutedColor = Color.FromArgb(100, 116, 139);

        private readonly PostsDataService _postsService;
        private readonly DataTable? _providedData;
        private readonly string? _initialKey;
        private readonly SeriesChartType? _initialType;

        private DataTable? _data;
        private PostChartSpec? _currentSpec;

        private ComboBox cmbChart = null!;
        private ComboBox cmbChartType = null!;
        private CheckBox chkShowTitle = null!;
        private Chart previewChart = null!;
        private DataGridView dgvStats = null!;
        private Label lblTotal = null!;
        private Label lblStatus = null!;
        private Label lblChartTitle = null!;
        private Button btnExportImage = null!;
        private Button btnPrint = null!;
        private Button btnRefresh = null!;
        private Button btnClose = null!;

        public FormExportPostCharts()
            : this(null, null, null)
        {
        }

        public FormExportPostCharts(DataTable? data, string? initialKey, SeriesChartType? initialType)
        {
            _postsService = new PostsDataService();
            _providedData = data;
            _initialKey = initialKey;
            _initialType = initialType;

            BuildUi();

            FontSettings.ApplyFontToForm(this);
            Load += FormExportPostCharts_Load;
        }

        private sealed class ChartTypeOption
        {
            public ChartTypeOption(SeriesChartType type, string name)
            {
                Type = type;
                Name = name;
            }

            public SeriesChartType Type { get; }
            public string Name { get; }
        }

        private static readonly ChartTypeOption[] ChartTypeOptions =
        {
            new(SeriesChartType.Pie, "دایره‌ای (Pie)"),
            new(SeriesChartType.Doughnut, "حلقه‌ای (Doughnut)"),
            new(SeriesChartType.Column, "ستونی (Column)"),
            new(SeriesChartType.Bar, "میله‌ای افقی (Bar)"),
            new(SeriesChartType.Line, "خطی (Line)"),
            new(SeriesChartType.Area, "سطحی (Area)")
        };

        private void BuildUi()
        {
            SuspendLayout();

            Text = "خروجی نمودارهای پست‌ها";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(980, 620);
            Size = new Size(1240, 780);
            BackColor = CanvasColor;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            Font = new Font(FontSettings.FontFamilyName, 9F);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = CanvasColor,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 92F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48F));
            Controls.Add(root);

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildBody(), 0, 1);
            root.Controls.Add(BuildFooter(), 0, 2);

            ResumeLayout(true);
        }

        private Control BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = HeaderColor,
                Padding = new Padding(24, 14, 24, 12),
                Margin = new Padding(0)
            };

            var title = new Label
            {
                Text = "خروجی حرفه‌ای نمودارهای پست‌ها",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font(FontSettings.FontFamilyName, 17F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };

            var subtitle = new Label
            {
                Text = "پیش‌نمایش زنده، ذخیره تصویر با کیفیت چاپ و چاپ مستقیم نمودار به همراه آمار تفکیکی",
                Dock = DockStyle.Fill,
                Font = new Font(FontSettings.FontFamilyName, 9.5F),
                ForeColor = Color.FromArgb(203, 213, 225),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };

            var accent = new Panel
            {
                Dock = DockStyle.Right,
                Width = 5,
                BackColor = HeaderAccent,
                Margin = new Padding(0, 6, 0, 6)
            };

            header.Controls.Add(subtitle);
            header.Controls.Add(title);
            header.Controls.Add(accent);
            return header;
        }

        private Control BuildBody()
        {
            var body = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = CanvasColor,
                Padding = new Padding(12),
                Margin = new Padding(0)
            };
            // در حالت راست‌به‌چپ، ستون صفر در سمت راست قرار می‌گیرد.
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 292F));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 286F));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            body.Controls.Add(BuildControlsPanel(), 0, 0);
            body.Controls.Add(BuildChartPanel(), 1, 0);
            body.Controls.Add(BuildStatsPanel(), 2, 0);
            return body;
        }

        private Control BuildControlsPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Padding = new Padding(16, 14, 16, 14),
                Margin = new Padding(0, 0, 10, 0),
                BorderStyle = BorderStyle.FixedSingle
            };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                BackColor = CardColor,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            panel.Controls.Add(flow);

            FlowLayoutPanel AddGroup(string text)
            {
                flow.Controls.Add(new Label
                {
                    Text = text,
                    AutoSize = false,
                    Width = 248,
                    Height = 26,
                    Font = new Font(FontSettings.FontFamilyName, 10.5F, FontStyle.Bold),
                    ForeColor = PrimaryColor,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Margin = new Padding(0, 6, 0, 4)
                });
                return flow;
            }

            void AddFieldLabel(string text)
            {
                flow.Controls.Add(new Label
                {
                    Text = text,
                    AutoSize = false,
                    Width = 248,
                    Height = 22,
                    Font = new Font(FontSettings.FontFamilyName, 9F),
                    ForeColor = MutedColor,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Margin = new Padding(0, 6, 0, 2)
                });
            }

            AddGroup("تنظیمات نمودار");

            AddFieldLabel("نمودار مورد نظر");
            cmbChart = new ComboBox
            {
                Width = 248,
                Height = 30,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font(FontSettings.FontFamilyName, 9.5F),
                RightToLeft = RightToLeft.Yes,
                Margin = new Padding(0, 0, 0, 6)
            };
            cmbChart.SelectedIndexChanged += (_, _) => RenderCurrentChart();
            flow.Controls.Add(cmbChart);

            AddFieldLabel("نوع نمایش");
            cmbChartType = new ComboBox
            {
                Width = 248,
                Height = 30,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font(FontSettings.FontFamilyName, 9.5F),
                RightToLeft = RightToLeft.Yes,
                Margin = new Padding(0, 0, 0, 6)
            };
            cmbChartType.Items.AddRange(ChartTypeOptions);
            cmbChartType.DisplayMember = nameof(ChartTypeOption.Name);
            cmbChartType.SelectedIndex = 0;
            cmbChartType.SelectedIndexChanged += (_, _) => RenderCurrentChart();
            flow.Controls.Add(cmbChartType);

            chkShowTitle = new CheckBox
            {
                Text = "نمایش عنوان روی نمودار",
                Width = 248,
                Height = 26,
                Checked = true,
                ForeColor = TextColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 4, 0, 12)
            };
            chkShowTitle.CheckedChanged += (_, _) => RenderCurrentChart();
            flow.Controls.Add(chkShowTitle);

            AddGroup("خروجی و چاپ");

            btnExportImage = CreateActionButton("ذخیره تصویر با کیفیت", SuccessColor);
            btnExportImage.Click += BtnExportImage_Click;
            flow.Controls.Add(btnExportImage);

            btnPrint = CreateActionButton("پیش‌نمایش چاپ", PrimaryColor);
            btnPrint.Click += BtnPrint_Click;
            flow.Controls.Add(btnPrint);

            btnRefresh = CreateActionButton("به‌روزرسانی نمودار", Color.FromArgb(71, 85, 105));
            btnRefresh.Click += (_, _) => RenderCurrentChart();
            flow.Controls.Add(btnRefresh);

            flow.Controls.Add(new Panel { Width = 248, Height = 8, Margin = new Padding(0) });

            btnClose = CreateActionButton("بستن", DangerColor);
            btnClose.Click += (_, _) => Close();
            flow.Controls.Add(btnClose);

            flow.Controls.Add(new Label
            {
                Text = "نمودارها با احتساب پست‌های دارای اطلاعات ناقص (نامشخص) ساخته می‌شوند.",
                AutoSize = false,
                Width = 248,
                Height = 54,
                Font = new Font(FontSettings.FontFamilyName, 8.5F),
                ForeColor = MutedColor,
                TextAlign = ContentAlignment.TopLeft,
                Margin = new Padding(0, 10, 0, 0)
            });

            return panel;
        }

        private static Button CreateActionButton(string text, Color backColor)
        {
            var button = new Button
            {
                Text = text,
                Width = 248,
                Height = 40,
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(FontSettings.FontFamilyName, 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 0, 8),
                TextAlign = ContentAlignment.MiddleCenter
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private Control BuildChartPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Padding = new Padding(14, 12, 14, 14),
                Margin = new Padding(0, 0, 10, 0),
                BorderStyle = BorderStyle.FixedSingle
            };

            lblChartTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 32,
                Font = new Font(FontSettings.FontFamilyName, 12F, FontStyle.Bold),
                ForeColor = TextColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "پیش‌نمایش نمودار",
                AutoEllipsis = true
            };

            previewChart = new Chart
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Margin = new Padding(0),
                BorderlineColor = BorderColor,
                BorderlineWidth = 1,
                MinimumSize = new Size(240, 200)
            };
            previewChart.ChartAreas.Add(new ChartArea("MainArea"));

            panel.Controls.Add(previewChart);
            panel.Controls.Add(lblChartTitle);
            return panel;
        }

        private Control BuildStatsPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Padding = new Padding(14, 12, 14, 12),
                Margin = new Padding(0),
                BorderStyle = BorderStyle.FixedSingle
            };

            lblTotal = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                Font = new Font(FontSettings.FontFamilyName, 11F, FontStyle.Bold),
                ForeColor = SuccessColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "تعداد کل: 0 پست"
            };

            var statsTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 26,
                Font = new Font(FontSettings.FontFamilyName, 10F, FontStyle.Bold),
                ForeColor = MutedColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "آمار تفکیکی"
            };

            dgvStats = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RightToLeft = RightToLeft.Yes,
                GridColor = BorderColor,
                EnableHeadersVisualStyles = false
            };
            dgvStats.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);
            dgvStats.ColumnHeadersDefaultCellStyle.ForeColor = TextColor;
            dgvStats.ColumnHeadersDefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F, FontStyle.Bold);
            dgvStats.DefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F);
            dgvStats.DefaultCellStyle.ForeColor = TextColor;
            dgvStats.DefaultCellStyle.SelectionBackColor = Color.FromArgb(224, 242, 254);
            dgvStats.DefaultCellStyle.SelectionForeColor = TextColor;
            dgvStats.Columns.Add("Label", "دسته");
            dgvStats.Columns.Add("Count", "تعداد");
            dgvStats.Columns.Add("Percent", "درصد");
            dgvStats.Columns["Label"]!.FillWeight = 56;
            dgvStats.Columns["Count"]!.FillWeight = 22;
            dgvStats.Columns["Percent"]!.FillWeight = 22;

            var note = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 34,
                Font = new Font(FontSettings.FontFamilyName, 8.5F),
                ForeColor = MutedColor,
                TextAlign = ContentAlignment.TopLeft,
                Text = "برای مشاهده جزئیات بیشتر، از بخش تحلیل پست‌ها استفاده کنید."
            };

            panel.Controls.Add(dgvStats);
            panel.Controls.Add(statsTitle);
            panel.Controls.Add(lblTotal);
            panel.Controls.Add(note);
            return panel;
        }

        private Control BuildFooter()
        {
            var footer = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(20, 8, 20, 8),
                Margin = new Padding(0)
            };

            lblStatus = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                ForeColor = MutedColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "در حال آماده‌سازی..."
            };

            footer.Controls.Add(lblStatus);
            return footer;
        }

        private void FormExportPostCharts_Load(object? sender, EventArgs e)
        {
            LoadData();
            PopulateChartList();
            SelectInitialChart();
            PopulateChartType();
            RenderCurrentChart();
        }

        private void LoadData()
        {
            if (_providedData != null && _providedData.Rows.Count > 0)
            {
                _data = _providedData;
                return;
            }

            try
            {
                _data = _postsService.LoadAllPosts(showError: true);
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormExportPostCharts.LoadData", ex);
                _data = null;
            }

            if (_data == null || _data.Rows.Count == 0)
            {
                MessageBox.Show(this,
                    "داده‌ای برای نمایش وجود ندارد. ابتدا اتصال پایگاه داده را در تنظیمات بررسی کنید.",
                    "بدون داده", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void PopulateChartList()
        {
            cmbChart.Items.Clear();
            foreach (PostChartSpec spec in PostChartCatalog.All)
                cmbChart.Items.Add(spec);

            cmbChart.DisplayMember = nameof(PostChartSpec.Title);
            if (cmbChart.Items.Count > 0)
                cmbChart.SelectedIndex = 0;
        }

        private void SelectInitialChart()
        {
            if (string.IsNullOrWhiteSpace(_initialKey))
                return;

            for (int i = 0; i < cmbChart.Items.Count; i++)
            {
                if (cmbChart.Items[i] is PostChartSpec spec &&
                    string.Equals(spec.Key, _initialKey, StringComparison.OrdinalIgnoreCase))
                {
                    cmbChart.SelectedIndex = i;
                    return;
                }
            }
        }

        private void PopulateChartType()
        {
            SeriesChartType target = _initialType ?? SeriesChartType.Pie;
            for (int i = 0; i < ChartTypeOptions.Length; i++)
            {
                if (ChartTypeOptions[i].Type == target)
                {
                    cmbChartType.SelectedIndex = i;
                    return;
                }
            }

            cmbChartType.SelectedIndex = 0;
        }

        private SeriesChartType SelectedChartType
        {
            get
            {
                if (cmbChartType.SelectedItem is ChartTypeOption option)
                    return option.Type;
                return SeriesChartType.Pie;
            }
        }

        private void RenderCurrentChart()
        {
            if (_data == null || _data.Rows.Count == 0)
                return;

            if (cmbChart.SelectedItem is not PostChartSpec spec)
                return;

            try
            {
                _currentSpec = spec;
                PostChartRenderer.Render(previewChart, _data, spec, SelectedChartType, chkShowTitle.Checked);

                lblChartTitle.Text = spec.Title;
                PopulateStats(spec);

                int total = PostChartRenderer.BuildStats(_data, spec).Sum(stat => stat.Count);
                lblStatus.Text = $"{total} پست در {PostChartCatalog.All.Count} شاخص قابل تحلیل • نمودار جاری: {spec.Title}";
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormExportPostCharts.RenderCurrentChart", ex);
                lblStatus.Text = "خطا در ساخت نمودار: " + ex.Message;
            }
        }

        private void PopulateStats(PostChartSpec spec)
        {
            List<PostStat> stats = PostChartRenderer.BuildStats(_data!, spec);
            int total = stats.Sum(stat => stat.Count);

            lblTotal.Text = $"تعداد کل: {total} پست";
            dgvStats.Rows.Clear();

            foreach (PostStat stat in stats)
            {
                double percent = total > 0 ? stat.Count * 100.0 / total : 0;
                int index = dgvStats.Rows.Add(stat.Label, stat.Count.ToString("N0"), $"{percent:F1}%");
                if (stat.IsOther)
                    dgvStats.Rows[index].DefaultCellStyle.ForeColor = MutedColor;
            }
        }

        private void BtnExportImage_Click(object? sender, EventArgs e)
        {
            if (!EnsureReady())
                return;

            PostChartImageExporter.Save(previewChart, _data!, _currentSpec!, this);
        }

        private void BtnPrint_Click(object? sender, EventArgs e)
        {
            if (!EnsureReady())
                return;

            PostChartImageExporter.PrintPreview(previewChart, _data!, _currentSpec!, this);
        }

        private bool EnsureReady()
        {
            if (_data == null || _data.Rows.Count == 0)
            {
                MessageBox.Show(this, "داده‌ای برای خروجی وجود ندارد.", "اطلاع",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            if (_currentSpec == null)
            {
                MessageBox.Show(this, "لطفاً ابتدا یک نمودار انتخاب کنید.", "اطلاع",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            return true;
        }
    }
}
