using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Linq;
using System.Drawing;

namespace PersonnelManagementApp
{
    public partial class FormPostsAnalytics : Form
    {
        private readonly DbHelper dbHelper;
        private readonly TabControl tabControl;

        // ====== Charts ======
        private readonly Chart chartProvincePie;
        private readonly Chart chartDeptPie;
        private readonly Chart chartVoltagePie;
        private readonly Chart chartTypePie;
        private readonly Chart chartStandardPie;
        private readonly Chart chartCircuitPie;
        private readonly Chart chartFMPie;
        private readonly Chart chartConnPie;
        private readonly Chart chartInsPie;
        private readonly Chart chartPT2Pie;
        private readonly Chart chartDieselPie;
        private readonly Chart chartOperationYearPie;

        // ====== DataGridViews ======
        private readonly DataGridView dgvPostStats;
        private readonly DataGridView dgvDeptDetails;

        // ====== Filter CheckedListBoxes ======
        private readonly CheckedListBox clbProvincesFilter;
        private readonly CheckedListBox clbCitiesFilter;
        private readonly CheckedListBox clbAffairsFilter;
        private readonly CheckedListBox clbDepartmentsFilter;
        private readonly CheckedListBox clbDistrictsFilter;
        private readonly CheckedListBox clbVoltageFilter;
        private readonly CheckedListBox clbTypeFilter;
        private readonly CheckedListBox clbStandardFilter;
        private readonly CheckedListBox clbCircuitFilter;
        private readonly CheckedListBox clbFMFilter;
        private readonly CheckedListBox clbDieselFilter;
        private readonly CheckedListBox clbWaterFilter;

        private readonly Button btnClearFilters;
        private readonly Label lblFilterInfo;
        private ContextMenuStrip? chartTypeMenu;

        // ====== Cached Data ======
        private DataTable? allPostsData;

        // ====== Cascading Filter Guard ======
        // Prevents recursive event loops when updating dependent filter options
        private bool _updatingFilters = false;

        public FormPostsAnalytics()
        {
            dbHelper = new DbHelper();
            tabControl = new TabControl();

            chartProvincePie      = new Chart();
            chartDeptPie          = new Chart();
            chartVoltagePie       = new Chart();
            chartTypePie          = new Chart();
            chartStandardPie      = new Chart();
            chartCircuitPie       = new Chart();
            chartFMPie            = new Chart();
            chartConnPie          = new Chart();
            chartInsPie           = new Chart();
            chartPT2Pie           = new Chart();
            chartDieselPie        = new Chart();
            chartOperationYearPie = new Chart();

            dgvPostStats   = new DataGridView();
            dgvDeptDetails = new DataGridView();

            clbProvincesFilter   = new CheckedListBox();
            clbCitiesFilter      = new CheckedListBox();
            clbAffairsFilter     = new CheckedListBox();
            clbDepartmentsFilter = new CheckedListBox();
            clbDistrictsFilter   = new CheckedListBox();
            clbVoltageFilter     = new CheckedListBox();
            clbTypeFilter        = new CheckedListBox();
            clbStandardFilter    = new CheckedListBox();
            clbCircuitFilter     = new CheckedListBox();
            clbFMFilter          = new CheckedListBox();
            clbDieselFilter      = new CheckedListBox();
            clbWaterFilter       = new CheckedListBox();

            btnClearFilters = new Button();
            lblFilterInfo   = new Label();

            InitializeComponent();
            BuildUI();
            InitializeChartTypeMenu();
            FontSettings.ApplyFontToForm(this);
            LoadData();
        }

        // ╔══════════════════════════════════════╗
        // ║           BUILD UI                   ║
        // ╚══════════════════════════════════════╝
        private void BuildUI()
        {
            Text = "⚡ تحلیل داده‌های پست‌ها - سیستم پیشرفته";
            WindowState = FormWindowState.Maximized;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            BackColor = Color.FromArgb(240, 248, 255);
            MinimumSize = new Size(1200, 700);
            Font = FontSettings.BodyFont;

            // ── Filter Panel ──────────────────────────────
            Panel panelFilter = new Panel
            {
                Dock = DockStyle.Top,
                Height = 260,
                BackColor = Color.FromArgb(225, 240, 255),
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = true,
                Padding = new Padding(6, 6, 6, 4)
            };

            TableLayoutPanel filterGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 210,
                ColumnCount = 6,
                RowCount = 2,
                RightToLeft = RightToLeft.Yes,
                BackColor = Color.Transparent,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            for (int i = 0; i < 6; i++)
                filterGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16.66f));
            filterGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
            filterGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));

            // Row 0
            filterGrid.Controls.Add(CreateFilterBox("استانها 🗺️",     clbProvincesFilter,   ClbProvincesFilter_ItemCheck), 0, 0);
            filterGrid.Controls.Add(CreateFilterBox("شهرها 🏙️",       clbCitiesFilter,      ClbCitiesFilter_ItemCheck),    1, 0);
            filterGrid.Controls.Add(CreateFilterBox("امور 📋",         clbAffairsFilter,     ClbAffairsFilter_ItemCheck),   2, 0);
            filterGrid.Controls.Add(CreateFilterBox("ادارات 🏛️",      clbDepartmentsFilter, ClbDeptFilter_ItemCheck),      3, 0);
            filterGrid.Controls.Add(CreateFilterBox("نواحی 🔺",        clbDistrictsFilter,   ClbDistrictFilter_ItemCheck),  4, 0);
            filterGrid.Controls.Add(CreateFilterBox("سطح ولتاژ ⚡",    clbVoltageFilter,     ClbVoltageFilter_ItemCheck),   5, 0);
            // Row 1
            filterGrid.Controls.Add(CreateFilterBox("نوع پست 🏗️",     clbTypeFilter,        ClbTypeFilter_ItemCheck),      0, 1);
            filterGrid.Controls.Add(CreateFilterBox("استاندارد 📐",    clbStandardFilter,    ClbStandardFilter_ItemCheck),  1, 1);
            filterGrid.Controls.Add(CreateFilterBox("وضعیت مدار 🔌",   clbCircuitFilter,     ClbCircuitFilter_ItemCheck),   2, 1);
            filterGrid.Controls.Add(CreateFilterBox("ثابت / سیار 🚗",  clbFMFilter,          ClbFMFilter_ItemCheck),        3, 1);
            filterGrid.Controls.Add(CreateFilterBox("دیزل ژنراتور 🔋", clbDieselFilter,      ClbDieselFilter_ItemCheck),    4, 1);
            filterGrid.Controls.Add(CreateFilterBox("وضعیت آب 💧",     clbWaterFilter,       ClbWaterFilter_ItemCheck),     5, 1);

            panelFilter.Controls.Add(filterGrid);

            // Bottom row of filter panel
            Panel filterBottomPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 42,
                BackColor = Color.Transparent
            };

            FlowLayoutPanel rowActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = true,
                BackColor = Color.Transparent,
                Padding = new Padding(0, 4, 0, 0)
            };

            btnClearFilters.Text = "🔄 غیرفعال کردن فیلترها";
            btnClearFilters.Size = new Size(190, 34);
            btnClearFilters.BackColor = Color.FromArgb(220, 53, 69);
            btnClearFilters.ForeColor = Color.White;
            btnClearFilters.Font = new Font(FontSettings.ButtonFont.FontFamily, 9.5F, FontStyle.Bold);
            btnClearFilters.FlatStyle = FlatStyle.Flat;
            btnClearFilters.FlatAppearance.BorderSize = 0;
            btnClearFilters.Margin = new Padding(4, 2, 4, 2);
            btnClearFilters.Click += BtnClearFilters_Click;

            rowActions.Controls.Add(btnClearFilters);
            filterBottomPanel.Controls.Add(rowActions);

            lblFilterInfo.Text = "✓ فیلتری فعال نیست";
            lblFilterInfo.Dock = DockStyle.Bottom;
            lblFilterInfo.Height = 28;
            lblFilterInfo.Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold);
            lblFilterInfo.ForeColor = Color.FromArgb(0, 102, 204);
            lblFilterInfo.TextAlign = ContentAlignment.MiddleLeft;

            panelFilter.Controls.Add(filterBottomPanel);
            panelFilter.Controls.Add(lblFilterInfo);

            // ── Main 66/34 split ──────────────────────────
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                RightToLeft = RightToLeft.No,
                BackColor = Color.Transparent
            };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 66f));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34f));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            // Charts panel
            Panel chartsPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(4) };

            tabControl.Dock = DockStyle.Fill;
            tabControl.RightToLeft = RightToLeft.Yes;
            tabControl.RightToLeftLayout = true;
            tabControl.ItemSize = new Size(130, 30);
            tabControl.Font = FontSettings.BodyFont;

            AddChartTab(tabControl, "🗺️ استان",          chartProvincePie);
            AddChartTab(tabControl, "🏛️ ادارات",         chartDeptPie);
            AddChartTab(tabControl, "⚡ ولتاژ",           chartVoltagePie);
            AddChartTab(tabControl, "🏗️ نوع پست",        chartTypePie);
            AddChartTab(tabControl, "📐 استاندارد",       chartStandardPie);
            AddChartTab(tabControl, "🔌 مدار",            chartCircuitPie);
            AddChartTab(tabControl, "🚗 ثابت/سیار",       chartFMPie);
            AddChartTab(tabControl, "🔗 اتصال توزیع",     chartConnPie);
            AddChartTab(tabControl, "🔆 نوع عایق",        chartInsPie);
            AddChartTab(tabControl, "📋 نوع پست ۲",       chartPT2Pie);
            AddChartTab(tabControl, "🔋 دیزل",            chartDieselPie);
            AddChartTab(tabControl, "📅 سال بهره‌برداری", chartOperationYearPie);

            chartsPanel.Controls.Add(tabControl);

            // Tables panel
            Panel tablesPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(4) };

            TabControl tablesTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true,
                Font = FontSettings.BodyFont,
                ItemSize = new Size(145, 30)
            };

            TabPage tabStats = new TabPage("📋 آمار کلی") { Padding = new Padding(0) };
            dgvPostStats.Dock = DockStyle.Fill;
            ConfigureDgv(dgvPostStats);
            tabStats.Controls.Add(dgvPostStats);

            TabPage tabDept = new TabPage("🏛️ جزئیات ادارات") { Padding = new Padding(0) };
            dgvDeptDetails.Dock = DockStyle.Fill;
            ConfigureDgv(dgvDeptDetails);
            tabDept.Controls.Add(dgvDeptDetails);

            tablesTabControl.TabPages.Add(tabStats);
            tablesTabControl.TabPages.Add(tabDept);
            tablesPanel.Controls.Add(tablesTabControl);

            mainLayout.Controls.Add(chartsPanel, 0, 0);
            mainLayout.Controls.Add(tablesPanel, 1, 0);

            Controls.Add(mainLayout);
            Controls.Add(panelFilter);
        }

        // ╔══════════════════════════════════════╗
        // ║         HELPERS                      ║
        // ╚══════════════════════════════════════╝
        private void ConfigureDgv(DataGridView dgv)
        {
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgv.ReadOnly = true;
            dgv.RightToLeft = RightToLeft.Yes;
            dgv.BackgroundColor = Color.White;
            dgv.EnableHeadersVisualStyles = false;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
            dgv.ColumnHeadersHeight = 35;
            dgv.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
        }

        private Panel CreateFilterBox(string title, CheckedListBox clb, ItemCheckEventHandler handler)
        {
            Panel box = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(2),
                Padding = new Padding(3),
                BackColor = Color.FromArgb(245, 252, 255),
                BorderStyle = BorderStyle.FixedSingle
            };
            Label lbl = new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 20,
                Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                TextAlign = ContentAlignment.MiddleLeft
            };
            clb.RightToLeft = RightToLeft.Yes;
            clb.CheckOnClick = true;
            clb.BackColor = Color.White;
            clb.Font = new Font(FontSettings.BodyFont.FontFamily, 9F);
            clb.IntegralHeight = false;
            clb.BorderStyle = BorderStyle.FixedSingle;
            clb.HorizontalScrollbar = false;
            clb.ScrollAlwaysVisible = false;
            clb.ItemHeight = 18;
            clb.Dock = DockStyle.Fill;
            clb.ItemCheck -= handler;
            clb.ItemCheck += handler;
            box.Controls.Add(clb);
            box.Controls.Add(lbl);
            return box;
        }

        private void AddChartTab(TabControl tc, string title, Chart chart)
        {
            TabPage tab = new TabPage(title);

            Button btnExport = new Button
            {
                Text = "📤",
                Size = new Size(40, 35),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                Font = new Font("Segoe UI Emoji", 11F),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnExport.FlatAppearance.BorderSize = 0;
            btnExport.Location = new Point(10, 10);
            btnExport.Click += (s, e) => new FormExportCharts().ShowDialog();

            chart.Dock = DockStyle.Fill;
            chart.BackColor = Color.White;
            chart.MinimumSize = new Size(100, 100);
            chart.ChartAreas.Add(new ChartArea("ChartArea1")
            {
                BackColor = Color.White,
                Area3DStyle = { Enable3D = true, Inclination = 15, Rotation = 45 }
            });
            chart.MouseClick += Chart_MouseClick;

            tab.Controls.Add(chart);
            tab.Controls.Add(btnExport);
            btnExport.BringToFront();
            tc.TabPages.Add(tab);
        }

        // ╔══════════════════════════════════════╗
        // ║        CHART TYPE MENU               ║
        // ╚══════════════════════════════════════╝
        private void InitializeChartTypeMenu()
        {
            chartTypeMenu = new ContextMenuStrip { RightToLeft = RightToLeft.Yes, ShowImageMargin = false };

            void Add(string text, SeriesChartType type)
                => chartTypeMenu.Items.Add(new ToolStripMenuItem(text) { Tag = type });

            Add("نمودار دایره‌ای (Pie)",         SeriesChartType.Pie);
            Add("نمودار حلقه‌ای (Doughnut)",      SeriesChartType.Doughnut);
            Add("نمودار میله‌ای افقی (Bar)",       SeriesChartType.Bar);
            Add("نمودار ستونی عمودی (Column)",     SeriesChartType.Column);
            Add("نمودار راداری (Radar)",           SeriesChartType.Radar);
            Add("نمودار قطبی (Polar)",             SeriesChartType.Polar);
            Add("میله‌ای انباشته (StackedBar)",    SeriesChartType.StackedBar);
            Add("ستونی انباشته (StackedColumn)",   SeriesChartType.StackedColumn);

            chartTypeMenu.ItemClicked += ChartTypeMenu_ItemClicked;

            foreach (var c in AllCharts())
                c.ContextMenuStrip = chartTypeMenu;
        }

        private IEnumerable<Chart> AllCharts() => new[]
        {
            chartProvincePie, chartDeptPie, chartVoltagePie, chartTypePie, chartStandardPie,
            chartCircuitPie, chartFMPie, chartConnPie, chartInsPie, chartPT2Pie,
            chartDieselPie, chartOperationYearPie
        };

        private void ChartTypeMenu_ItemClicked(object? sender, ToolStripItemClickedEventArgs e)
        {
            chartTypeMenu?.Hide();
            if (e.ClickedItem?.Tag is SeriesChartType type)
            {
                var chart = (sender as ContextMenuStrip)?.SourceControl as Chart;
                if (chart != null) { chart.Tag = type; ApplyChartTypeToChart(chart, type); }
            }
        }

        private static bool IsPieType(SeriesChartType t)
            => t == SeriesChartType.Pie || t == SeriesChartType.Doughnut;

        private SeriesChartType GetChartTypeOrDefault(Chart c)
            => c?.Tag is SeriesChartType t ? t : SeriesChartType.Pie;

        private void ConfigureChartAreaForType(Chart chart, SeriesChartType type)
        {
            if (chart == null || chart.ChartAreas.Count == 0) return;
            var area = chart.ChartAreas[0];
            bool pie = IsPieType(type);
            area.Area3DStyle.Enable3D = pie;
            if (pie) { area.Area3DStyle.Inclination = 15; area.Area3DStyle.Rotation = 45; }
            area.AxisX.Enabled = pie ? AxisEnabled.False : AxisEnabled.True;
            area.AxisY.Enabled = pie ? AxisEnabled.False : AxisEnabled.True;
            area.AxisX.MajorGrid.Enabled = !pie;
            area.AxisY.MajorGrid.Enabled = !pie;
            area.AxisX.MajorGrid.LineColor = Color.Gainsboro;
            area.AxisY.MajorGrid.LineColor = Color.Gainsboro;
            if (!pie)
            {
                if (type == SeriesChartType.Bar || type == SeriesChartType.StackedBar)
                    { area.AxisY.Interval = 1; area.AxisX.LabelStyle.Angle = 0; }
                else
                    { area.AxisX.Interval = 1; area.AxisX.LabelStyle.Angle = -45; }
            }
        }

        private void ConfigureSeriesForType(Series series, SeriesChartType type)
        {
            if (series == null) return;
            series.ChartType = type;
            series.XValueType = ChartValueType.String;
            series.YValueType = ChartValueType.Int32;
            series.IsXValueIndexed = true;
            series.Font = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
            series.IsValueShownAsLabel = true;
            if (IsPieType(type)) series["PieLabelStyle"] = "Outside";
        }

        private void ApplyChartTypeToChart(Chart chart, SeriesChartType type)
        {
            if (chart == null) return;
            ConfigureChartAreaForType(chart, type);
            foreach (Series s in chart.Series) ConfigureSeriesForType(s, type);
        }

        // ╔══════════════════════════════════════╗
        // ║           DATA LOADING               ║
        // ╚══════════════════════════════════════╝
        private void LoadData()
        {
            try
            {
                if (!dbHelper.TestConnection())
                {
                    MessageBox.Show("❌ اتصال به دیتابیس ناموفق بود.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // MS Access (ACE.OLEDB) requires nested parentheses for multiple INNER JOINs
                string query = @"
                    SELECT Posts.PostID, Posts.OperationYear, Posts.DistributedCapacity, 
                    Posts.CapacityHV, Posts.CapacityMV, 
                    Provinces.ProvinceName, Cities.CityName, TransferAffairs.AffairName, 
                    OperationDepartments.DeptName, Districts.DistrictName, PostsNames.PostName, 
                    VoltageLevels.VoltageName, PostStandards.StandardName, PostTypes.TypeName, 
                    DistributedConnections.ConnName, InsulationTypes.InsName, PostTypeTwos.PT2Name, 
                    FixedMobiles.FMName, CircuitStatuses.CircuitName, DieselGenerators.DieselName, 
                    DistributionFeeds.FeedName, WaterStatuses.WaterName, GuestHouses.GuestName 
                    FROM (((((((((((((((((Posts 
                    INNER JOIN Provinces ON Posts.ProvinceID = Provinces.ProvinceID)
                    INNER JOIN Cities ON Posts.CityID = Cities.CityID)
                    INNER JOIN TransferAffairs ON Posts.AffairID = TransferAffairs.AffairID)
                    INNER JOIN OperationDepartments ON Posts.DeptID = OperationDepartments.DeptID)
                    INNER JOIN Districts ON Posts.DistrictID = Districts.DistrictID)
                    INNER JOIN PostsNames ON Posts.PostNameID = PostsNames.PostNameID)
                    INNER JOIN VoltageLevels ON Posts.VoltageID = VoltageLevels.VoltageID)
                    INNER JOIN PostStandards ON Posts.StandardID = PostStandards.StandardID)
                    INNER JOIN PostTypes ON Posts.TypeID = PostTypes.TypeID)
                    INNER JOIN DistributedConnections ON Posts.ConnID = DistributedConnections.ConnID)
                    INNER JOIN InsulationTypes ON Posts.InsID = InsulationTypes.InsID)
                    INNER JOIN PostTypeTwos ON Posts.PT2ID = PostTypeTwos.PT2ID)
                    INNER JOIN FixedMobiles ON Posts.FMID = FixedMobiles.FMID)
                    INNER JOIN CircuitStatuses ON Posts.CircuitID = CircuitStatuses.CircuitID)
                    INNER JOIN DieselGenerators ON Posts.DieselID = DieselGenerators.DieselID)
                    INNER JOIN DistributionFeeds ON Posts.FeedID = DistributionFeeds.FeedID)
                    INNER JOIN WaterStatuses ON Posts.WaterID = WaterStatuses.WaterID)
                    INNER JOIN GuestHouses ON Posts.GuestID = GuestHouses.GuestID
                ";

                allPostsData = dbHelper.ExecuteQuery(query);

                if (allPostsData == null || allPostsData.Rows.Count == 0)
                {
                    MessageBox.Show("⚠️ داده‌ای در جدول پست‌ها یافت نشد.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                LoadFilterOptions();
                RefreshAllCharts();
                MessageBox.Show(
                    $"✅ داده‌ها با موفقیت بارگذاری شدند.\n⚡ تعداد پست‌ها: {allPostsData.Rows.Count}",
                    "موفق", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadFilterOptions()
        {
            if (allPostsData == null) return;
            // Reload ALL options for every filter (used on initial load and after clear)
            FillFilter(clbProvincesFilter,   "ProvinceName");
            FillFilter(clbCitiesFilter,      "CityName");
            FillFilter(clbAffairsFilter,     "AffairName");
            FillFilter(clbDepartmentsFilter, "DeptName");
            FillFilter(clbDistrictsFilter,   "DistrictName");
            FillFilter(clbVoltageFilter,     "VoltageName");
            FillFilter(clbTypeFilter,        "TypeName");
            FillFilter(clbStandardFilter,    "StandardName");
            FillFilter(clbCircuitFilter,     "CircuitName");
            FillFilter(clbFMFilter,          "FMName");
            FillFilter(clbDieselFilter,      "DieselName");
            FillFilter(clbWaterFilter,       "WaterName");
        }

        private void FillFilter(CheckedListBox clb, string col)
        {
            clb.Items.Clear();
            if (allPostsData == null) return;
            foreach (var v in allPostsData.AsEnumerable()
                .Select(r => r[col]?.ToString() ?? "")
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct().OrderBy(v => v))
                clb.Items.Add(v, false);
        }

        // ── Cascading Filter Helpers ────────────────────
        // Refreshes one CLB's items based on a filtered row set,
        // preserving currently-checked items that are still valid.
        private void RefreshClb(CheckedListBox clb, IEnumerable<DataRow> rows, string col)
        {
            var kept = clb.CheckedItems.Cast<string>().ToList();
            var opts = rows
                .Select(r => r[col]?.ToString() ?? "")
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct().OrderBy(v => v).ToList();
            clb.Items.Clear();
            foreach (var item in opts)
                clb.Items.Add(item, kept.Contains(item));
        }

        // Cascade chain:
        //   Province  -->  Cities, Affairs
        //   Province + Cities + Affairs  -->  Departments
        //   Province + Cities + Affairs + Departments  -->  Districts
        private void UpdateDependentFilters()
        {
            if (allPostsData == null || _updatingFilters) return;
            _updatingFilters = true;
            try
            {
                var allRows = allPostsData.AsEnumerable();

                // --- Level 1: Province narrows Cities and Affairs ---
                var selProv = clbProvincesFilter.CheckedItems.Cast<string>().ToList();
                var afterProv = selProv.Count > 0
                    ? allRows.Where(r => selProv.Contains(r["ProvinceName"]?.ToString() ?? ""))
                    : allRows;

                RefreshClb(clbCitiesFilter,  afterProv, "CityName");
                RefreshClb(clbAffairsFilter, afterProv, "AffairName");

                // --- Level 2: Province + Cities + Affairs narrows Departments ---
                var selCity = clbCitiesFilter.CheckedItems.Cast<string>().ToList();
                var selAff  = clbAffairsFilter.CheckedItems.Cast<string>().ToList();
                var afterAff = afterProv;
                if (selCity.Count > 0)
                    afterAff = afterAff.Where(r => selCity.Contains(r["CityName"]?.ToString()  ?? ""));
                if (selAff.Count > 0)
                    afterAff = afterAff.Where(r => selAff.Contains(r["AffairName"]?.ToString() ?? ""));

                RefreshClb(clbDepartmentsFilter, afterAff, "DeptName");

                // --- Level 3: + Departments narrows Districts ---
                var selDept  = clbDepartmentsFilter.CheckedItems.Cast<string>().ToList();
                var afterDept = afterAff;
                if (selDept.Count > 0)
                    afterDept = afterDept.Where(r => selDept.Contains(r["DeptName"]?.ToString() ?? ""));

                RefreshClb(clbDistrictsFilter, afterDept, "DistrictName");
            }
            finally
            {
                _updatingFilters = false;
            }
        }

        // ── Filtered DataTable ──────────────────
        private DataTable GetFilteredData()
        {
            if (allPostsData == null) return new DataTable();

            var p   = clbProvincesFilter.CheckedItems.Cast<string>().ToList();
            var ci  = clbCitiesFilter.CheckedItems.Cast<string>().ToList();
            var af  = clbAffairsFilter.CheckedItems.Cast<string>().ToList();
            var d   = clbDepartmentsFilter.CheckedItems.Cast<string>().ToList();
            var di  = clbDistrictsFilter.CheckedItems.Cast<string>().ToList();
            var v   = clbVoltageFilter.CheckedItems.Cast<string>().ToList();
            var ty  = clbTypeFilter.CheckedItems.Cast<string>().ToList();
            var st  = clbStandardFilter.CheckedItems.Cast<string>().ToList();
            var cr  = clbCircuitFilter.CheckedItems.Cast<string>().ToList();
            var fm  = clbFMFilter.CheckedItems.Cast<string>().ToList();
            var di2 = clbDieselFilter.CheckedItems.Cast<string>().ToList();
            var w   = clbWaterFilter.CheckedItems.Cast<string>().ToList();

            string S(DataRow r, string col) => r[col]?.ToString() ?? "";

            var filtered = allPostsData.AsEnumerable().Where(r =>
                (p.Count   == 0 || p.Contains(S(r, "ProvinceName")))   &&
                (ci.Count  == 0 || ci.Contains(S(r, "CityName")))      &&
                (af.Count  == 0 || af.Contains(S(r, "AffairName")))    &&
                (d.Count   == 0 || d.Contains(S(r, "DeptName")))       &&
                (di.Count  == 0 || di.Contains(S(r, "DistrictName")))  &&
                (v.Count   == 0 || v.Contains(S(r, "VoltageName")))    &&
                (ty.Count  == 0 || ty.Contains(S(r, "TypeName")))      &&
                (st.Count  == 0 || st.Contains(S(r, "StandardName")))  &&
                (cr.Count  == 0 || cr.Contains(S(r, "CircuitName")))   &&
                (fm.Count  == 0 || fm.Contains(S(r, "FMName")))        &&
                (di2.Count == 0 || di2.Contains(S(r, "DieselName")))   &&
                (w.Count   == 0 || w.Contains(S(r, "WaterName")))
            ).ToList();

            return filtered.Count > 0 ? filtered.CopyToDataTable() : new DataTable();
        }

        private List<(string Name, int Count)> GroupStats(DataTable dt, string col)
        {
            if (dt == null || dt.Rows.Count == 0) return new List<(string Name, int Count)>();
            return dt.AsEnumerable()
                .GroupBy(r => r[col]?.ToString() ?? "نامشخص")
                .Select(g => (Name: g.Key, Count: g.Count()))
                .OrderByDescending(x => x.Count)
                .ToList();
        }

        // ╔══════════════════════════════════════╗
        // ║          REFRESH CHARTS              ║
        // ╚══════════════════════════════════════╝
        private void RefreshAllCharts()
        {
            var dt = GetFilteredData();
            DrawChart(chartProvincePie,      dt, "ProvinceName",  "🗺️ توزیع پست‌ها بر اساس استان");
            DrawChart(chartDeptPie,          dt, "DeptName",      "🏛️ توزیع پست‌ها بر اساس اداره");
            DrawChart(chartVoltagePie,       dt, "VoltageName",   "⚡ توزیع بر اساس سطح ولتاژ");
            DrawChart(chartTypePie,          dt, "TypeName",      "🏗️ توزیع بر اساس نوع پست");
            DrawChart(chartStandardPie,      dt, "StandardName",  "📐 توزیع بر اساس استاندارد پست");
            DrawChart(chartCircuitPie,       dt, "CircuitName",   "🔌 توزیع بر اساس وضعیت مدار");
            DrawChart(chartFMPie,            dt, "FMName",        "🚗 توزیع بر اساس ثابت / سیار");
            DrawChart(chartConnPie,          dt, "ConnName",      "🔗 توزیع بر اساس اتصال توزیع");
            DrawChart(chartInsPie,           dt, "InsName",       "🔆 توزیع بر اساس نوع عایق");
            DrawChart(chartPT2Pie,           dt, "PT2Name",       "📋 توزیع بر اساس نوع پست ۲");
            DrawChart(chartDieselPie,        dt, "DieselName",    "🔋 توزیع بر اساس دیزل ژنراتور");
            DrawOperationYearChart(dt);
            LoadSummaryTable(dt);
            LoadDeptDetailsTable(dt);
        }

        private void DrawChart(Chart chart, DataTable dt, string col, string title)
        {
            try
            {
                chart.Series.Clear();
                Font safeFont  = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font titleFont = FontSettings.HeaderFont     ?? new Font("Tahoma", 11F, FontStyle.Bold);

                var stats = GroupStats(dt, col);
                int total = stats.Sum(x => x.Count);
                var type  = GetChartTypeOrDefault(chart);
                bool pie  = IsPieType(type);
                ConfigureChartAreaForType(chart, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var (name, count) in pie ? stats.Take(15) : (IEnumerable<(string, int)>)stats)
                {
                    double pct = total > 0 ? (count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(name, count);
                    series.Points[idx].AxisLabel = name;
                    series.Points[idx].ToolTip   = $"{name}: {count} پست ({pct:F1}%)";
                    series.Points[idx].Font      = safeFont;
                    series.Points[idx].Label     = pie ? $"{name}\n{count} ({pct:F1}%)" : count.ToString();
                }

                chart.Series.Add(series);
                chart.Titles.Clear();
                chart.Titles.Add(new Title(title) { Font = titleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا در نمودار: {ex.Message}"); }
        }

        private void DrawOperationYearChart(DataTable dt)
        {
            try
            {
                chartOperationYearPie.Series.Clear();
                Font safeFont  = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font titleFont = FontSettings.HeaderFont     ?? new Font("Tahoma", 11F, FontStyle.Bold);

                var stats = dt.AsEnumerable()
                    .Where(r => r["OperationYear"] != DBNull.Value)
                    .GroupBy(r =>
                    {
                        if (int.TryParse(r["OperationYear"]?.ToString(), out int y))
                            return $"{(y / 10) * 10}\u2013{(y / 10) * 10 + 9}";
                        return "نامشخص";
                    })
                    .Select(g => (Name: g.Key, Count: g.Count()))
                    .OrderBy(x => x.Name)
                    .ToList();

                int total = stats.Sum(x => x.Count);
                var type  = GetChartTypeOrDefault(chartOperationYearPie);
                bool pie  = IsPieType(type);
                ConfigureChartAreaForType(chartOperationYearPie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var (name, count) in stats)
                {
                    double pct = total > 0 ? (count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(name, count);
                    series.Points[idx].AxisLabel = name;
                    series.Points[idx].ToolTip   = $"{name}: {count} پست ({pct:F1}%)";
                    series.Points[idx].Font      = safeFont;
                    series.Points[idx].Label     = pie ? $"{name}\n{count} ({pct:F1}%)" : count.ToString();
                }

                chartOperationYearPie.Series.Add(series);
                chartOperationYearPie.Titles.Clear();
                chartOperationYearPie.Titles.Add(new Title("📅 توزیع بر اساس دهه بهره‌برداری") { Font = titleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا در نمودار سال: {ex.Message}"); }
        }

        // ── Statistics Tables ────────────────────
        private void LoadSummaryTable(DataTable dt)
        {
            try
            {
                dgvPostStats.DataSource = null;
                dgvPostStats.Columns.Clear();
                dgvPostStats.Columns.Add("Metric", "معیار");
                dgvPostStats.Columns.Add("Value",  "مقدار");

                dgvPostStats.Rows.Add("⚡ کل پست‌ها",               dt.Rows.Count);
                dgvPostStats.Rows.Add("🗺️ تعداد استان‌ها",          GroupStats(dt, "ProvinceName").Count);
                dgvPostStats.Rows.Add("🏙️ تعداد شهرها",            GroupStats(dt, "CityName").Count);
                dgvPostStats.Rows.Add("📋 تعداد امور",              GroupStats(dt, "AffairName").Count);
                dgvPostStats.Rows.Add("🏛️ تعداد ادارات",           GroupStats(dt, "DeptName").Count);
                dgvPostStats.Rows.Add("🔺 تعداد نواحی",             GroupStats(dt, "DistrictName").Count);
                dgvPostStats.Rows.Add("⚡ سطوح ولتاژ",              GroupStats(dt, "VoltageName").Count);
                dgvPostStats.Rows.Add("🏗️ انواع پست",              GroupStats(dt, "TypeName").Count);
                dgvPostStats.Rows.Add("📐 تعداد استانداردها",       GroupStats(dt, "StandardName").Count);

                void Section(string header, string col)
                {
                    dgvPostStats.Rows.Add("", "");
                    dgvPostStats.Rows.Add($"════ {header} ════", "");
                    foreach (var (name, count) in GroupStats(dt, col))
                        dgvPostStats.Rows.Add($"  \u2022 {name}", count);
                }

                Section("توزیع ولتاژ",      "VoltageName");
                Section("توزیع نوع پست",    "TypeName");
                Section("توزیع ثابت/سیار",  "FMName");
                Section("توزیع مدار",       "CircuitName");
                Section("توزیع دیزل",       "DieselName");
                Section("توزیع وضعیت آب",  "WaterName");
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadDeptDetailsTable(DataTable dt)
        {
            try
            {
                dgvDeptDetails.DataSource = null;
                dgvDeptDetails.Columns.Clear();
                dgvDeptDetails.Columns.Add("Dept",    "اداره");
                dgvDeptDetails.Columns.Add("Count",   "تعداد پست");
                dgvDeptDetails.Columns.Add("Percent", "درصد");

                var stats = GroupStats(dt, "DeptName");
                int total = stats.Sum(x => x.Count);
                foreach (var (name, count) in stats)
                {
                    double pct = total > 0 ? (count * 100.0) / total : 0;
                    dgvDeptDetails.Rows.Add(name, count, $"{pct:F1}%");
                }
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        // ╔══════════════════════════════════════╗
        // ║           FILTER EVENTS              ║
        // ╚══════════════════════════════════════╝
        private void FilterChanged()
        {
            // Guard: if we are already updating dependent filters, skip to avoid recursion
            if (_updatingFilters) return;

            BeginInvoke((MethodInvoker)delegate
            {
                if (_updatingFilters) return;

                // Step 1: cascade-update dependent filter options
                UpdateDependentFilters();

                // Step 2: count active filters
                int n =
                    clbProvincesFilter.CheckedItems.Count   + clbCitiesFilter.CheckedItems.Count    +
                    clbAffairsFilter.CheckedItems.Count     + clbDepartmentsFilter.CheckedItems.Count +
                    clbDistrictsFilter.CheckedItems.Count   + clbVoltageFilter.CheckedItems.Count   +
                    clbTypeFilter.CheckedItems.Count        + clbStandardFilter.CheckedItems.Count  +
                    clbCircuitFilter.CheckedItems.Count     + clbFMFilter.CheckedItems.Count        +
                    clbDieselFilter.CheckedItems.Count      + clbWaterFilter.CheckedItems.Count;

                lblFilterInfo.Text = n > 0 ? $"🔴 {n} فیلتر فعال" : "✓ فیلتری فعال نیست";

                // Step 3: refresh charts and tables
                RefreshAllCharts();
            });
        }

        private void ClbProvincesFilter_ItemCheck(object? s, ItemCheckEventArgs e)  => FilterChanged();
        private void ClbCitiesFilter_ItemCheck(object? s, ItemCheckEventArgs e)      => FilterChanged();
        private void ClbAffairsFilter_ItemCheck(object? s, ItemCheckEventArgs e)     => FilterChanged();
        private void ClbDeptFilter_ItemCheck(object? s, ItemCheckEventArgs e)        => FilterChanged();
        private void ClbDistrictFilter_ItemCheck(object? s, ItemCheckEventArgs e)    => FilterChanged();
        private void ClbVoltageFilter_ItemCheck(object? s, ItemCheckEventArgs e)     => FilterChanged();
        private void ClbTypeFilter_ItemCheck(object? s, ItemCheckEventArgs e)        => FilterChanged();
        private void ClbStandardFilter_ItemCheck(object? s, ItemCheckEventArgs e)    => FilterChanged();
        private void ClbCircuitFilter_ItemCheck(object? s, ItemCheckEventArgs e)     => FilterChanged();
        private void ClbFMFilter_ItemCheck(object? s, ItemCheckEventArgs e)          => FilterChanged();
        private void ClbDieselFilter_ItemCheck(object? s, ItemCheckEventArgs e)      => FilterChanged();
        private void ClbWaterFilter_ItemCheck(object? s, ItemCheckEventArgs e)       => FilterChanged();

        private void BtnClearFilters_Click(object? sender, EventArgs e)
        {
            // Uncheck all items without triggering cascade loops
            _updatingFilters = true;
            try
            {
                foreach (var clb in new[] {
                    clbProvincesFilter, clbCitiesFilter, clbAffairsFilter, clbDepartmentsFilter,
                    clbDistrictsFilter, clbVoltageFilter, clbTypeFilter, clbStandardFilter,
                    clbCircuitFilter, clbFMFilter, clbDieselFilter, clbWaterFilter })
                    for (int i = 0; i < clb.Items.Count; i++) clb.SetItemChecked(i, false);
            }
            finally
            {
                _updatingFilters = false;
            }

            // Restore ALL options in all filters (cascade may have narrowed them)
            LoadFilterOptions();
            lblFilterInfo.Text = "✓ فیلتری فعال نیست";
            RefreshAllCharts();
        }

        // ╔══════════════════════════════════════╗
        // ║       CHART CLICK → DETAIL FORM      ║
        // ╚══════════════════════════════════════╝
        private void Chart_MouseClick(object? sender, MouseEventArgs e)
        {
            try
            {
                if (sender is not Chart chart) return;
                var result = chart.HitTest(e.X, e.Y);
                if (result.ChartElementType != ChartElementType.DataPoint) return;

                string itemName = result.Series.Points[result.PointIndex].AxisLabel;
                string col = ChartToColumn(chart);
                if (string.IsNullOrEmpty(col)) return;

                var dt = GetFilteredData();
                if (dt.Rows.Count == 0) return;

                var rows = dt.AsEnumerable()
                    .Where(r => r[col]?.ToString() == itemName)
                    .ToList();

                if (rows.Count > 0)
                    ShowPostDetails(itemName, rows.CopyToDataTable());
                else
                    MessageBox.Show("❌ داده‌ای برای نمایش وجود ندارد.", "پیام", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private string ChartToColumn(Chart c)
        {
            if (c == chartProvincePie)  return "ProvinceName";
            if (c == chartDeptPie)      return "DeptName";
            if (c == chartVoltagePie)   return "VoltageName";
            if (c == chartTypePie)      return "TypeName";
            if (c == chartStandardPie)  return "StandardName";
            if (c == chartCircuitPie)   return "CircuitName";
            if (c == chartFMPie)        return "FMName";
            if (c == chartConnPie)      return "ConnName";
            if (c == chartInsPie)       return "InsName";
            if (c == chartPT2Pie)       return "PT2Name";
            if (c == chartDieselPie)    return "DieselName";
            return "";
        }

        private void ShowPostDetails(string category, DataTable postData)
        {
            Form detailsForm = new Form
            {
                Text = $"⚡ جزئیات پست‌ها \u2014 {category}",
                StartPosition = FormStartPosition.CenterScreen,
                WindowState = FormWindowState.Maximized,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true,
                BackColor = Color.FromArgb(240, 248, 255),
                Font = FontSettings.BodyFont
            };

            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                ReadOnly = true,
                RightToLeft = RightToLeft.Yes,
                BackgroundColor = Color.White,
                EnableHeadersVisualStyles = false,
                AllowUserToAddRows = false,
                ColumnHeadersHeight = 40,
                RowTemplate = { Height = 30 }
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 120, 212);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
            dgv.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            string[] cols =
            {
                "PostName|نام پست", "ProvinceName|استان", "CityName|شهر",
                "AffairName|امور", "DeptName|اداره", "DistrictName|ناحیه",
                "VoltageName|سطح ولتاژ", "TypeName|نوع پست", "StandardName|استاندارد",
                "CircuitName|وضعیت مدار", "FMName|ثابت/سیار", "OperationYear|سال بهره‌برداری"
            };
            foreach (var c in cols)
            {
                var parts = c.Split('|');
                dgv.Columns.Add(parts[0], parts[1]);
            }

            foreach (DataRow row in postData.Rows)
                dgv.Rows.Add(cols.Select(c => postData.Columns.Contains(c.Split('|')[0])
                    ? row[c.Split('|')[0]]?.ToString() : "").ToArray());

            Panel bottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 55,
                BackColor = Color.FromArgb(225, 240, 255)
            };
            bottom.Controls.Add(new Label
            {
                Text = $"📊 تعداد: {postData.Rows.Count} پست  |  دسته: {category}",
                AutoSize = true,
                Location = new Point(20, 16),
                Font = new Font(FontSettings.SubtitleFont.FontFamily, 10.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204)
            });

            detailsForm.Controls.Add(dgv);
            detailsForm.Controls.Add(bottom);
            detailsForm.ShowDialog();
        }
    }
}
