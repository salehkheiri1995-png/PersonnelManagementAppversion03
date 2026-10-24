using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using ClosedXML.Excel;

namespace PersonnelManagementApp
{
    public partial class FormPostsAnalytics : Form
    {
        private readonly DbHelper dbHelper;
        private readonly PostsDataService postsService;

        private readonly List<FilterBinding> filters = new();
        private readonly Dictionary<string, Chart> charts = new();
        private readonly Dictionary<string, PostChartSpec> specs = new();

        private readonly TabControl tabControl = new();
        private readonly TabControl tablesTabControl = new();

        private readonly DataGridView dgvSummary = new();
        private readonly DataGridView dgvDeptDetails = new();
        private readonly DataGridView dgvPosts = new();
        private readonly TextBox txtPostSearch = new();

        private readonly RadioButton rbShowSummary = new();
        private readonly RadioButton rbShowFullStats = new();
        private readonly Label lblFilterInfo = new();
        private readonly Button btnClearFilters = new();
        private readonly Button btnExportExcel = new();

        private ContextMenuStrip? chartTypeMenu;

        private DataTable? allPostsData;
        private bool suppressEvents;
        private bool loading;

        // ترتیب تب‌های نمودار
        private static readonly string[] ChartKeys =
        {
            "province", "city", "affair", "department", "district", "postname",
            "voltage", "standard", "type", "circuit", "diesel", "water",
            "operationyear", "capacityhv", "capacitymv"
        };

        public FormPostsAnalytics()
        {
            dbHelper = new DbHelper();
            postsService = new PostsDataService(dbHelper);

            InitializeComponent();
            BuildUI();
            InitializeChartTypeMenu();
            FontSettings.ApplyFontToForm(this);

            LoadData();
        }

        private void BuildUI()
        {
            Text = "🎯 تحلیل داده‌های پست‌ها - سیستم پیشرفته";
            WindowState = FormWindowState.Maximized;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            BackColor = Color.FromArgb(240, 248, 255);
            MinimumSize = new Size(1200, 700);
            Font = FontSettings.BodyFont;

            // ── پنل فیلترها (بالا) ──
            Panel panelFilter = new Panel
            {
                Dock = DockStyle.Top,
                Height = 300,
                BackColor = Color.FromArgb(230, 240, 250),
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = true,
                Padding = new Padding(6, 6, 6, 4)
            };

            TableLayoutPanel filterGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 200,
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

            filterGrid.Controls.Add(CreateFilterBox("استانها 🗺️", "ProvinceName"), 0, 0);
            filterGrid.Controls.Add(CreateFilterBox("شهرها 🏙️", "CityName"), 1, 0);
            filterGrid.Controls.Add(CreateFilterBox("امور 📋", "AffairName"), 2, 0);
            filterGrid.Controls.Add(CreateFilterBox("ادارات 🏛️", "DeptName"), 3, 0);
            filterGrid.Controls.Add(CreateFilterBox("نواحی 🔺", "DistrictName"), 4, 0);
            filterGrid.Controls.Add(CreateFilterBox("پستها ⚡", "PostName"), 5, 0);

            filterGrid.Controls.Add(CreateFilterBox("سطح ولتاژ 🔌", "VoltageName"), 0, 1);
            filterGrid.Controls.Add(CreateFilterBox("استاندارد 📐", "StandardName"), 1, 1);
            filterGrid.Controls.Add(CreateFilterBox("نوع پست 🏗️", "TypeName"), 2, 1);
            filterGrid.Controls.Add(CreateFilterBox("وضعیت مدار 🔄", "CircuitName"), 3, 1);
            filterGrid.Controls.Add(CreateFilterBox("دیزل ⚙️", "DieselName"), 4, 1);
            filterGrid.Controls.Add(CreateFilterBox("وضعیت آب 💧", "WaterName"), 5, 1);

            panelFilter.Controls.Add(filterGrid);

            // نوار پایین فیلترها
            Panel filterBottomPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 82,
                BackColor = Color.Transparent,
                Padding = new Padding(0)
            };

            FlowLayoutPanel rowActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 46,
                RightToLeft = RightToLeft.Yes,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = true,
                BackColor = Color.Transparent,
                Padding = new Padding(0, 2, 0, 0)
            };

            btnClearFilters.Text = "🔄 غیرفعال کردن فیلترها";
            btnClearFilters.Size = new Size(185, 34);
            btnClearFilters.BackColor = Color.FromArgb(220, 53, 69);
            btnClearFilters.ForeColor = Color.White;
            btnClearFilters.Font = new Font(FontSettings.ButtonFont.FontFamily, 9.5F, FontStyle.Bold);
            btnClearFilters.FlatStyle = FlatStyle.Flat;
            btnClearFilters.FlatAppearance.BorderSize = 0;
            btnClearFilters.Margin = new Padding(4, 6, 4, 4);
            btnClearFilters.Click += (_, _) => ClearAllFilters();

            btnExportExcel.Text = "📥 خروجی اکسل پست‌ها";
            btnExportExcel.Size = new Size(210, 34);
            btnExportExcel.BackColor = Color.FromArgb(40, 167, 69);
            btnExportExcel.ForeColor = Color.White;
            btnExportExcel.Font = new Font(FontSettings.ButtonFont.FontFamily, 9.5F, FontStyle.Bold);
            btnExportExcel.FlatStyle = FlatStyle.Flat;
            btnExportExcel.FlatAppearance.BorderSize = 0;
            btnExportExcel.Margin = new Padding(4, 6, 4, 4);
            btnExportExcel.Click += (_, _) => ExportFilteredPostsToExcel();

            rowActions.Controls.Add(btnExportExcel);
            rowActions.Controls.Add(btnClearFilters);

            lblFilterInfo.Text = "✓ فیلتری فعال نیست";
            lblFilterInfo.Dock = DockStyle.Bottom;
            lblFilterInfo.Height = 30;
            lblFilterInfo.Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold);
            lblFilterInfo.ForeColor = Color.FromArgb(0, 102, 204);
            lblFilterInfo.TextAlign = ContentAlignment.MiddleLeft;

            filterBottomPanel.Controls.Add(lblFilterInfo);
            filterBottomPanel.Controls.Add(rowActions);
            panelFilter.Controls.Add(filterBottomPanel);

            // ── چیدمان اصلی ──
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

            Panel chartsPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(4)
            };

            tabControl.Dock = DockStyle.Fill;
            tabControl.RightToLeft = RightToLeft.Yes;
            tabControl.RightToLeftLayout = true;
            tabControl.ItemSize = new Size(120, 30);
            tabControl.Font = FontSettings.BodyFont;

            foreach (string key in ChartKeys)
            {
                PostChartSpec? spec = PostChartCatalog.Find(key);
                if (spec == null)
                    continue;
                specs[key] = spec;
                AddChartTab(key, spec);
            }

            chartsPanel.Controls.Add(tabControl);

            Panel tablesPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(4)
            };

            tablesTabControl.Dock = DockStyle.Fill;
            tablesTabControl.RightToLeft = RightToLeft.Yes;
            tablesTabControl.RightToLeftLayout = true;
            tablesTabControl.Font = FontSettings.BodyFont;
            tablesTabControl.ItemSize = new Size(130, 30);

            tablesTabControl.TabPages.Add(BuildStatsTab());
            tablesTabControl.TabPages.Add(BuildDeptTab());
            tablesTabControl.TabPages.Add(BuildPostsTab());

            tablesPanel.Controls.Add(tablesTabControl);

            mainLayout.Controls.Add(chartsPanel, 0, 0);
            mainLayout.Controls.Add(tablesPanel, 1, 0);

            Controls.Add(mainLayout);
            Controls.Add(panelFilter);
        }

        // ─────────────────────────────────────────────
        // فیلترها
        // ─────────────────────────────────────────────
        private Panel CreateFilterBox(string title, string column)
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

            var clb = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                CheckOnClick = true,
                BackColor = Color.White,
                Font = new Font(FontSettings.BodyFont.FontFamily, 9F),
                IntegralHeight = false,
                BorderStyle = BorderStyle.FixedSingle,
                HorizontalScrollbar = false,
                ScrollAlwaysVisible = false,
                ItemHeight = 18
            };
            clb.ItemCheck += Filter_ItemCheck;

            filters.Add(new FilterBinding(clb, lbl, column, title));
            box.Controls.Add(clb);
            box.Controls.Add(lbl);
            return box;
        }

        private void Filter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            if (suppressEvents || loading)
                return;

            if (!IsHandleCreated || IsDisposed)
                return;

            try
            {
                BeginInvoke((MethodInvoker)ApplyFilterChange);
            }
            catch (InvalidOperationException)
            {
                // فرم در حال بسته شدن است
            }
        }

        private void ApplyFilterChange()
        {
            if (suppressEvents || IsDisposed || loading || allPostsData == null)
                return;

            suppressEvents = true;
            try
            {
                UpdateDependentFilters();
            }
            finally
            {
                suppressEvents = false;
            }

            UpdateFilterTitles();
            DataTable filtered = GetFilteredData();
            RefreshAllCharts();
            UpdateTables(filtered);
        }

        private void ClearAllFilters()
        {
            if (allPostsData == null)
                return;

            suppressEvents = true;
            try
            {
                foreach (FilterBinding binding in filters)
                {
                    for (int i = 0; i < binding.List.Items.Count; i++)
                        binding.List.SetItemChecked(i, false);
                }
            }
            finally
            {
                suppressEvents = false;
            }

            FillAllFilterOptions();
            UpdateFilterTitles();
            DataTable filtered = GetFilteredData();
            RefreshAllCharts();
            UpdateTables(filtered);
        }

        private void UpdateDependentFilters()
        {
            if (allPostsData == null)
                return;

            var allRows = allPostsData.AsEnumerable();
            List<string> provinces = GetChecked(FindFilter("ProvinceName"));

            var afterProvince = provinces.Count == 0
                ? allRows
                : allRows.Where(row => provinces.Contains(Value(row, "ProvinceName")));

            RefreshFilterOptions(FindFilter("CityName"), afterProvince, "CityName");
            RefreshFilterOptions(FindFilter("AffairName"), afterProvince, "AffairName");

            List<string> cities = GetChecked(FindFilter("CityName"));
            List<string> affairs = GetChecked(FindFilter("AffairName"));

            var afterAffair = afterProvince;
            if (cities.Count > 0)
                afterAffair = afterAffair.Where(row => cities.Contains(Value(row, "CityName")));
            if (affairs.Count > 0)
                afterAffair = afterAffair.Where(row => affairs.Contains(Value(row, "AffairName")));

            RefreshFilterOptions(FindFilter("DeptName"), afterAffair, "DeptName");

            List<string> departments = GetChecked(FindFilter("DeptName"));
            var afterDepartment = departments.Count == 0
                ? afterAffair
                : afterAffair.Where(row => departments.Contains(Value(row, "DeptName")));

            RefreshFilterOptions(FindFilter("DistrictName"), afterDepartment, "DistrictName");
        }

        private CheckedListBox FindFilter(string column)
            => filters.First(filter => string.Equals(filter.Column, column, StringComparison.OrdinalIgnoreCase)).List;

        private static List<string> GetChecked(CheckedListBox list)
            => list.CheckedItems.Cast<object>()
                .Select(item => item?.ToString() ?? string.Empty)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();

        private void RefreshFilterOptions(CheckedListBox list, IEnumerable<DataRow> rows, string column)
        {
            List<string> kept = GetChecked(list);
            List<string> options = rows
                .Select(row => Value(row, column))
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(value => value, StringComparer.CurrentCulture)
                .ToList();

            list.BeginUpdate();
            try
            {
                list.Items.Clear();
                foreach (string option in options)
                    list.Items.Add(option, kept.Contains(option, StringComparer.CurrentCultureIgnoreCase));
            }
            finally
            {
                list.EndUpdate();
            }
        }

        private void FillAllFilterOptions()
        {
            if (allPostsData == null)
                return;

            suppressEvents = true;
            try
            {
                foreach (FilterBinding binding in filters)
                {
                    List<string> options = allPostsData.AsEnumerable()
                        .Select(row => Value(row, binding.Column))
                        .Where(value => !string.IsNullOrWhiteSpace(value))
                        .Distinct(StringComparer.CurrentCultureIgnoreCase)
                        .OrderBy(value => value, StringComparer.CurrentCulture)
                        .ToList();

                    binding.List.BeginUpdate();
                    try
                    {
                        binding.List.Items.Clear();
                        foreach (string option in options)
                            binding.List.Items.Add(option, false);
                    }
                    finally
                    {
                        binding.List.EndUpdate();
                    }
                }
            }
            finally
            {
                suppressEvents = false;
            }
        }

        private void UpdateFilterTitles()
        {
            foreach (FilterBinding binding in filters)
            {
                int selected = binding.List.CheckedItems.Count;
                binding.Title.Text = selected > 0
                    ? $"{binding.DisplayName} ({selected})"
                    : binding.DisplayName;

                binding.Title.ForeColor = selected > 0
                    ? Color.FromArgb(220, 53, 69)
                    : Color.FromArgb(0, 102, 204);
            }

            int active = filters.Sum(filter => filter.List.CheckedItems.Count);
            lblFilterInfo.Text = active == 0
                ? "✓ فیلتری فعال نیست"
                : $"🔎 {active} فیلتر فعال است  |  {GetFilteredData().Rows.Count} پست از {allPostsData?.Rows.Count ?? 0}";
        }

        private DataTable GetFilteredData()
        {
            if (allPostsData == null)
                return new DataTable();

            var selected = filters.ToDictionary(
                filter => filter.Column,
                filter => GetChecked(filter.List),
                StringComparer.OrdinalIgnoreCase);

            DataTable result = allPostsData.Clone();
            foreach (DataRow row in allPostsData.Rows)
            {
                bool match = true;
                foreach (var pair in selected)
                {
                    if (pair.Value.Count == 0)
                        continue;

                    if (!pair.Value.Contains(Value(row, pair.Key), StringComparer.CurrentCultureIgnoreCase))
                    {
                        match = false;
                        break;
                    }
                }

                if (match)
                    result.ImportRow(row);
            }

            return result;
        }

        // ─────────────────────────────────────────────
        // نمودارها
        // ─────────────────────────────────────────────
        private void InitializeChartTypeMenu()
        {
            chartTypeMenu = new ContextMenuStrip
            {
                RightToLeft = RightToLeft.Yes,
                ShowImageMargin = false
            };

            AddChartTypeMenuItem("نمودار دایره‌ای (Pie)", SeriesChartType.Pie);
            AddChartTypeMenuItem("نمودار حلقه‌ای (Doughnut)", SeriesChartType.Doughnut);
            AddChartTypeMenuItem("نمودار میله‌ای افقی (Bar)", SeriesChartType.Bar);
            AddChartTypeMenuItem("نمودار ستونی عمودی (Column)", SeriesChartType.Column);
            AddChartTypeMenuItem("نمودار راداری (Radar)", SeriesChartType.Radar);
            AddChartTypeMenuItem("نمودار قطبی (Polar)", SeriesChartType.Polar);
            AddChartTypeMenuItem("میله‌ای انباشته (StackedBar)", SeriesChartType.StackedBar);
            AddChartTypeMenuItem("ستونی انباشته (StackedColumn)", SeriesChartType.StackedColumn);

            chartTypeMenu.ItemClicked += ChartTypeMenu_ItemClicked;
        }

        private void AddChartTypeMenuItem(string text, SeriesChartType type)
        {
            var item = new ToolStripMenuItem(text) { Tag = type };
            if (chartTypeMenu != null)
                chartTypeMenu.Items.Add(item);
        }

        private void ChartTypeMenu_ItemClicked(object? sender, ToolStripItemClickedEventArgs e)
        {
            chartTypeMenu?.Hide();

            if (e.ClickedItem?.Tag is not SeriesChartType type)
                return;

            if ((sender as ContextMenuStrip)?.SourceControl is not Chart chart)
                return;

            chart.Tag = type;
            if (chart.Tag is string key && specs.TryGetValue(key, out PostChartSpec? spec))
            {
                PostChartRenderer.Render(chart, GetFilteredData(), spec, type, true, true);
            }
        }

        private void AddChartTab(string key, PostChartSpec spec)
        {
            TabPage tab = new TabPage(GetTabTitle(key));

            var chart = new Chart
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                MinimumSize = new Size(100, 100),
                Tag = key
            };
            chart.ContextMenuStrip = chartTypeMenu;
            chart.MouseClick += (_, e) => Chart_MouseClick(key, e);

            // دکمه خروجی نمودار جاری
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
            btnExport.Click += (_, _) => ExportThisChart(key, chart);

            tab.Controls.Add(chart);
            tab.Controls.Add(btnExport);
            btnExport.BringToFront();

            charts[key] = chart;
            tabControl.TabPages.Add(tab);
        }

        private static string GetTabTitle(string key)
        {
            return key switch
            {
                "province" => "🗺️ استان",
                "city" => "🏙️ شهر",
                "affair" => "📋 امور",
                "department" => "🏛️ ادارات",
                "district" => "🔺 نواحی",
                "postname" => "⚡ پستها",
                "voltage" => "🔌 ولتاژ",
                "standard" => "📐 استاندارد",
                "type" => "🏗️ نوع پست",
                "circuit" => "🔄 مدار",
                "diesel" => "⚙️ دیزل",
                "water" => "💧 آب",
                "operationyear" => "📅 سال بهره‌برداری",
                "capacityhv" => "📊 ظرفیت HV",
                "capacitymv" => "📊 ظرفیت MV",
                _ => key
            };
        }

        private void RefreshAllCharts()
        {
            if (allPostsData == null)
                return;

            DataTable filtered = GetFilteredData();
            foreach (string key in charts.Keys)
            {
                Chart chart = charts[key];
                if (!specs.TryGetValue(key, out PostChartSpec? spec))
                    continue;

                SeriesChartType type = chart.Tag is SeriesChartType t ? t : SeriesChartType.Pie;
                PostChartRenderer.Render(chart, filtered, spec, type, true, true);
            }
        }

        private void Chart_MouseClick(string key, MouseEventArgs e)
        {
            if (allPostsData == null || !charts.TryGetValue(key, out Chart? chart))
                return;

            if (!specs.TryGetValue(key, out PostChartSpec? spec))
                return;

            try
            {
                HitTestResult result = chart.HitTest(e.X, e.Y);
                if (result.ChartElementType != ChartElementType.DataPoint || result.PointIndex < 0)
                    return;

                string clickedLabel = result.Series.Points[result.PointIndex].AxisLabel;
                if (string.IsNullOrWhiteSpace(clickedLabel))
                    clickedLabel = result.Series.Points[result.PointIndex].Label;

                DataTable filtered = GetFilteredData();
                if (filtered.Rows.Count == 0)
                    return;

                List<PostStat> stats = PostChartRenderer.BuildStats(filtered, spec,
                    PostChartRenderer.IsPie(PostChartRenderer.GetChartType(chart)) ? 12 : 0);
                List<string> topLabels = stats
                    .Where(stat => !stat.IsOther)
                    .Select(stat => stat.Label)
                    .ToList();

                List<DataRow> rows;
                if (clickedLabel == "سایر")
                {
                    rows = filtered.AsEnumerable()
                        .Where(row => !topLabels.Contains(PostChartRenderer.GetLabel(spec, row), StringComparer.CurrentCultureIgnoreCase))
                        .ToList();
                }
                else
                {
                    rows = filtered.AsEnumerable()
                        .Where(row => string.Equals(PostChartRenderer.GetLabel(spec, row), clickedLabel, StringComparison.CurrentCultureIgnoreCase))
                        .ToList();
                }

                if (rows.Count == 0)
                {
                    MessageBox.Show(this, "داده‌ای برای نمایش این دسته وجود ندارد.", "اطلاع",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DataTable details = filtered.Clone();
                foreach (DataRow row in rows)
                    details.ImportRow(row);

                if (ShowPostDetails(clickedLabel, details))
                {
                    allPostsData = postsService.LoadAllPosts(showError: false);
                    if (allPostsData != null)
                    {
                        FillAllFilterOptions();
                        UpdateFilterTitles();
                        RefreshAllCharts();
                        UpdateTables(GetFilteredData());
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPostsAnalytics.Chart_MouseClick", ex);
                MessageBox.Show(this, $"خطا در نمایش جزئیات:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportThisChart(string key, Chart chart)
        {
            if (!specs.TryGetValue(key, out PostChartSpec? spec))
                return;

            SeriesChartType type = chart.Tag is SeriesChartType t ? t : SeriesChartType.Pie;
            using var form = new FormExportPostCharts(GetFilteredData(), key, type);
            form.ShowDialog(this);
        }

        private bool ShowPostDetails(string category, DataTable postData)
        {
            bool changed = false;
            using var detailsForm = new Form
            {
                Text = $"جزئیات پست‌ها — {category}",
                StartPosition = FormStartPosition.CenterScreen,
                WindowState = FormWindowState.Maximized,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true,
                BackColor = Color.FromArgb(240, 248, 255),
                Font = FontSettings.BodyFont
            };

            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                ReadOnly = true,
                RightToLeft = RightToLeft.Yes,
                BackgroundColor = Color.White,
                EnableHeadersVisualStyles = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                ColumnHeadersHeight = 40,
                RowTemplate = { Height = 30 }
            };
            ConfigureGrid(grid);

            string[] columns =
            {
                "PostName|نام پست", "ProvinceName|استان", "CityName|شهر", "AffairName|امور",
                "DeptName|اداره", "DistrictName|ناحیه", "VoltageName|سطح ولتاژ",
                "TypeName|نوع پست", "StandardName|استاندارد", "CircuitName|وضعیت مدار",
                "FMName|ثابت/سیار", "OperationYear|سال بهره‌برداری", "CapacityHV|ظرفیت HV",
                "CapacityMV|ظرفیت MV", "PostalAddress|آدرس"
            };

            foreach (string column in columns)
            {
                string[] parts = column.Split('|');
                grid.Columns.Add(parts[0], parts[1]);
            }

            foreach (DataRow row in postData.Rows)
            {
                grid.Rows.Add(columns.Select(column =>
                {
                    string source = column.Split('|')[0];
                    return postData.Columns.Contains(source) ? row[source]?.ToString() ?? "" : "";
                }).ToArray());
            }

            var footer = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 54,
                BackColor = Color.FromArgb(230, 240, 250)
            };

            var lblInfo = new Label
            {
                Text = $"تعداد: {postData.Rows.Count} پست  |  دسته: {category}",
                Dock = DockStyle.Fill,
                Font = new Font(FontSettings.FontFamilyName, 10F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 16, 0)
            };

            var btnEdit = CreateSmallAction("✏️ ویرایش", Color.FromArgb(0, 102, 204));
            var btnDelete = CreateSmallAction("🗑️ حذف", Color.FromArgb(220, 53, 69));
            var btnClose = CreateSmallAction("❌ بستن", Color.FromArgb(108, 117, 125));

            btnEdit.Click += (_, _) =>
            {
                if (grid.CurrentRow == null || grid.CurrentRow.Cells["PostID"].Value == null)
                {
                    MessageBox.Show(detailsForm, "ابتدا یک پست را از جدول انتخاب کنید.", "اطلاع",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int postId = Convert.ToInt32(grid.CurrentRow.Cells["PostID"].Value);
                using var editor = new FormPostsEditor(postId);
                if (editor.ShowDialog(detailsForm) == DialogResult.OK)
                {
                    changed = true;
                    MessageBox.Show(detailsForm, "✅ پست با موفقیت ویرایش شد.", "موفقیت",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    detailsForm.Close();
                }
            };

            btnDelete.Click += (_, _) =>
            {
                if (grid.CurrentRow == null || grid.CurrentRow.Cells["PostID"].Value == null)
                {
                    MessageBox.Show(detailsForm, "ابتدا یک پست را از جدول انتخاب کنید.", "اطلاع",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int postId = Convert.ToInt32(grid.CurrentRow.Cells["PostID"].Value);
                string postName = grid.CurrentRow.Cells["PostName"].Value?.ToString() ?? "";
                if (MessageBox.Show(detailsForm,
                        $"آیا مطمئن هستید که می‌خواهید پست «{postName}» را حذف کنید؟",
                        "تأیید حذف", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                {
                    return;
                }

                var crud = new PostsCrudService();
                if (crud.DeletePost(postId))
                {
                    changed = true;
                    MessageBox.Show(detailsForm, "✅ پست با موفقیت حذف شد.", "موفقیت",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    detailsForm.Close();
                }
                else
                {
                    MessageBox.Show(detailsForm, "هیچ پستی حذف نشد.", "هشدار",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };

            btnClose.Click += (_, _) => detailsForm.Close();

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 400,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.Transparent,
                Padding = new Padding(0, 9, 8, 0)
            };
            actions.Controls.Add(btnClose);
            actions.Controls.Add(btnDelete);
            actions.Controls.Add(btnEdit);

            footer.Controls.Add(lblInfo);
            footer.Controls.Add(actions);

            detailsForm.Controls.Add(grid);
            detailsForm.Controls.Add(footer);
            detailsForm.ShowDialog(this);
            return changed;
        }

        private static Button CreateSmallAction(string text, Color backColor)
        {
            var button = new Button
            {
                Text = text,
                Size = new Size(130, 36),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(FontSettings.FontFamilyName, 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Margin = new Padding(4, 0, 0, 0)
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        // ─────────────────────────────────────────────
        // جداول
        // ─────────────────────────────────────────────
        private TabPage BuildStatsTab()
        {
            TabPage tab = new TabPage("📋 آمار")
            {
                Padding = new Padding(0)
            };

            Panel radioPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 42,
                BackColor = Color.FromArgb(230, 240, 250)
            };

            rbShowSummary.Text = "📊 خلاصه آماری";
            rbShowSummary.Location = new Point(10, 10);
            rbShowSummary.Size = new Size(150, 25);
            rbShowSummary.Checked = true;
            rbShowSummary.Font = FontSettings.ButtonFont;
            rbShowSummary.CheckedChanged += (_, _) =>
            {
                if (rbShowSummary.Checked && allPostsData != null)
                    LoadSummaryTable(GetFilteredData());
            };

            rbShowFullStats.Text = "📋 جدول کامل آمار";
            rbShowFullStats.Location = new Point(170, 10);
            rbShowFullStats.Size = new Size(170, 25);
            rbShowFullStats.Font = FontSettings.ButtonFont;
            rbShowFullStats.CheckedChanged += (_, _) =>
            {
                if (rbShowFullStats.Checked && allPostsData != null)
                    LoadStatisticalTable(GetFilteredData());
            };

            radioPanel.Controls.Add(rbShowFullStats);
            radioPanel.Controls.Add(rbShowSummary);

            ConfigureGrid(dgvSummary);
            dgvSummary.AutoGenerateColumns = false;
            dgvSummary.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "SummaryLabel",
                HeaderText = "شاخص",
                FillWeight = 72
            });
            dgvSummary.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "SummaryValue",
                HeaderText = "مقدار",
                FillWeight = 28
            });

            tab.Controls.Add(dgvSummary);
            tab.Controls.Add(radioPanel);
            return tab;
        }

        private TabPage BuildDeptTab()
        {
            TabPage tab = new TabPage("🏛️ جزئیات ادارات")
            {
                Padding = new Padding(0)
            };

            ConfigureGrid(dgvDeptDetails);
            dgvDeptDetails.AutoGenerateColumns = false;
            dgvDeptDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeptName", HeaderText = "اداره", FillWeight = 52 });
            dgvDeptDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeptCount", HeaderText = "تعداد", FillWeight = 24 });
            dgvDeptDetails.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeptPercent", HeaderText = "درصد", FillWeight = 24 });

            tab.Controls.Add(dgvDeptDetails);
            return tab;
        }

        private TabPage BuildPostsTab()
        {
            TabPage tab = new TabPage("⚡ لیست پست‌ها")
            {
                Padding = new Padding(0)
            };

            Panel searchPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 42,
                BackColor = Color.FromArgb(230, 240, 250)
            };

            txtPostSearch.Dock = DockStyle.Fill;
            txtPostSearch.Margin = new Padding(6);
            txtPostSearch.Font = FontSettings.TextBoxFont;
            txtPostSearch.PlaceholderText = "جستجو در نام پست، شهر، اداره یا آدرس...";
            txtPostSearch.TextChanged += (_, _) => LoadPostsTable(GetFilteredData());
            searchPanel.Controls.Add(txtPostSearch);

            ConfigureGrid(dgvPosts);
            dgvPosts.AutoGenerateColumns = false;
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "PostName", HeaderText = "نام پست", DataPropertyName = "PostName", Width = 120 });
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "ProvinceName", HeaderText = "استان", DataPropertyName = "ProvinceName", Width = 70 });
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "CityName", HeaderText = "شهر", DataPropertyName = "CityName", Width = 70 });
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "DeptName", HeaderText = "اداره", DataPropertyName = "DeptName", Width = 110 });
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "VoltageName", HeaderText = "ولتاژ", DataPropertyName = "VoltageName", Width = 80 });
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "TypeName", HeaderText = "نوع", DataPropertyName = "TypeName", Width = 75 });
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "OperationYear", HeaderText = "سال", DataPropertyName = "OperationYear", Width = 55 });
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "CapacityHV", HeaderText = "ظرفیت HV", DataPropertyName = "CapacityHV", Width = 75 });
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "CapacityMV", HeaderText = "ظرفیت MV", DataPropertyName = "CapacityMV", Width = 75 });
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn { Name = "PostalAddress", HeaderText = "آدرس", DataPropertyName = "PostalAddress", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });

            Panel gridPanel = new Panel { Dock = DockStyle.Fill };
            gridPanel.Controls.Add(dgvPosts);

            tab.Controls.Add(gridPanel);
            tab.Controls.Add(searchPanel);
            return tab;
        }

        private void UpdateTables(DataTable filtered)
        {
            if (rbShowSummary.Checked)
                LoadSummaryTable(filtered);
            else
                LoadStatisticalTable(filtered);

            LoadDeptDetailsTable(filtered);
            LoadPostsTable(filtered);
        }

        private void LoadSummaryTable(DataTable filtered)
        {
            dgvSummary.Rows.Clear();

            AddSummaryRow("کل پست‌ها", filtered.Rows.Count);
            AddSummaryRow("تعداد استان‌ها", DistinctCount(filtered, "ProvinceName"));
            AddSummaryRow("تعداد شهرها", DistinctCount(filtered, "CityName"));
            AddSummaryRow("تعداد امور", DistinctCount(filtered, "AffairName"));
            AddSummaryRow("تعداد ادارات", DistinctCount(filtered, "DeptName"));
            AddSummaryRow("تعداد نواحی", DistinctCount(filtered, "DistrictName"));
            AddSummaryRow("تعداد نام پست", DistinctCount(filtered, "PostName"));
            AddSummaryRow("سطوح ولتاژ", DistinctCount(filtered, "VoltageName"));
            AddSummaryRow("انواع پست", DistinctCount(filtered, "TypeName"));
            AddSummaryRow("استانداردها", DistinctCount(filtered, "StandardName"));

            dgvSummary.Rows.Add("", "");
            dgvSummary.Rows.Add("کیفیت داده", "");
            AddSummaryRow("بدون استان", CountValue(filtered, "ProvinceName", "نامشخص"));
            AddSummaryRow("بدون اداره", CountValue(filtered, "DeptName", "نامشخص"));
            AddSummaryRow("بدون سطح ولتاژ", CountValue(filtered, "VoltageName", "نامشخص"));
            AddSummaryRow("بدون سال بهره‌برداری", filtered.AsEnumerable().Count(row => ToInt(row, "OperationYear") <= 0));
            AddSummaryRow("بدون ظرفیت HV", filtered.AsEnumerable().Count(row => ToDouble(row, "CapacityHV") <= 0));
            AddSummaryRow("بدون ظرفیت MV", filtered.AsEnumerable().Count(row => ToDouble(row, "CapacityMV") <= 0));

            dgvSummary.Rows.Add("", "");
            dgvSummary.Rows.Add("ظرفیت‌ها", "");
            AddSummaryRow("مجموع ظرفیت HV", filtered.AsEnumerable().Sum(row => ToDouble(row, "CapacityHV")).ToString("N2"));
            AddSummaryRow("مجموع ظرفیت MV", filtered.AsEnumerable().Sum(row => ToDouble(row, "CapacityMV")).ToString("N2"));
            AddSummaryRow("مجموع ظرفیت توزیع‌شده", filtered.AsEnumerable().Sum(row => ToDouble(row, "DistributedCapacity")).ToString("N2"));
        }

        private void AddSummaryRow(string label, object value)
        {
            dgvSummary.Rows.Add(label, value);
        }

        private void LoadStatisticalTable(DataTable filtered)
        {
            dgvSummary.Rows.Clear();

            if (!dgvSummary.Columns.Contains("SummaryLabel"))
                return;

            foreach (PostChartSpec spec in PostChartCatalog.All)
            {
                List<PostStat> stats = PostChartRenderer.BuildStats(filtered, spec);
                if (stats.Count == 0)
                    continue;

                dgvSummary.Rows.Add(spec.Title, "");
                int total = stats.Sum(stat => stat.Count);
                foreach (PostStat stat in stats)
                {
                    double percent = total > 0 ? stat.Count * 100.0 / total : 0;
                    dgvSummary.Rows.Add($"  {stat.Label}", $"{stat.Count:N0}  ({percent:F1}%)");
                }
            }
        }

        private void LoadDeptDetailsTable(DataTable filtered)
        {
            dgvDeptDetails.Rows.Clear();
            var statsList = filtered.AsEnumerable()
                .GroupBy(row => Value(row, "DeptName"))
                .Select(group => new { Name = group.Key, Count = group.Count() })
                .OrderByDescending(item => item.Count)
                .ThenBy(item => item.Name, StringComparer.CurrentCulture)
                .ToList();

            int total = statsList.Sum(item => item.Count);
            foreach (var item in statsList)
            {
                double percent = total > 0 ? item.Count * 100.0 / total : 0;
                dgvDeptDetails.Rows.Add(item.Name, item.Count, $"{percent:F1}%");
            }
        }

        private void LoadPostsTable(DataTable filtered)
        {
            dgvPosts.Rows.Clear();
            string search = txtPostSearch?.Text?.Trim() ?? string.Empty;

            IEnumerable<DataRow> rows = filtered.AsEnumerable();
            if (!string.IsNullOrWhiteSpace(search))
            {
                rows = rows.Where(row =>
                    Contains(row, "PostName", search) ||
                    Contains(row, "ProvinceName", search) ||
                    Contains(row, "CityName", search) ||
                    Contains(row, "DeptName", search) ||
                    Contains(row, "PostalAddress", search));
            }

            foreach (DataRow row in rows)
            {
                dgvPosts.Rows.Add(
                    Value(row, "PostName"),
                    Value(row, "ProvinceName"),
                    Value(row, "CityName"),
                    Value(row, "DeptName"),
                    Value(row, "VoltageName"),
                    Value(row, "TypeName"),
                    row.Table.Columns.Contains("OperationYear") ? row["OperationYear"] : "",
                    ToDouble(row, "CapacityHV").ToString("N2"),
                    ToDouble(row, "CapacityMV").ToString("N2"),
                    Value(row, "PostalAddress"));
            }
        }

        private static bool Contains(DataRow row, string column, string search)
            => row.Table.Columns.Contains(column) &&
               Value(row, column).Contains(search, StringComparison.CurrentCultureIgnoreCase);

        private static int DistinctCount(DataTable table, string column)
        {
            if (!table.Columns.Contains(column))
                return 0;

            return table.AsEnumerable()
                .Select(row => Value(row, column))
                .Where(value => !string.IsNullOrWhiteSpace(value) && value != "نامشخص")
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .Count();
        }

        private static int CountValue(DataTable table, string column, string value)
        {
            if (!table.Columns.Contains(column))
                return 0;

            return table.AsEnumerable().Count(row =>
                string.Equals(Value(row, column), value, StringComparison.CurrentCultureIgnoreCase));
        }

        private static string Value(DataRow row, string column)
            => row.Table.Columns.Contains(column)
                ? row[column]?.ToString()?.Trim() ?? string.Empty
                : string.Empty;

        private static double ToDouble(DataRow row, string column)
        {
            if (!row.Table.Columns.Contains(column))
                return 0;

            object? value = row[column];
            if (value == null || value == DBNull.Value)
                return 0;

            return double.TryParse(value.ToString(), System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double result)
                ? result
                : 0;
        }

        private static int ToInt(DataRow row, string column)
        {
            if (!row.Table.Columns.Contains(column))
                return 0;

            object? value = row[column];
            if (value == null || value == DBNull.Value)
                return 0;

            return int.TryParse(value.ToString(), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out int result)
                ? result
                : 0;
        }

        // ─────────────────────────────────────────────
        // بارگذاری داده‌ها
        // ─────────────────────────────────────────────
        private void LoadData()
        {
            if (loading)
                return;

            loading = true;
            UseWaitCursor = true;
            try
            {
                if (!dbHelper.TestConnectionSilent())
                {
                    MessageBox.Show(this,
                        "❌ اتصال به پایگاه‌داده برقرار نشد.\nاز بخش تنظیمات، پیش‌نیاز Access را بررسی کنید.",
                        "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                allPostsData = postsService.LoadAllPosts(showError: false);
                if (allPostsData == null || allPostsData.Rows.Count == 0)
                {
                    MessageBox.Show(this, "⚠️ داده‌ای در جدول پست‌ها یافت نشد.", "اخطار",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                FillAllFilterOptions();
                UpdateFilterTitles();
                RefreshAllCharts();
                UpdateTables(GetFilteredData());
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPostsAnalytics.LoadData", ex);
                MessageBox.Show(this, $"خطا در بارگذاری داده‌ها:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                loading = false;
                UseWaitCursor = false;
            }
        }

        // ─────────────────────────────────────────────
        // خروجی اکسل
        // ─────────────────────────────────────────────
        private void ExportFilteredPostsToExcel()
        {
            DataTable filtered = GetFilteredData();
            if (filtered.Rows.Count == 0)
            {
                MessageBox.Show(this, "داده‌ای برای خروجی وجود ندارد.", "اطلاع",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dialog = new SaveFileDialog
            {
                Filter = "Excel Workbook|*.xlsx",
                Title = "ذخیره خروجی اکسل پست‌ها",
                FileName = $"پست‌ها_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                AddExtension = true
            };

            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            try
            {
                UseWaitCursor = true;
                Cursor = Cursors.WaitCursor;
                WritePostsWorkbook(filtered, dialog.FileName);
                UseWaitCursor = false;
                Cursor = Cursors.Default;

                MessageBox.Show(this, "✅ فایل اکسل با موفقیت ذخیره شد.", "خروجی اکسل",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                UseWaitCursor = false;
                Cursor = Cursors.Default;
                ErrorLogger.Log("FormPostsAnalytics.ExportFilteredPostsToExcel", ex);
                MessageBox.Show(this, $"خطا در ایجاد فایل اکسل:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void WritePostsWorkbook(DataTable data, string path)
        {
            var columns = new (string Column, string Header)[]
            {
                ("PostName", "نام پست"),
                ("ProvinceName", "استان"),
                ("CityName", "شهر"),
                ("AffairName", "امور"),
                ("DeptName", "اداره"),
                ("DistrictName", "ناحیه"),
                ("VoltageName", "سطح ولتاژ"),
                ("StandardName", "استاندارد"),
                ("TypeName", "نوع پست"),
                ("ConnName", "اتصال توزیع"),
                ("InsName", "نوع عایق"),
                ("PT2Name", "نوع پست ۲"),
                ("FMName", "ثابت / سیار"),
                ("CircuitName", "وضعیت مدار"),
                ("DieselName", "دیزل ژنراتور"),
                ("FeedName", "نوع توزیع"),
                ("WaterName", "وضعیت آب"),
                ("GuestName", "میهمان‌سرا"),
                ("OperationYear", "سال بهره‌برداری"),
                ("DistributedCapacity", "ظرفیت توزیع‌شده"),
                ("CapacityHV", "ظرفیت HV"),
                ("CapacityMV", "ظرفیت MV"),
                ("TotalArea", "مساحت کل"),
                ("BuildingsArea", "مساحت ساختمان"),
                ("SiteArea", "مساحت محوطه"),
                ("ParkingBays", "تعداد پارکینگ"),
                ("DistanceToCapital", "فاصله تا مرکز"),
                ("PostalCode", "کد پستی"),
                ("PostalAddress", "آدرس"),
                ("Description", "توضیحات")
            };

            using var workbook = new XLWorkbook();
            IXLWorksheet sheet = workbook.Worksheets.Add("پست‌ها");
            sheet.RightToLeft = true;

            for (int c = 0; c < columns.Length; c++)
                sheet.Cell(1, c + 1).Value = columns[c].Header;

            IXLRange headerRange = sheet.Range(1, 1, 1, columns.Length);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 102, 204);
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Row(1).Height = 28;

            for (int r = 0; r < data.Rows.Count; r++)
            {
                for (int c = 0; c < columns.Length; c++)
                {
                    string source = columns[c].Column;
                    object? value = data.Columns.Contains(source) ? data.Rows[r][source] : null;
                    IXLCell cell = sheet.Cell(r + 2, c + 1);

                    if (value == null || value == DBNull.Value)
                    {
                        cell.Value = "";
                    }
                    else if (value is double or float or decimal or int or long)
                    {
                        cell.Value = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        cell.Value = value.ToString();
                    }
                }
            }

            sheet.Range(1, 1, Math.Max(1, data.Rows.Count + 1), columns.Length).SetAutoFilter();
            sheet.SheetView.FreezeRows(1);
            sheet.Columns().AdjustToContents(10, 40);

            workbook.SaveAs(path);
        }

        private void ConfigureGrid(DataGridView grid)
        {
            grid.Dock = DockStyle.Fill;
            grid.ReadOnly = true;
            grid.RightToLeft = RightToLeft.Yes;
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false;
            grid.RowHeadersVisible = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            grid.ColumnHeadersDefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.ColumnHeadersHeight = 36;
            grid.DefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F);
            grid.DefaultCellStyle.BackColor = Color.White;
            grid.DefaultCellStyle.ForeColor = Color.FromArgb(30, 41, 59);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(30, 41, 59);
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            grid.RowTemplate.Height = 28;
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
        }

        private sealed class FilterBinding
        {
            public FilterBinding(CheckedListBox list, Label title, string column, string displayName)
            {
                List = list;
                Title = title;
                Column = column;
                DisplayName = displayName;
            }

            public CheckedListBox List { get; }
            public Label Title { get; }
            public string Column { get; }
            public string DisplayName { get; }
        }
    }
}
