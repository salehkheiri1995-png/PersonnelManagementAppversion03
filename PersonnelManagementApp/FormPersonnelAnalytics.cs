using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace PersonnelManagementApp
{
    public partial class FormPersonnelAnalytics : Form
    {
        private readonly DbHelper dbHelper;
        private readonly TabControl tabControl;
        private readonly AnalyticsDataModel analyticsModel;

        private readonly Chart chartDepartmentPie;
        private readonly Chart chartPositionPie;
        private readonly Chart chartGenderPie;
        private readonly Chart chartJobLevelPie;
        private readonly Chart chartContractTypePie;
        private readonly Chart chartProvincePie;
        private readonly Chart chartEducationPie;
        private readonly Chart chartCompanyPie;
        private readonly Chart chartWorkShiftPie;
        private readonly Chart chartAgePie;
        private readonly Chart chartWorkExperiencePie;

        private readonly DataGridView dgvPersonnelStats;
        private readonly DataGridView dgvDepartmentDetails;
        private readonly DataGridView dgvPositionDetails;

        private readonly CheckedListBox clbProvincesFilter;
        private readonly CheckedListBox clbCitiesFilter;
        private readonly CheckedListBox clbAffairsFilter;
        private readonly CheckedListBox clbDepartmentsFilter;
        private readonly CheckedListBox clbDistrictsFilter;
        private readonly CheckedListBox clbPositionsFilter;
        private readonly CheckedListBox clbEducationFilter;
        private readonly CheckedListBox clbJobLevelFilter;
        private readonly CheckedListBox clbContractTypeFilter;
        private readonly CheckedListBox clbCompanyFilter;
        private readonly CheckedListBox clbWorkShiftFilter;
        private readonly CheckedListBox clbGenderFilter;

        private readonly Button btnClearFilters;
        private readonly Label lblFilterInfo;

        private DateTimePicker? dtpHireDateFrom;
        private DateTimePicker? dtpHireDateTo;
        private CheckBox? chkHireDateFilter;

        private NumericUpDown? nudMinAge;
        private NumericUpDown? nudMaxAge;
        private CheckBox? chkAgeFilter;

        private NumericUpDown? nudMinExperience;
        private NumericUpDown? nudMaxExperience;
        private CheckBox? chkExperienceFilter;

        private RadioButton? rbShowSummary;
        private RadioButton? rbShowFullStats;

        private ContextMenuStrip? chartTypeMenu;

        private NumericUpDown? nudAgeRangeSize;
        private Label? lblAgeRangeSize;

        public FormPersonnelAnalytics()
        {
            dbHelper = new DbHelper();
            analyticsModel = new AnalyticsDataModel();
            tabControl = new TabControl();

            chartDepartmentPie = new Chart();
            chartPositionPie = new Chart();
            chartGenderPie = new Chart();
            chartJobLevelPie = new Chart();
            chartContractTypePie = new Chart();
            chartProvincePie = new Chart();
            chartEducationPie = new Chart();
            chartCompanyPie = new Chart();
            chartWorkShiftPie = new Chart();
            chartAgePie = new Chart();
            chartWorkExperiencePie = new Chart();

            dgvPersonnelStats = new DataGridView();
            dgvDepartmentDetails = new DataGridView();
            dgvPositionDetails = new DataGridView();

            clbProvincesFilter = new CheckedListBox();
            clbCitiesFilter = new CheckedListBox();
            clbAffairsFilter = new CheckedListBox();
            clbDepartmentsFilter = new CheckedListBox();
            clbDistrictsFilter = new CheckedListBox();
            clbPositionsFilter = new CheckedListBox();
            clbEducationFilter = new CheckedListBox();
            clbJobLevelFilter = new CheckedListBox();
            clbContractTypeFilter = new CheckedListBox();
            clbCompanyFilter = new CheckedListBox();
            clbWorkShiftFilter = new CheckedListBox();
            clbGenderFilter = new CheckedListBox();

            btnClearFilters = new Button();
            lblFilterInfo = new Label();

            rbShowSummary = new RadioButton();
            rbShowFullStats = new RadioButton();

            InitializeComponent();
            BuildUI();
            InitializeChartTypeMenu();

            FontSettings.ApplyFontToForm(this);

            LoadData();
        }

        private void BuildUI()
        {
            Text = "🎯 تحلیل داده‌های پرسنل - سیستم پیشرفته";
            WindowState = FormWindowState.Maximized;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            BackColor = Color.FromArgb(240, 248, 255);
            MinimumSize = new Size(1200, 700);
            Font = FontSettings.BodyFont;

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

            filterGrid.Controls.Add(CreateFilterBox("استانها 🗺️", clbProvincesFilter, ClbProvincesFilter_ItemCheck), 0, 0);
            filterGrid.Controls.Add(CreateFilterBox("شهرها 🏙️", clbCitiesFilter, ClbCitiesFilter_ItemCheck), 1, 0);
            filterGrid.Controls.Add(CreateFilterBox("امور 📋", clbAffairsFilter, ClbAffairsFilter_ItemCheck), 2, 0);
            filterGrid.Controls.Add(CreateFilterBox("ادارات 🏛️", clbDepartmentsFilter, ClbDepartmentsFilter_ItemCheck), 3, 0);
            filterGrid.Controls.Add(CreateFilterBox("نواحی 🔺", clbDistrictsFilter, ClbDistrictsFilter_ItemCheck), 4, 0);
            filterGrid.Controls.Add(CreateFilterBox("پستها ⚡", clbPositionsFilter, ClbPositionsFilter_ItemCheck), 5, 0);

            filterGrid.Controls.Add(CreateFilterBox("جنسیت 👥", clbGenderFilter, ClbGenderFilter_ItemCheck), 0, 1);
            filterGrid.Controls.Add(CreateFilterBox("تحصیلات 📚", clbEducationFilter, ClbEducationFilter_ItemCheck), 1, 1);
            filterGrid.Controls.Add(CreateFilterBox("سطح شغلی 📊", clbJobLevelFilter, ClbJobLevelFilter_ItemCheck), 2, 1);
            filterGrid.Controls.Add(CreateFilterBox("نوع قرارداد 📄", clbContractTypeFilter, ClbContractTypeFilter_ItemCheck), 3, 1);
            filterGrid.Controls.Add(CreateFilterBox("شرکت 🏢", clbCompanyFilter, ClbCompanyFilter_ItemCheck), 4, 1);
            filterGrid.Controls.Add(CreateFilterBox("شیفت کاری ⏰", clbWorkShiftFilter, ClbWorkShiftFilter_ItemCheck), 5, 1);

            panelFilter.Controls.Add(filterGrid);

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
            btnClearFilters.Click += BtnClearFilters_Click;

            Label lblHireDate = new Label
            {
                Text = "📅 تاریخ استخدام",
                AutoSize = true,
                Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Margin = new Padding(4, 11, 4, 4)
            };

            chkHireDateFilter = new CheckBox
            {
                Text = "فعال",
                AutoSize = true,
                Font = new Font(FontSettings.BodyFont.FontFamily, 9F),
                Margin = new Padding(4, 11, 4, 4)
            };
            chkHireDateFilter.CheckedChanged += ChkHireDateFilter_CheckedChanged;

            dtpHireDateFrom = new DateTimePicker
            {
                Size = new Size(140, 28),
                Font = new Font(FontSettings.TextBoxFont.FontFamily, 9F),
                Enabled = false,
                Value = DateTime.Now.AddYears(-10),
                Format = DateTimePickerFormat.Short,
                Margin = new Padding(4, 8, 4, 4)
            };

            Label lblTo = new Label
            {
                Text = "تا",
                AutoSize = true,
                Font = new Font(FontSettings.LabelFont.FontFamily, 9F),
                Margin = new Padding(4, 11, 4, 4)
            };

            dtpHireDateTo = new DateTimePicker
            {
                Size = new Size(140, 28),
                Font = new Font(FontSettings.TextBoxFont.FontFamily, 9F),
                Enabled = false,
                Value = DateTime.Now,
                Format = DateTimePickerFormat.Short,
                Margin = new Padding(4, 8, 4, 4)
            };

            Label lblAge = new Label
            {
                Text = "🎂 سن",
                AutoSize = true,
                Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Margin = new Padding(10, 11, 4, 4)
            };

            chkAgeFilter = new CheckBox
            {
                Text = "فعال",
                AutoSize = true,
                Font = new Font(FontSettings.BodyFont.FontFamily, 9F),
                Margin = new Padding(4, 11, 4, 4)
            };
            chkAgeFilter.CheckedChanged += (s, e) =>
            {
                if (nudMinAge != null) nudMinAge.Enabled = chkAgeFilter.Checked;
                if (nudMaxAge != null) nudMaxAge.Enabled = chkAgeFilter.Checked;
                UpdateFilters();
                RefreshAllCharts();
            };

            nudMinAge = new NumericUpDown
            {
                Minimum = 10,
                Maximum = 100,
                Value = 18,
                Width = 70,
                Enabled = false,
                Margin = new Padding(4, 8, 4, 4)
            };

            Label lblAgeTo = new Label
            {
                Text = "تا",
                AutoSize = true,
                Margin = new Padding(4, 11, 4, 4)
            };

            nudMaxAge = new NumericUpDown
            {
                Minimum = 10,
                Maximum = 100,
                Value = 65,
                Width = 70,
                Enabled = false,
                Margin = new Padding(4, 8, 4, 4)
            };

            Label lblExp = new Label
            {
                Text = "💼 سابقه کاری",
                AutoSize = true,
                Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Margin = new Padding(10, 11, 4, 4)
            };

            chkExperienceFilter = new CheckBox
            {
                Text = "فعال",
                AutoSize = true,
                Font = new Font(FontSettings.BodyFont.FontFamily, 9F),
                Margin = new Padding(4, 11, 4, 4)
            };
            chkExperienceFilter.CheckedChanged += (s, e) =>
            {
                if (nudMinExperience != null) nudMinExperience.Enabled = chkExperienceFilter.Checked;
                if (nudMaxExperience != null) nudMaxExperience.Enabled = chkExperienceFilter.Checked;
                UpdateFilters();
                RefreshAllCharts();
            };

            nudMinExperience = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 50,
                Value = 0,
                Width = 70,
                Enabled = false,
                Margin = new Padding(4, 8, 4, 4)
            };

            Label lblExpTo = new Label
            {
                Text = "تا",
                AutoSize = true,
                Margin = new Padding(4, 11, 4, 4)
            };

            nudMaxExperience = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 50,
                Value = 40,
                Width = 70,
                Enabled = false,
                Margin = new Padding(4, 8, 4, 4)
            };

            rowActions.Controls.Add(btnClearFilters);
            rowActions.Controls.Add(nudMaxExperience);
            rowActions.Controls.Add(lblExpTo);
            rowActions.Controls.Add(nudMinExperience);
            rowActions.Controls.Add(chkExperienceFilter);
            rowActions.Controls.Add(lblExp);
            rowActions.Controls.Add(nudMaxAge);
            rowActions.Controls.Add(lblAgeTo);
            rowActions.Controls.Add(nudMinAge);
            rowActions.Controls.Add(chkAgeFilter);
            rowActions.Controls.Add(lblAge);
            rowActions.Controls.Add(dtpHireDateTo);
            rowActions.Controls.Add(lblTo);
            rowActions.Controls.Add(dtpHireDateFrom);
            rowActions.Controls.Add(chkHireDateFilter);
            rowActions.Controls.Add(lblHireDate);

            lblFilterInfo.Text = "✓ فیلتری فعال نیست";
            lblFilterInfo.Dock = DockStyle.Bottom;
            lblFilterInfo.Height = 30;
            lblFilterInfo.Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold);
            lblFilterInfo.ForeColor = Color.FromArgb(0, 102, 204);
            lblFilterInfo.TextAlign = ContentAlignment.MiddleLeft;

            filterBottomPanel.Controls.Add(lblFilterInfo);
            filterBottomPanel.Controls.Add(rowActions);

            panelFilter.Controls.Add(filterBottomPanel);

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

            // 🔥 حذف دکمه بالایی - فقط عنوان باقی می‌مونه
            Panel chartHeaderPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = Color.FromArgb(0, 102, 204),
                Padding = new Padding(10, 8, 10, 8)
            };

            Label lblChartsTitle = new Label
            {
                Text = "📊 نمودارهای آماری",
                Font = new Font(FontSettings.HeaderFont.FontFamily, 12F, FontStyle.Bold),
                ForeColor = Color.White,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleRight
            };
            chartHeaderPanel.Controls.Add(lblChartsTitle);

            tabControl.Dock = DockStyle.Fill;
            tabControl.RightToLeft = RightToLeft.Yes;
            tabControl.RightToLeftLayout = true;
            tabControl.ItemSize = new Size(120, 30);
            tabControl.Font = FontSettings.BodyFont;

            AddChartTab(tabControl, "📊 ادارات", chartDepartmentPie, null);
            AddChartTab(tabControl, "💼 پستها", chartPositionPie, null);
            AddChartTab(tabControl, "👥 جنسیت", chartGenderPie, null);
            AddChartTab(tabControl, "📈 سطح شغلی", chartJobLevelPie, null);
            AddChartTab(tabControl, "📋 قرارداد", chartContractTypePie, null);
            AddChartTab(tabControl, "🗺️ استان", chartProvincePie, null);
            AddChartTab(tabControl, "📚 تحصیلات", chartEducationPie, null);
            AddChartTab(tabControl, "🏢 شرکت", chartCompanyPie, null);
            AddChartTab(tabControl, "⏰ شیفت", chartWorkShiftPie, null);
            AddChartTab(tabControl, "🎂 سن", chartAgePie, null);
            AddChartTab(tabControl, "💼 سابقه کاری", chartWorkExperiencePie, null);

            chartsPanel.Controls.Add(tabControl);
            chartsPanel.Controls.Add(chartHeaderPanel);

            Panel tablesPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(4)
            };

            TabControl tablesTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true,
                Font = FontSettings.BodyFont,
                ItemSize = new Size(130, 30)
            };

            TabPage tabStats = new TabPage("📋 آمار")
            {
                Padding = new Padding(0)
            };

            Panel radioPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 42,
                BackColor = Color.FromArgb(230, 240, 250)
            };

            rbShowSummary = new RadioButton
            {
                Text = "📊 خلاصه آماری",
                Location = new Point(10, 10),
                Size = new Size(150, 25),
                Checked = true,
                Font = FontSettings.ButtonFont
            };
            rbShowSummary.CheckedChanged += RbShowSummary_CheckedChanged;
            radioPanel.Controls.Add(rbShowSummary);

            rbShowFullStats = new RadioButton
            {
                Text = "📋 جدول کامل آمار",
                Location = new Point(170, 10),
                Size = new Size(170, 25),
                Font = FontSettings.ButtonFont
            };
            rbShowFullStats.CheckedChanged += RbShowFullStats_CheckedChanged;
            radioPanel.Controls.Add(rbShowFullStats);

            dgvPersonnelStats.Dock = DockStyle.Fill;
            dgvPersonnelStats.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvPersonnelStats.ReadOnly = true;
            dgvPersonnelStats.RightToLeft = RightToLeft.Yes;
            dgvPersonnelStats.BackgroundColor = Color.White;
            dgvPersonnelStats.EnableHeadersVisualStyles = false;
            dgvPersonnelStats.AllowUserToAddRows = false;
            dgvPersonnelStats.AllowUserToDeleteRows = false;
            dgvPersonnelStats.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
            dgvPersonnelStats.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPersonnelStats.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
            dgvPersonnelStats.ColumnHeadersHeight = 35;
            dgvPersonnelStats.DefaultCellStyle.BackColor = Color.White;
            dgvPersonnelStats.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgvPersonnelStats.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            tabStats.Controls.Add(dgvPersonnelStats);
            tabStats.Controls.Add(radioPanel);

            TabPage tabDeptDetails = new TabPage("🏛️ جزئیات ادارات")
            {
                Padding = new Padding(0)
            };

            dgvDepartmentDetails.Dock = DockStyle.Fill;
            dgvDepartmentDetails.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvDepartmentDetails.ReadOnly = true;
            dgvDepartmentDetails.RightToLeft = RightToLeft.Yes;
            dgvDepartmentDetails.BackgroundColor = Color.White;
            dgvDepartmentDetails.EnableHeadersVisualStyles = false;
            dgvDepartmentDetails.AllowUserToAddRows = false;
            dgvDepartmentDetails.AllowUserToDeleteRows = false;
            dgvDepartmentDetails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
            dgvDepartmentDetails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvDepartmentDetails.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
            dgvDepartmentDetails.ColumnHeadersHeight = 35;
            dgvDepartmentDetails.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgvDepartmentDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            tabDeptDetails.Controls.Add(dgvDepartmentDetails);

            TabPage tabPosDetails = new TabPage("💼 جزئیات پست‌ها")
            {
                Padding = new Padding(0)
            };

            dgvPositionDetails.Dock = DockStyle.Fill;
            dgvPositionDetails.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvPositionDetails.ReadOnly = true;
            dgvPositionDetails.RightToLeft = RightToLeft.Yes;
            dgvPositionDetails.BackgroundColor = Color.White;
            dgvPositionDetails.EnableHeadersVisualStyles = false;
            dgvPositionDetails.AllowUserToAddRows = false;
            dgvPositionDetails.AllowUserToDeleteRows = false;
            dgvPositionDetails.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
            dgvPositionDetails.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPositionDetails.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
            dgvPositionDetails.ColumnHeadersHeight = 35;
            dgvPositionDetails.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgvPositionDetails.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);
            tabPosDetails.Controls.Add(dgvPositionDetails);

            tablesTabControl.TabPages.Add(tabStats);
            tablesTabControl.TabPages.Add(tabDeptDetails);
            tablesTabControl.TabPages.Add(tabPosDetails);

            tablesPanel.Controls.Add(tablesTabControl);

            mainLayout.Controls.Add(chartsPanel, 0, 0);
            mainLayout.Controls.Add(tablesPanel, 1, 0);

            Controls.Add(mainLayout);
            Controls.Add(panelFilter);
        }

        private Panel CreateFilterBox(string title, CheckedListBox clb, ItemCheckEventHandler eventHandler)
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

            SetupFilterCheckedListBox(clb, eventHandler);
            clb.Dock = DockStyle.Fill;

            box.Controls.Add(clb);
            box.Controls.Add(lbl);
            return box;
        }

        private void SetupFilterCheckedListBox(CheckedListBox clb, ItemCheckEventHandler eventHandler)
        {
            clb.RightToLeft = RightToLeft.Yes;
            clb.CheckOnClick = true;
            clb.BackColor = Color.White;
            clb.Font = new Font(FontSettings.BodyFont.FontFamily, 9F);
            clb.IntegralHeight = false;
            clb.BorderStyle = BorderStyle.FixedSingle;
            clb.HorizontalScrollbar = false;
            clb.HorizontalExtent = 0;
            clb.ScrollAlwaysVisible = false;
            clb.ItemHeight = 18;

            clb.ItemCheck -= eventHandler;
            clb.ItemCheck += eventHandler;
        }

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

            chartDepartmentPie.ContextMenuStrip = chartTypeMenu;
            chartPositionPie.ContextMenuStrip = chartTypeMenu;
            chartGenderPie.ContextMenuStrip = chartTypeMenu;
            chartJobLevelPie.ContextMenuStrip = chartTypeMenu;
            chartContractTypePie.ContextMenuStrip = chartTypeMenu;
            chartProvincePie.ContextMenuStrip = chartTypeMenu;
            chartEducationPie.ContextMenuStrip = chartTypeMenu;
            chartCompanyPie.ContextMenuStrip = chartTypeMenu;
            chartWorkShiftPie.ContextMenuStrip = chartTypeMenu;
            chartAgePie.ContextMenuStrip = chartTypeMenu;
            chartWorkExperiencePie.ContextMenuStrip = chartTypeMenu;
        }

        private void AddChartTypeMenuItem(string text, SeriesChartType type)
        {
            var item = new ToolStripMenuItem(text) { Tag = type };
            if (chartTypeMenu != null)
                chartTypeMenu.Items.Add(item);
        }

        private void ChartTypeMenu_ItemClicked(object? sender, ToolStripItemClickedEventArgs e)
        {
            if (chartTypeMenu != null)
                chartTypeMenu.Hide();

            if (e.ClickedItem == null || e.ClickedItem.Tag == null)
                return;

            if (!(e.ClickedItem.Tag is SeriesChartType type))
                return;

            var cms = sender as ContextMenuStrip;
            var chart = cms?.SourceControl as Chart;
            if (chart == null)
                return;

            chart.Tag = type;
            ApplyChartTypeToChart(chart, type);
        }

        private static bool IsPieType(SeriesChartType type)
        {
            return type == SeriesChartType.Pie || type == SeriesChartType.Doughnut;
        }

        private SeriesChartType GetChartTypeOrDefault(Chart chart)
        {
            if (chart != null && chart.Tag is SeriesChartType type)
                return type;

            return SeriesChartType.Pie;
        }

        private void ConfigureChartAreaForType(Chart chart, SeriesChartType type)
        {
            if (chart == null || chart.ChartAreas.Count == 0)
                return;

            var area = chart.ChartAreas[0];
            bool pie = IsPieType(type);

            area.Area3DStyle.Enable3D = pie;
            if (pie)
            {
                area.Area3DStyle.Inclination = 15;
                area.Area3DStyle.Rotation = 45;
            }

            area.AxisX.Enabled = pie ? AxisEnabled.False : AxisEnabled.True;
            area.AxisY.Enabled = pie ? AxisEnabled.False : AxisEnabled.True;

            area.AxisX.MajorGrid.Enabled = !pie;
            area.AxisY.MajorGrid.Enabled = !pie;
            area.AxisX.MajorGrid.LineColor = Color.Gainsboro;
            area.AxisY.MajorGrid.LineColor = Color.Gainsboro;

            area.AxisX.IsLabelAutoFit = true;
            area.AxisY.IsLabelAutoFit = true;

            if (!pie)
            {
                if (type == SeriesChartType.Bar || type == SeriesChartType.StackedBar)
                {
                    area.AxisY.Interval = 1;
                    area.AxisX.Interval = 0;
                    area.AxisX.LabelStyle.Angle = 0;
                }
                else
                {
                    area.AxisX.Interval = 1;
                    area.AxisY.Interval = 0;

                    if (chart.Series.Count > 0 && chart.Series[0].Points.Count > 10)
                        area.AxisX.LabelStyle.Angle = -45;
                    else
                        area.AxisX.LabelStyle.Angle = 0;
                }
            }
        }

        private void ConfigureSeriesForType(Series series, SeriesChartType type)
        {
            if (series == null)
                return;

            series.ChartType = type;
            series.XValueType = ChartValueType.String;
            series.YValueType = ChartValueType.Int32;
            series.IsXValueIndexed = true;
            
            // 🎯 Safe font fallback
            series.Font = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);

            if (IsPieType(type))
            {
                series.IsValueShownAsLabel = true;
                series["PieLabelStyle"] = "Outside";
            }
            else
            {
                series.IsValueShownAsLabel = true;
                if (series.CustomProperties != null && series.CustomProperties.Contains("PieLabelStyle"))
                    series["PieLabelStyle"] = "";
            }
        }

        private void ApplyChartTypeToChart(Chart chart, SeriesChartType type)
        {
            if (chart == null)
                return;

            ConfigureChartAreaForType(chart, type);

            foreach (Series series in chart.Series)
            {
                ConfigureSeriesForType(series, type);
            }
        }

        private void ApplyChartTypeFromTag(Chart chart)
        {
            if (chart.Tag is SeriesChartType type)
            {
                ApplyChartTypeToChart(chart, type);
            }
        }

        private void ReapplyChartTypes()
        {
            ApplyChartTypeFromTag(chartDepartmentPie);
            ApplyChartTypeFromTag(chartPositionPie);
            ApplyChartTypeFromTag(chartGenderPie);
            ApplyChartTypeFromTag(chartJobLevelPie);
            ApplyChartTypeFromTag(chartContractTypePie);
            ApplyChartTypeFromTag(chartProvincePie);
            ApplyChartTypeFromTag(chartEducationPie);
            ApplyChartTypeFromTag(chartCompanyPie);
            ApplyChartTypeFromTag(chartWorkShiftPie);
            ApplyChartTypeFromTag(chartAgePie);
            ApplyChartTypeFromTag(chartWorkExperiencePie);
        }

        private void RbShowSummary_CheckedChanged(object? sender, EventArgs e)
        {
            if (rbShowSummary != null && rbShowSummary.Checked)
            {
                LoadSummaryTable();
            }
        }

        private void RbShowFullStats_CheckedChanged(object? sender, EventArgs e)
        {
            if (rbShowFullStats != null && rbShowFullStats.Checked)
            {
                LoadStatisticalTable();
            }
        }

        private void AddChartTab(TabControl tabControl, string title, Chart chart, DataGridView? detailsGrid)
        {
            TabPage tab = new TabPage(title);

            if (detailsGrid != null)
            {
                SplitContainer split = new SplitContainer
                {
                    Dock = DockStyle.Fill,
                    Orientation = Orientation.Horizontal,
                    SplitterDistance = 400
                };

                chart.Dock = DockStyle.Fill;
                chart.BackColor = Color.White;
                chart.MinimumSize = new Size(100, 100);
                chart.ChartAreas.Add(new ChartArea("ChartArea1")
                {
                    BackColor = Color.White,
                    Area3DStyle = { Enable3D = true, Inclination = 15, Rotation = 45 }
                });
                chart.MouseClick += Chart_MouseClick;
                split.Panel1.Controls.Add(chart);

                detailsGrid.Dock = DockStyle.Fill;
                detailsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                detailsGrid.ReadOnly = true;
                detailsGrid.RightToLeft = RightToLeft.Yes;
                detailsGrid.BackgroundColor = Color.White;
                detailsGrid.EnableHeadersVisualStyles = false;
                detailsGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
                detailsGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                detailsGrid.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
                split.Panel2.Controls.Add(detailsGrid);

                tab.Controls.Add(split);
            }
            else
            {
                // 🔥 دکمه Export کوچیک گوشه نمودار
                Button btnExportThis = new Button
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
                btnExportThis.FlatAppearance.BorderSize = 0;
                btnExportThis.Location = new Point(10, 10);
                btnExportThis.Click += (s, e) =>
                {
                    FormExportCharts exportForm = new FormExportCharts();
                    exportForm.ShowDialog();
                };

                chart.Dock = DockStyle.Fill;
                chart.BackColor = Color.White;
                chart.MinimumSize = new Size(100, 100);
                chart.ChartAreas.Add(new ChartArea("ChartArea1")
                {
                    BackColor = Color.White,
                    Area3DStyle = { Enable3D = true, Inclination = 15, Rotation = 45 }
                });
                chart.MouseClick += Chart_MouseClick;

                if (title == "🎂 سن")
                {
                    Panel topPanel = new Panel
                    {
                        Dock = DockStyle.Top,
                        Height = 45,
                        BackColor = Color.FromArgb(230, 240, 250),
                        Padding = new Padding(10, 8, 10, 8)
                    };

                    lblAgeRangeSize = new Label
                    {
                        Text = "📊 بازه سنی (سال):",
                        AutoSize = true,
                        Font = new Font(FontSettings.SubtitleFont.FontFamily, 10F, FontStyle.Bold),
                        ForeColor = Color.FromArgb(0, 102, 204),
                        Location = new Point(10, 12)
                    };

                    nudAgeRangeSize = new NumericUpDown
                    {
                        Minimum = 1,
                        Maximum = 10,
                        Value = 10,
                        Width = 70,
                        Font = new Font(FontSettings.TextBoxFont.FontFamily, 10F),
                        Location = new Point(150, 9)
                    };
                    nudAgeRangeSize.ValueChanged += (s, e) => LoadAgePieChart();

                    topPanel.Controls.Add(nudAgeRangeSize);
                    topPanel.Controls.Add(lblAgeRangeSize);

                    tab.Controls.Add(chart);
                    tab.Controls.Add(topPanel);
                    tab.Controls.Add(btnExportThis);
                    btnExportThis.BringToFront();
                }
                else
                {
                    tab.Controls.Add(chart);
                    tab.Controls.Add(btnExportThis);
                    btnExportThis.BringToFront();
                }
            }

            tabControl.TabPages.Add(tab);
        }

        private void LoadData()
        {
            try
            {
                if (!dbHelper.TestConnection())
                {
                    MessageBox.Show("❌ اتصال به دیتابیس ناموفق بود.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!analyticsModel.LoadData(dbHelper))
                {
                    MessageBox.Show("❌ خطا در بارگذاری داده‌ها.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                LoadFilterOptions();
                RefreshAllCharts();
                MessageBox.Show($"✅ داده‌ها با موفقیت بارگذاری شدند.\n👥 تعداد پرسنل: {analyticsModel.TotalPersonnel}", "موفق", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadFilterOptions()
        {
            var allProvinces = analyticsModel.GetAllProvinces().Distinct().OrderBy(x => x).ToList();

            clbProvincesFilter.Items.Clear();
            foreach (var p in allProvinces)
                clbProvincesFilter.Items.Add(p, false);

            clbCitiesFilter.Items.Clear();
            foreach (var c in analyticsModel.GetCitiesByProvinces(allProvinces).Distinct().OrderBy(x => x))
                clbCitiesFilter.Items.Add(c, false);

            clbAffairsFilter.Items.Clear();
            foreach (var a in analyticsModel.GetAffairsByProvinces(allProvinces).Distinct().OrderBy(x => x))
                clbAffairsFilter.Items.Add(a, false);

            var allCities = clbCitiesFilter.Items.Cast<string>().ToList();
            var allAffairs = clbAffairsFilter.Items.Cast<string>().ToList();

            var allDepts = analyticsModel.GetDepartmentsByFilters(allProvinces, allCities, allAffairs).Distinct().OrderBy(x => x).ToList();
            clbDepartmentsFilter.Items.Clear();
            foreach (var d in allDepts)
                clbDepartmentsFilter.Items.Add(d, false);

            var allDistricts = allDepts.Count > 0
                ? analyticsModel.GetDistrictsByDepartments(allDepts).Distinct().OrderBy(x => x).ToList()
                : new List<string>();
            clbDistrictsFilter.Items.Clear();
            foreach (var dist in allDistricts)
                clbDistrictsFilter.Items.Add(dist, false);

            var allPositions = allDistricts.Count > 0
                ? analyticsModel.GetPositionsByDistricts(allDistricts).Distinct().OrderBy(x => x).ToList()
                : new List<string>();
            clbPositionsFilter.Items.Clear();
            foreach (var pos in allPositions)
                clbPositionsFilter.Items.Add(pos, false);

            clbGenderFilter.Items.Clear();
            foreach (var g in analyticsModel.GetAllGenders())
                clbGenderFilter.Items.Add(g, false);

            clbEducationFilter.Items.Clear();
            foreach (var e in analyticsModel.GetAllEducations())
                clbEducationFilter.Items.Add(e, false);

            clbJobLevelFilter.Items.Clear();
            foreach (var j in analyticsModel.GetAllJobLevels())
                clbJobLevelFilter.Items.Add(j, false);

            clbContractTypeFilter.Items.Clear();
            foreach (var c in analyticsModel.GetAllContractTypes())
                clbContractTypeFilter.Items.Add(c, false);

            clbCompanyFilter.Items.Clear();
            foreach (var co in analyticsModel.GetAllCompanies())
                clbCompanyFilter.Items.Add(co, false);

            clbWorkShiftFilter.Items.Clear();
            foreach (var ws in analyticsModel.GetAllWorkShifts())
                clbWorkShiftFilter.Items.Add(ws, false);
        }

        private void ClbProvincesFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdateCitiesAndAffairs();
                UpdateDepartmentsAndDistricts();
                RefreshAllCharts();
            });
        }

        private void ClbCitiesFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdateDepartmentsAndDistricts();
                RefreshAllCharts();
            });
        }

        private void ClbAffairsFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdateDepartmentsAndDistricts();
                RefreshAllCharts();
            });
        }

        private void ClbDepartmentsFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdateDistrictsAndPositions();
                RefreshAllCharts();
            });
        }

        private void ClbDistrictsFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdatePositions();
                RefreshAllCharts();
            });
        }

        private void ClbPositionsFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbGenderFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbEducationFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbJobLevelFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbContractTypeFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbCompanyFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbWorkShiftFilter_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ChkHireDateFilter_CheckedChanged(object? sender, EventArgs e)
        {
            if (dtpHireDateFrom != null) dtpHireDateFrom.Enabled = chkHireDateFilter?.Checked ?? false;
            if (dtpHireDateTo != null) dtpHireDateTo.Enabled = chkHireDateFilter?.Checked ?? false;
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void UpdateFilters()
        {
            List<string> selectedProvinces = clbProvincesFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedCities = clbCitiesFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedAffairs = clbAffairsFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedDepts = clbDepartmentsFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedDistricts = clbDistrictsFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedPositions = clbPositionsFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedGenders = clbGenderFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedEducations = clbEducationFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedJobLevels = clbJobLevelFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedContractTypes = clbContractTypeFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedCompanies = clbCompanyFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedWorkShifts = clbWorkShiftFilter.CheckedItems.Cast<string>().ToList();

            DateTime? hireFromDate = (chkHireDateFilter?.Checked ?? false) ? dtpHireDateFrom?.Value : null;
            DateTime? hireToDate = (chkHireDateFilter?.Checked ?? false) ? dtpHireDateTo?.Value : null;

            int? minAgeValue = (chkAgeFilter?.Checked ?? false) ? (int?)nudMinAge?.Value : null;
            int? maxAgeValue = (chkAgeFilter?.Checked ?? false) ? (int?)nudMaxAge?.Value : null;

            int? minExpValue = (chkExperienceFilter?.Checked ?? false) ? (int?)nudMinExperience?.Value : null;
            int? maxExpValue = (chkExperienceFilter?.Checked ?? false) ? (int?)nudMaxExperience?.Value : null;

            analyticsModel.SetFilters(selectedProvinces, selectedCities, selectedAffairs, selectedDepts,
                selectedDistricts, selectedPositions, selectedGenders, selectedEducations, selectedJobLevels,
                selectedContractTypes, selectedCompanies, selectedWorkShifts, hireFromDate, hireToDate,
                minAgeValue, maxAgeValue, minExpValue, maxExpValue);

            int filterCount = selectedProvinces.Count + selectedCities.Count + selectedAffairs.Count +
                selectedDepts.Count + selectedDistricts.Count + selectedPositions.Count +
                selectedGenders.Count + selectedEducations.Count + selectedJobLevels.Count +
                selectedContractTypes.Count + selectedCompanies.Count + selectedWorkShifts.Count +
                ((chkHireDateFilter?.Checked ?? false) ? 1 : 0) + ((chkAgeFilter?.Checked ?? false) ? 1 : 0) + ((chkExperienceFilter?.Checked ?? false) ? 1 : 0);

            lblFilterInfo.Text = filterCount > 0 ? $"🔴 {filterCount} فیلتر فعال" : "✓ فیلتری فعال نیست";
        }

        private void UpdateCitiesAndAffairs()
        {
            clbCitiesFilter.Items.Clear();
            clbAffairsFilter.Items.Clear();
            var selectedProvinces = clbProvincesFilter.CheckedItems.Cast<string>().ToList();

            if (selectedProvinces.Count > 0)
            {
                foreach (var city in analyticsModel.GetCitiesByProvinces(selectedProvinces).Distinct().OrderBy(x => x))
                    clbCitiesFilter.Items.Add(city, false);

                foreach (var affair in analyticsModel.GetAffairsByProvinces(selectedProvinces).Distinct().OrderBy(x => x))
                    clbAffairsFilter.Items.Add(affair, false);
            }
            else
            {
                var allProvinces = analyticsModel.GetAllProvinces().Distinct().OrderBy(x => x).ToList();
                foreach (var city in analyticsModel.GetCitiesByProvinces(allProvinces).Distinct().OrderBy(x => x))
                    clbCitiesFilter.Items.Add(city, false);

                foreach (var affair in analyticsModel.GetAffairsByProvinces(allProvinces).Distinct().OrderBy(x => x))
                    clbAffairsFilter.Items.Add(affair, false);
            }
        }

        private void UpdateDepartmentsAndDistricts()
        {
            clbDepartmentsFilter.Items.Clear();
            clbDistrictsFilter.Items.Clear();

            var selectedProvinces = clbProvincesFilter.CheckedItems.Cast<string>().ToList();
            var selectedCities = clbCitiesFilter.CheckedItems.Cast<string>().ToList();
            var selectedAffairs = clbAffairsFilter.CheckedItems.Cast<string>().ToList();

            if (selectedProvinces.Count == 0)
                selectedProvinces = analyticsModel.GetAllProvinces().Distinct().OrderBy(x => x).ToList();

            if (selectedCities.Count == 0)
                selectedCities = clbCitiesFilter.Items.Cast<string>().ToList();

            if (selectedAffairs.Count == 0)
                selectedAffairs = clbAffairsFilter.Items.Cast<string>().ToList();

            var depts = analyticsModel.GetDepartmentsByFilters(selectedProvinces, selectedCities, selectedAffairs)
                .Distinct().OrderBy(x => x).ToList();

            foreach (var dept in depts)
                clbDepartmentsFilter.Items.Add(dept, false);

            if (depts.Count > 0)
            {
                foreach (var district in analyticsModel.GetDistrictsByDepartments(depts).Distinct().OrderBy(x => x))
                    clbDistrictsFilter.Items.Add(district, false);
            }
        }

        private void UpdateDistrictsAndPositions()
        {
            clbDistrictsFilter.Items.Clear();
            var selectedDepts = clbDepartmentsFilter.CheckedItems.Cast<string>().ToList();

            if (selectedDepts.Count > 0)
            {
                foreach (var district in analyticsModel.GetDistrictsByDepartments(selectedDepts).Distinct().OrderBy(x => x))
                    clbDistrictsFilter.Items.Add(district, false);
            }
            else
            {
                var allDepts = clbDepartmentsFilter.Items.Cast<string>().ToList();
                if (allDepts.Count > 0)
                {
                    foreach (var district in analyticsModel.GetDistrictsByDepartments(allDepts).Distinct().OrderBy(x => x))
                        clbDistrictsFilter.Items.Add(district, false);
                }
            }
        }

        private void UpdatePositions()
        {
            clbPositionsFilter.Items.Clear();
            var selectedDistricts = clbDistrictsFilter.CheckedItems.Cast<string>().ToList();

            if (selectedDistricts.Count > 0)
            {
                foreach (var pos in analyticsModel.GetPositionsByDistricts(selectedDistricts).Distinct().OrderBy(x => x))
                    clbPositionsFilter.Items.Add(pos, false);
            }
            else
            {
                var allDistricts = clbDistrictsFilter.Items.Cast<string>().ToList();
                if (allDistricts.Count > 0)
                {
                    foreach (var pos in analyticsModel.GetPositionsByDistricts(allDistricts).Distinct().OrderBy(x => x))
                        clbPositionsFilter.Items.Add(pos, false);
                }
            }
        }

        private void BtnClearFilters_Click(object? sender, EventArgs e)
        {
            for (int i = 0; i < clbProvincesFilter.Items.Count; i++) clbProvincesFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbCitiesFilter.Items.Count; i++) clbCitiesFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbAffairsFilter.Items.Count; i++) clbAffairsFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbDepartmentsFilter.Items.Count; i++) clbDepartmentsFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbDistrictsFilter.Items.Count; i++) clbDistrictsFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbPositionsFilter.Items.Count; i++) clbPositionsFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbGenderFilter.Items.Count; i++) clbGenderFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbEducationFilter.Items.Count; i++) clbEducationFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbJobLevelFilter.Items.Count; i++) clbJobLevelFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbContractTypeFilter.Items.Count; i++) clbContractTypeFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbCompanyFilter.Items.Count; i++) clbCompanyFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbWorkShiftFilter.Items.Count; i++) clbWorkShiftFilter.SetItemChecked(i, false);
            if (chkHireDateFilter != null) chkHireDateFilter.Checked = false;
            if (chkAgeFilter != null) chkAgeFilter.Checked = false;
            if (chkExperienceFilter != null) chkExperienceFilter.Checked = false;

            analyticsModel.ClearFilters();
            lblFilterInfo.Text = "✓ فیلتری فعال نیست";
            LoadFilterOptions();
            RefreshAllCharts();
        }

        private void RefreshAllCharts()
        {
            if (rbShowSummary?.Checked ?? true)
                LoadSummaryTable();
            else
                LoadStatisticalTable();

            LoadDepartmentPieChart();
            LoadPositionPieChart();
            LoadGenderPieChart();
            LoadJobLevelPieChart();
            LoadContractTypePieChart();
            LoadProvincePieChart();
            LoadEducationPieChart();
            LoadCompanyPieChart();
            LoadWorkShiftPieChart();
            LoadAgePieChart();
            LoadWorkExperiencePieChart();

            ReapplyChartTypes();
        }

        private void LoadSummaryTable()
        {
            try
            {
                dgvPersonnelStats.DataSource = null;
                dgvPersonnelStats.Columns.Clear();
                dgvPersonnelStats.Columns.Add("Metric", "معیار");
                dgvPersonnelStats.Columns.Add("Value", "مقدار");

                dgvPersonnelStats.Rows.Add("👥 کل پرسنل", analyticsModel.GetFilteredTotal());
                dgvPersonnelStats.Rows.Add("🏛️ تعداد ادارهها", analyticsModel.GetFilteredDepartmentCount());
                dgvPersonnelStats.Rows.Add("💼 تعداد پستهای شغلی", analyticsModel.GetFilteredPositionCount());
                dgvPersonnelStats.Rows.Add("🗺️ تعداد استانها", analyticsModel.ProvinceCount);
                dgvPersonnelStats.Rows.Add("🏢 تعداد شرکتها", analyticsModel.CompanyCount);
                dgvPersonnelStats.Rows.Add("📈 تعداد سطحهای شغلی", analyticsModel.JobLevelCount);
                dgvPersonnelStats.Rows.Add("📋 تعداد انواع قرارداد", analyticsModel.ContractTypeCount);
                dgvPersonnelStats.Rows.Add("📚 تعداد مدارک تحصیلی", analyticsModel.EducationCount);
                dgvPersonnelStats.Rows.Add("⏰ تعداد شیفت‌های کاری", analyticsModel.WorkShiftCount);
                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("👩 افراد خانم", analyticsModel.GetFilteredFemaleCount());
                dgvPersonnelStats.Rows.Add("👨 افراد آقا", analyticsModel.GetFilteredMaleCount());
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadDepartmentPieChart()
        {
            try
            {
                chartDepartmentPie.Series.Clear();
                
                // 🎯 Safe font with fallback
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredDepartmentStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartDepartmentPie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartDepartmentPie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                var items = pie ? stats.Take(15).ToList() : stats.ToList();
                foreach (var item in items)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartDepartmentPie.Series.Add(series);
                chartDepartmentPie.Titles.Clear();
                chartDepartmentPie.Titles.Add(new Title("📊 توزیع پرسنل در ادارهها") { Font = safeTitleFont });

                dgvDepartmentDetails.DataSource = null;
                dgvDepartmentDetails.Columns.Clear();
                dgvDepartmentDetails.Columns.Add("Name", "اداره");
                dgvDepartmentDetails.Columns.Add("Count", "تعداد");
                dgvDepartmentDetails.Columns.Add("Percent", "درصد");
                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    dgvDepartmentDetails.Rows.Add(item.Name, item.Count, $"{pct:F1}%");
                }
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadPositionPieChart()
        {
            try
            {
                chartPositionPie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredPositionStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartPositionPie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartPositionPie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                var items = pie ? stats.Take(15).ToList() : stats.ToList();
                foreach (var item in items)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartPositionPie.Series.Add(series);
                chartPositionPie.Titles.Clear();
                chartPositionPie.Titles.Add(new Title("💼 توزیع پستهای شغلی") { Font = safeTitleFont });

                dgvPositionDetails.DataSource = null;
                dgvPositionDetails.Columns.Clear();
                dgvPositionDetails.Columns.Add("Name", "پست");
                dgvPositionDetails.Columns.Add("Count", "تعداد");
                dgvPositionDetails.Columns.Add("Percent", "درصد");
                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    dgvPositionDetails.Rows.Add(item.Name, item.Count, $"{pct:F1}%");
                }
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadGenderPieChart()
        {
            try
            {
                chartGenderPie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredGenderStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartGenderPie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartGenderPie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartGenderPie.Series.Add(series);
                chartGenderPie.Titles.Clear();
                chartGenderPie.Titles.Add(new Title("👥 توزیع جنسیت") { Font = safeTitleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadJobLevelPieChart()
        {
            try
            {
                chartJobLevelPie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredJobLevelStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartJobLevelPie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartJobLevelPie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartJobLevelPie.Series.Add(series);
                chartJobLevelPie.Titles.Clear();
                chartJobLevelPie.Titles.Add(new Title("📈 توزیع سطح شغلی") { Font = safeTitleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadContractTypePieChart()
        {
            try
            {
                chartContractTypePie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredContractTypeStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartContractTypePie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartContractTypePie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartContractTypePie.Series.Add(series);
                chartContractTypePie.Titles.Clear();
                chartContractTypePie.Titles.Add(new Title("📋 توزیع نوع قرارداد") { Font = safeTitleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadProvincePieChart()
        {
            try
            {
                chartProvincePie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredProvinceStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartProvincePie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartProvincePie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                var items = pie ? stats.Take(20).ToList() : stats.ToList();
                foreach (var item in items)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartProvincePie.Series.Add(series);
                chartProvincePie.Titles.Clear();
                chartProvincePie.Titles.Add(new Title("🗺️ توزیع بر اساس استان") { Font = safeTitleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadEducationPieChart()
        {
            try
            {
                chartEducationPie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredEducationStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartEducationPie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartEducationPie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartEducationPie.Series.Add(series);
                chartEducationPie.Titles.Clear();
                chartEducationPie.Titles.Add(new Title("📚 توزیع مدارک تحصیلی") { Font = safeTitleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadCompanyPieChart()
        {
            try
            {
                chartCompanyPie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredCompanyStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartCompanyPie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartCompanyPie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartCompanyPie.Series.Add(series);
                chartCompanyPie.Titles.Clear();
                chartCompanyPie.Titles.Add(new Title("🏢 توزیع شرکتها") { Font = safeTitleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadWorkShiftPieChart()
        {
            try
            {
                chartWorkShiftPie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredWorkShiftStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartWorkShiftPie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartWorkShiftPie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartWorkShiftPie.Series.Add(series);
                chartWorkShiftPie.Titles.Clear();
                chartWorkShiftPie.Titles.Add(new Title("⏰ توزیع شیفت‌های کاری") { Font = safeTitleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadAgePieChart()
        {
            try
            {
                chartAgePie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                int rangeSize = nudAgeRangeSize != null ? (int)nudAgeRangeSize.Value : 10;
                var stats = analyticsModel.GetFilteredAgeStatistics(rangeSize);
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartAgePie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartAgePie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartAgePie.Series.Add(series);
                chartAgePie.Titles.Clear();
                chartAgePie.Titles.Add(new Title($"🎂 توزیع بر اساس سن (بازه: {rangeSize} سال)") { Font = safeTitleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadWorkExperiencePieChart()
        {
            try
            {
                chartWorkExperiencePie.Series.Clear();
                
                Font safeFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font safeTitleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 11F, FontStyle.Bold);
                
                var stats = analyticsModel.GetFilteredWorkExperienceStatistics();
                int total = stats.Sum(x => x.Count);

                var type = GetChartTypeOrDefault(chartWorkExperiencePie);
                bool pie = IsPieType(type);
                ConfigureChartAreaForType(chartWorkExperiencePie, type);

                Series series = new Series("تعداد");
                ConfigureSeriesForType(series, type);

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].AxisLabel = item.Name;
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].Font = safeFont;

                    if (pie)
                        series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    else
                        series.Points[idx].Label = item.Count.ToString();
                }

                chartWorkExperiencePie.Series.Add(series);
                chartWorkExperiencePie.Titles.Clear();
                chartWorkExperiencePie.Titles.Add(new Title("💼 توزیع بر اساس سابقه کاری") { Font = safeTitleFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void Chart_MouseClick(object? sender, MouseEventArgs e)
        {
            try
            {
                Chart? chart = sender as Chart;
                if (chart == null) return;

                HitTestResult result = chart.HitTest(e.X, e.Y);
                if (result.ChartElementType == ChartElementType.DataPoint)
                {
                    int pointIndex = result.PointIndex;
                    DataPoint point = result.Series.Points[pointIndex];

                    string itemName = point.AxisLabel;

                    List<PersonnelDetail> personnel;
                    if (chart == chartAgePie)
                    {
                        int rangeSize = nudAgeRangeSize != null ? (int)nudAgeRangeSize.Value : 10;
                        personnel = analyticsModel.GetPersonnelByFilter(itemName, chart, rangeSize);
                    }
                    else
                    {
                        personnel = analyticsModel.GetPersonnelByFilter(itemName, chart);
                    }

                    if (personnel.Count > 0)
                        ShowPersonnelDetails(itemName, personnel);
                    else
                        MessageBox.Show("❌ داده‌ای برای نمایش وجود ندارد.", "پیام", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void ShowPersonnelDetails(string category, List<PersonnelDetail> personnel)
        {
            Form detailsForm = new Form
            {
                Text = $"👥 جزئیات پرسنل - {category}",
                Size = new Size(1400, 800),
                StartPosition = FormStartPosition.CenterScreen,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true,
                BackColor = Color.FromArgb(240, 248, 255),
                WindowState = FormWindowState.Maximized,
                Font = FontSettings.BodyFont
            };

            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                ReadOnly = false,
                RightToLeft = RightToLeft.Yes,
                BackgroundColor = Color.White,
                EnableHeadersVisualStyles = false,
                AllowUserToAddRows = false,
                ColumnHeadersHeight = 40,
                RowTemplate = { Height = 35 }
            };

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
            dgv.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            dgv.Columns.Add("PersonnelID", "ID");
            dgv.Columns["PersonnelID"].Visible = false;
            dgv.Columns.Add("FirstName", "نام");
            dgv.Columns.Add("LastName", "نام‌خانوادگی");
            dgv.Columns.Add("PersonnelNumber", "شماره پرسنلی");
            dgv.Columns.Add("NationalID", "شناسه ملی");
            dgv.Columns.Add("PostName", "پست");
            dgv.Columns.Add("DeptName", "اداره");
            dgv.Columns.Add("Province", "استان");
            dgv.Columns.Add("ContractType", "نوع قرارداد");
            dgv.Columns.Add("HireDate", "تاریخ استخدام");
            dgv.Columns.Add("MobileNumber", "تلفن");

            DataGridViewButtonColumn editColumn = new DataGridViewButtonColumn
            {
                Name = "Edit",
                HeaderText = "✏️ ویرایش",
                Text = "ویرایش",
                UseColumnTextForButtonValue = true,
                Width = 120,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(40, 167, 69),
                    ForeColor = Color.White,
                    Font = FontSettings.ButtonFont,
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    Padding = new Padding(5)
                }
            };
            dgv.Columns.Add(editColumn);

            DataGridViewButtonColumn deleteColumn = new DataGridViewButtonColumn
            {
                Name = "Delete",
                HeaderText = "🗑️ حذف",
                Text = "حذف",
                UseColumnTextForButtonValue = true,
                Width = 120,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(220, 53, 69),
                    ForeColor = Color.White,
                    Font = FontSettings.ButtonFont,
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    Padding = new Padding(5)
                }
            };
            dgv.Columns.Add(deleteColumn);

            foreach (var p in personnel)
            {
                dgv.Rows.Add(p.PersonnelID, p.FirstName, p.LastName, p.PersonnelNumber, p.NationalID, p.PostName,
                    p.DeptName, p.Province, p.ContractType, p.HireDate?.ToString("yyyy/MM/dd") ?? "", p.MobileNumber, "ویرایش", "حذف");
            }

            dgv.CellClick += (sender, e) =>
            {
                if (e.ColumnIndex == dgv.Columns["Edit"].Index && e.RowIndex >= 0)
                {
                    int personnelID = Convert.ToInt32(dgv.Rows[e.RowIndex].Cells["PersonnelID"].Value);
                    OpenEditForm(personnelID, detailsForm);
                }
                else if (e.ColumnIndex == dgv.Columns["Delete"].Index && e.RowIndex >= 0)
                {
                    int personnelID = Convert.ToInt32(dgv.Rows[e.RowIndex].Cells["PersonnelID"].Value);
                    DeletePersonnel(personnelID, detailsForm, dgv, e.RowIndex);
                }
            };

            Panel bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 70,
                BackColor = Color.FromArgb(230, 240, 250)
            };

            Button btnExportExcel = new Button
            {
                Text = "📊 خروجی اکسل",
                Location = new Point(20, 15),
                Size = new Size(200, 40),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                Font = FontSettings.ButtonFont,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnExportExcel.FlatAppearance.BorderSize = 0;
            btnExportExcel.Click += (s, ev) =>
            {
                using (ExportColumnsForm exportForm = new ExportColumnsForm())
                {
                    if (exportForm.ShowDialog() == DialogResult.OK)
                    {
                        var selectedColumns = exportForm.SelectedColumns;
                        if (selectedColumns != null && selectedColumns.Count > 0)
                        {
                            ExcelExportHelper.ExportToExcel(personnel, selectedColumns, $"Personnel_{category}");
                        }
                    }
                }
            };
            bottomPanel.Controls.Add(btnExportExcel);

            detailsForm.Controls.Add(dgv);
            detailsForm.Controls.Add(bottomPanel);
            detailsForm.ShowDialog();
        }

        private void OpenEditForm(int personnelID, Form parentForm)
        {
            try
            {
                FormPersonnelEdit editForm = new FormPersonnelEdit();
                editForm.txtPersonnelID.Text = personnelID.ToString();
                editForm.BtnLoad_Click(null, EventArgs.Empty);
                editForm.ShowDialog(parentForm);

                RefreshAllCharts();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در باز کردن فرم ویرایش: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeletePersonnel(int personnelID, Form parentForm, DataGridView dgv, int rowIndex)
        {
            try
            {
                DialogResult result = MessageBox.Show(
                    $"❓ آیا مطمئن هستید که می‌خواهید این پرسنل را حذف کنید؟",
                    "تایید حذف",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string query = $"DELETE FROM Personnel WHERE PersonnelID = {personnelID}";
                    dbHelper.ExecuteNonQuery(query);

                    MessageBox.Show("✅ پرسنل با موفقیت حذف شد.", "موفق", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    dgv.Rows.RemoveAt(rowIndex);

                    LoadData();
                    RefreshAllCharts();

                    if (dgv.Rows.Count == 0)
                    {
                        parentForm.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در حذف پرسنل: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadStatisticalTable()
        {
            try
            {
                dgvPersonnelStats.DataSource = null;
                dgvPersonnelStats.Columns.Clear();
                dgvPersonnelStats.Columns.Add("Metric", "معیار");
                dgvPersonnelStats.Columns.Add("Value", "مقدار");

                dgvPersonnelStats.Rows.Add("═══════════════════", "");
                dgvPersonnelStats.Rows.Add("👥 کل پرسنل", analyticsModel.GetFilteredTotal());
                dgvPersonnelStats.Rows.Add("🏛️ تعداد ادارهها", analyticsModel.GetFilteredDepartmentCount());
                dgvPersonnelStats.Rows.Add("💼 تعداد پستهای شغلی", analyticsModel.GetFilteredPositionCount());
                dgvPersonnelStats.Rows.Add("🗺️ تعداد استانها", analyticsModel.ProvinceCount);
                dgvPersonnelStats.Rows.Add("🏢 تعداد شرکتها", analyticsModel.CompanyCount);
                dgvPersonnelStats.Rows.Add("📈 تعداد سطحهای شغلی", analyticsModel.JobLevelCount);
                dgvPersonnelStats.Rows.Add("📋 تعداد انواع قرارداد", analyticsModel.ContractTypeCount);
                dgvPersonnelStats.Rows.Add("📚 تعداد مدارک تحصیلی", analyticsModel.EducationCount);
                dgvPersonnelStats.Rows.Add("⏰ تعداد شیفت‌های کاری", analyticsModel.WorkShiftCount);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("═════ توزیع جنسیت ═════", "");
                foreach (var g in analyticsModel.GetFilteredGenderStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {g.Name}", g.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════ توزیع سطح شغلی ════", "");
                foreach (var j in analyticsModel.GetFilteredJobLevelStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {j.Name}", j.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════ توزیع نوع قرارداد ════", "");
                foreach (var c in analyticsModel.GetFilteredContractTypeStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {c.Name}", c.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("═════════ تمام ادارات ═════════", "");
                foreach (var d in analyticsModel.GetFilteredDepartmentStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {d.Name}", d.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════════ تمام پستهای شغلی ════════", "");
                foreach (var p in analyticsModel.GetFilteredPositionStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {p.Name}", p.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════════════ تمام استان‌ها ════════════", "");
                foreach (var pr in analyticsModel.GetFilteredProvinceStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {pr.Name}", pr.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════════════ تمام شرکتها ════════════", "");
                foreach (var co in analyticsModel.GetFilteredCompanyStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {co.Name}", co.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("═════════ تمام مدارک تحصیلی ═════════", "");
                foreach (var e in analyticsModel.GetFilteredEducationStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {e.Name}", e.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("═════════ تمام شیفت‌های کاری ═════════", "");
                foreach (var ws in analyticsModel.GetFilteredWorkShiftStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {ws.Name}", ws.Count);
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }
    }
}
