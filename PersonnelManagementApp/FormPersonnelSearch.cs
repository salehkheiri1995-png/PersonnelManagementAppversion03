using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    /// <summary>
    /// جستجو و حذف پرسنل (رابط مدرن).
    /// </summary>
    public partial class FormPersonnelSearch : Form
    {
        private static readonly Color HeaderColor = Color.FromArgb(15, 23, 42);
        private static readonly Color HeaderAccent = Color.FromArgb(56, 189, 248);
        private static readonly Color CardColor = Color.White;
        private static readonly Color CanvasColor = Color.FromArgb(241, 245, 249);
        private static readonly Color BorderColor = Color.FromArgb(226, 232, 240);
        private static readonly Color PrimaryColor = Color.FromArgb(0, 122, 204);
        private static readonly Color SuccessColor = Color.FromArgb(16, 185, 129);
        private static readonly Color DangerColor = Color.FromArgb(239, 68, 68);
        private static readonly Color WarningColor = Color.FromArgb(245, 158, 11);
        private static readonly Color TextColor = Color.FromArgb(30, 41, 59);
        private static readonly Color MutedColor = Color.FromArgb(100, 116, 139);

        private readonly DbHelper db = new();

        private DataTable? personnelTable;
        private DataTable? provincesTable;
        private DataTable? citiesTable;
        private DataTable? transferAffairsTable;
        private DataTable? operationDepartmentsTable;
        private DataTable? districtsTable;
        private DataTable? postsNamesTable;
        private DataTable? searchTablePersonnel;

        private TextBox txtFreeSearch = new();
        private CheckedListBox clbProvinces = new();
        private CheckedListBox clbCities = new();
        private CheckedListBox clbAffairs = new();
        private CheckedListBox clbDepartments = new();
        private CheckedListBox clbDistricts = new();
        private CheckedListBox clbPosts = new();
        private DataGridView dgvResults = new();
        private Label lblStatus = new();
        private Label lblCount = new();

        private bool _loading;

        public FormPersonnelSearch()
        {
            BuildUi();
            FontSettings.ApplyFontToForm(this);
            LoadData();
            CreateSearchTable();
            LoadProvinces();
            ApplyResults();
        }

        private void BuildUi()
        {
            Text = "جستجو و حذف پرسنل";
            WindowState = FormWindowState.Maximized;
            MinimumSize = new Size(1100, 700);
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            BackColor = CanvasColor;
            Font = new Font(FontSettings.FontFamilyName, 9F);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                BackColor = CanvasColor,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 86F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
            Controls.Add(root);

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildToolbar(), 0, 1);
            root.Controls.Add(BuildContent(), 0, 2);
            root.Controls.Add(BuildStatusBar(), 0, 3);
        }

        private Control BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = HeaderColor,
                Padding = new Padding(24, 12, 24, 10),
                Margin = new Padding(0)
            };

            var title = new Label
            {
                Text = "جستجو و حذف پرسنل",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font(FontSettings.FontFamilyName, 17F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var subtitle = new Label
            {
                Text = "جستجوی سریع بر اساس نام، مشخصات و موقعیت سازمانی — به‌همراه حذف پرسنل انتخابی",
                Dock = DockStyle.Fill,
                Font = new Font(FontSettings.FontFamilyName, 9.5F),
                ForeColor = Color.FromArgb(203, 213, 225),
                TextAlign = ContentAlignment.MiddleLeft
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

        private Control BuildToolbar()
        {
            var toolbar = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Padding = new Padding(12, 8, 12, 8),
                Margin = new Padding(0)
            };

            txtFreeSearch.Dock = DockStyle.Fill;
            txtFreeSearch.Font = new Font(FontSettings.FontFamilyName, 9.5F);
            txtFreeSearch.PlaceholderText = "جستجوی آزاد (نام، نام خانوادگی، شماره پرسنلی، کد ملی)...";
            txtFreeSearch.KeyDown += (_, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    ApplyResults();
                    e.SuppressKeyPress = true;
                }
            };

            var btnSearch = CreateToolButton("🔎 جستجو", PrimaryColor);
            btnSearch.Click += (_, _) => ApplyResults();

            var btnClear = CreateToolButton("🧹 پاک کردن فیلترها", Color.FromArgb(71, 85, 105));
            btnClear.Click += (_, _) => ClearFilters();

            var btnExport = CreateToolButton("📥 خروجی CSV", Color.FromArgb(20, 184, 166));
            btnExport.Click += (_, _) => ExportCsv();

            var btnDelete = CreateToolButton("🗑️ حذف پرسنل انتخابی", DangerColor);
            btnDelete.Click += (_, _) => DeleteSelected();

            var btnRefresh = CreateToolButton("🔄 به‌روزرسانی", Color.FromArgb(139, 92, 246));
            btnRefresh.Click += (_, _) =>
            {
                LoadData();
                CreateSearchTable();
                LoadProvinces();
                ApplyResults();
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = CardColor,
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 830F));

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = CardColor,
                Margin = new Padding(0)
            };
            actions.Controls.Add(btnRefresh);
            actions.Controls.Add(btnDelete);
            actions.Controls.Add(btnExport);
            actions.Controls.Add(btnClear);
            actions.Controls.Add(btnSearch);

            layout.Controls.Add(txtFreeSearch, 0, 0);
            layout.Controls.Add(actions, 1, 0);
            toolbar.Controls.Add(layout);
            return toolbar;
        }

        private static Button CreateToolButton(string text, Color backColor)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = false,
                Width = 160,
                Height = 36,
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

        private Control BuildContent()
        {
            var content = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = CanvasColor,
                Padding = new Padding(12),
                Margin = new Padding(0)
            };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 320F));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            content.Controls.Add(BuildFilterPanel(), 0, 0);
            content.Controls.Add(BuildGridPanel(), 1, 0);
            return content;
        }

        private Control BuildFilterPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Padding = new Padding(12, 10, 12, 10),
                Margin = new Padding(0, 0, 10, 0),
                BorderStyle = BorderStyle.FixedSingle
            };

            var title = new Label
            {
                Text = "فیلترهای سازمانی",
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font(FontSettings.FontFamilyName, 11F, FontStyle.Bold),
                ForeColor = PrimaryColor,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var scroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = CardColor
            };

            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                BackColor = CardColor,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };

            AddFilterGroup(flow, "استان‌ها", clbProvinces, true, ClbProvinces_ItemCheck);
            AddFilterGroup(flow, "شهرها", clbCities, false, ClbCities_ItemCheck);
            AddFilterGroup(flow, "امور", clbAffairs, false, ClbAffairs_ItemCheck);
            AddFilterGroup(flow, "ادارات", clbDepartments, false, ClbDepartments_ItemCheck);
            AddFilterGroup(flow, "نواحی", clbDistricts, false, ClbDistricts_ItemCheck);
            AddFilterGroup(flow, "پست‌ها", clbPosts, false, null);

            scroll.Controls.Add(flow);
            panel.Controls.Add(scroll);
            panel.Controls.Add(title);
            return panel;
        }

        private static void AddFilterGroup(FlowLayoutPanel flow, string caption, CheckedListBox list,
            bool enabled, ItemCheckEventHandler? onCheck)
        {
            var group = new Panel
            {
                Width = 258,
                Height = 140,
                Margin = new Padding(0, 0, 0, 10),
                Padding = new Padding(8, 4, 8, 4),
                BackColor = Color.FromArgb(248, 250, 252),
                BorderStyle = BorderStyle.FixedSingle
            };

            var header = new Label
            {
                Text = caption,
                Dock = DockStyle.Top,
                Height = 24,
                Font = new Font(FontSettings.FontFamilyName, 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 64, 175),
                TextAlign = ContentAlignment.MiddleLeft
            };

            list.Dock = DockStyle.Fill;
            list.CheckOnClick = true;
            list.BorderStyle = BorderStyle.None;
            list.BackColor = Color.FromArgb(248, 250, 252);
            list.Font = new Font(FontSettings.FontFamilyName, 8.5F);
            list.IntegralHeight = false;
            list.Enabled = enabled;
            if (onCheck != null)
                list.ItemCheck += onCheck;

            group.Controls.Add(list);
            group.Controls.Add(header);
            flow.Controls.Add(group);
        }

        private Control BuildGridPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Margin = new Padding(0)
            };

            dgvResults.Dock = DockStyle.Fill;
            dgvResults.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvResults.AllowUserToAddRows = false;
            dgvResults.ReadOnly = true;
            dgvResults.BackgroundColor = Color.White;
            dgvResults.BorderStyle = BorderStyle.None;
            dgvResults.ColumnHeadersHeight = 40;
            dgvResults.RowTemplate.Height = 30;
            dgvResults.RowHeadersVisible = false;
            dgvResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgvResults.MultiSelect = false;
            dgvResults.RightToLeft = RightToLeft.Yes;
            dgvResults.EnableHeadersVisualStyles = false;
            dgvResults.GridColor = BorderColor;
            dgvResults.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 64, 175);
            dgvResults.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvResults.ColumnHeadersDefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F, FontStyle.Bold);
            dgvResults.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvResults.DefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F);
            dgvResults.DefaultCellStyle.ForeColor = TextColor;
            dgvResults.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            dgvResults.DefaultCellStyle.SelectionForeColor = TextColor;
            dgvResults.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);

            panel.Controls.Add(dgvResults);
            return panel;
        }

        private Control BuildStatusBar()
        {
            var statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(248, 250, 252),
                Margin = new Padding(0)
            };
            statusPanel.Paint += (_, e) =>
            {
                using var pen = new Pen(BorderColor);
                e.Graphics.DrawLine(pen, 0, 0, statusPanel.Width, 0);
            };

            lblStatus.Dock = DockStyle.Fill;
            lblStatus.Font = new Font(FontSettings.FontFamilyName, 8.5F);
            lblStatus.ForeColor = Color.FromArgb(71, 85, 105);
            lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            lblStatus.Padding = new Padding(18, 0, 18, 0);
            lblStatus.Text = "در حال بارگذاری...";

            lblCount.Dock = DockStyle.Right;
            lblCount.Width = 220;
            lblCount.Font = new Font(FontSettings.FontFamilyName, 9F, FontStyle.Bold);
            lblCount.ForeColor = PrimaryColor;
            lblCount.TextAlign = ContentAlignment.MiddleRight;
            lblCount.Padding = new Padding(0, 0, 12, 0);
            lblCount.Text = "۰ پرسنل";

            statusPanel.Controls.Add(lblStatus);
            statusPanel.Controls.Add(lblCount);
            return statusPanel;
        }

        // ─────────────────────────────────────────────
        // داده‌ها
        // ─────────────────────────────────────────────
        private void LoadData()
        {
            try
            {
                provincesTable = db.GetProvinces();
                citiesTable = db.GetCitiesByProvince(0);
                transferAffairsTable = db.GetAffairsByProvince(0);
                operationDepartmentsTable = db.GetDeptsByAffair(0);
                districtsTable = db.GetDistrictsByDept(0);
                postsNamesTable = db.GetPostNamesByDistrict(0);
                personnelTable = db.ExecuteQuery(
                    "SELECT PersonnelID, ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, VoltageID, " +
                    "WorkShiftID, GenderID, FirstName, LastName, FatherName, PersonnelNumber, NationalID, " +
                    "MobileNumber, BirthDate, HireDate, StartDateOperation, ContractTypeID, JobLevelID, " +
                    "CompanyID, DegreeID, DegreeFieldID, MainJobTitle, CurrentActivity, Inconsistency, " +
                    "Description, StatusID FROM Personnel");

                if (personnelTable == null)
                    personnelTable = new DataTable("Personnel");

                lblStatus.Text = $"✅ {personnelTable.Rows.Count} پرسنل بارگذاری شد.";
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPersonnelSearch.LoadData", ex);
                lblStatus.Text = "❌ خطا در بارگذاری داده‌ها.";
                MessageBox.Show(this, $"خطا در بارگذاری داده‌ها:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CreateSearchTable()
        {
            searchTablePersonnel = new DataTable("SearchTablePersonnel");
            searchTablePersonnel.Columns.Add("شناسه پرسنل", typeof(int));
            searchTablePersonnel.Columns.Add("نام", typeof(string));
            searchTablePersonnel.Columns.Add("نام خانوادگی", typeof(string));
            searchTablePersonnel.Columns.Add("نام پدر", typeof(string));
            searchTablePersonnel.Columns.Add("شماره پرسنلی", typeof(string));
            searchTablePersonnel.Columns.Add("کد ملی", typeof(string));
            searchTablePersonnel.Columns.Add("شماره موبایل", typeof(string));
            searchTablePersonnel.Columns.Add("استان", typeof(string));
            searchTablePersonnel.Columns.Add("شهر", typeof(string));
            searchTablePersonnel.Columns.Add("امور", typeof(string));
            searchTablePersonnel.Columns.Add("اداره", typeof(string));
            searchTablePersonnel.Columns.Add("ناحیه", typeof(string));
            searchTablePersonnel.Columns.Add("نام پست", typeof(string));
            searchTablePersonnel.Columns.Add("سطح ولتاژ", typeof(string));
            searchTablePersonnel.Columns.Add("شیفت کاری", typeof(string));
            searchTablePersonnel.Columns.Add("جنسیت", typeof(string));
            searchTablePersonnel.Columns.Add("نوع قرارداد", typeof(string));
            searchTablePersonnel.Columns.Add("سطح شغل", typeof(string));
            searchTablePersonnel.Columns.Add("شرکت", typeof(string));
            searchTablePersonnel.Columns.Add("مدرک تحصیلی", typeof(string));
            searchTablePersonnel.Columns.Add("رشته تحصیلی", typeof(string));
            searchTablePersonnel.Columns.Add("عنوان شغلی اصلی", typeof(string));
            searchTablePersonnel.Columns.Add("فعالیت فعلی", typeof(string));
            searchTablePersonnel.Columns.Add("مغایرت", typeof(string));
            searchTablePersonnel.Columns.Add("توضیحات", typeof(string));
            searchTablePersonnel.Columns.Add("وضعیت حضور", typeof(string));
            searchTablePersonnel.Columns.Add("تاریخ تولد", typeof(string));
            searchTablePersonnel.Columns.Add("تاریخ استخدام", typeof(string));
            searchTablePersonnel.Columns.Add("تاریخ شروع بکار", typeof(string));

            DataTable? voltageLevelsTable = db.GetVoltageLevels();
            DataTable? workShiftsTable = db.GetWorkShifts();
            DataTable? gendersTable = db.GetGenders();
            DataTable? contractTypesTable = db.GetContractTypes();
            DataTable? jobLevelsTable = db.GetJobLevels();
            DataTable? companiesTable = db.GetCompanies();
            DataTable? degreesTable = db.GetDegrees();
            DataTable? degreeFieldsTable = db.GetDegreeFields();
            DataTable? chartAffairsTable = db.GetChartAffairs();
            DataTable? statusPresenceTable = db.GetStatusPresence();

            foreach (DataRow p in personnelTable!.Rows)
            {
                DataRow r = searchTablePersonnel.NewRow();
                r["شناسه پرسنل"] = p["PersonnelID"];
                r["نام"] = p["FirstName"];
                r["نام خانوادگی"] = p["LastName"];
                r["نام پدر"] = p["FatherName"];
                r["شماره پرسنلی"] = p["PersonnelNumber"];
                r["کد ملی"] = p["NationalID"];
                r["شماره موبایل"] = p["MobileNumber"];
                r["تاریخ تولد"] = p["BirthDate"];
                r["تاریخ استخدام"] = p["HireDate"];
                r["تاریخ شروع بکار"] = p["StartDateOperation"];

                r["استان"] = LookupName(provincesTable, "ProvinceID", "ProvinceName", p["ProvinceID"]);
                r["شهر"] = LookupName(citiesTable, "CityID", "CityName", p["CityID"]);
                r["امور"] = LookupName(transferAffairsTable, "AffairID", "AffairName", p["AffairID"]);
                r["اداره"] = LookupName(operationDepartmentsTable, "DeptID", "DeptName", p["DeptID"]);
                r["ناحیه"] = LookupName(districtsTable, "DistrictID", "DistrictName", p["DistrictID"]);
                r["نام پست"] = LookupName(postsNamesTable, "PostNameID", "PostName", p["PostNameID"]);
                r["سطح ولتاژ"] = LookupName(voltageLevelsTable, "VoltageID", "VoltageName", p["VoltageID"]);
                r["شیفت کاری"] = LookupName(workShiftsTable, "WorkShiftID", "WorkShiftName", p["WorkShiftID"]);
                r["جنسیت"] = LookupName(gendersTable, "GenderID", "GenderName", p["GenderID"]);
                r["نوع قرارداد"] = LookupName(contractTypesTable, "ContractTypeID", "ContractTypeName", p["ContractTypeID"]);
                r["سطح شغل"] = LookupName(jobLevelsTable, "JobLevelID", "JobLevelName", p["JobLevelID"]);
                r["شرکت"] = LookupName(companiesTable, "CompanyID", "CompanyName", p["CompanyID"]);
                r["مدرک تحصیلی"] = LookupName(degreesTable, "DegreeID", "DegreeName", p["DegreeID"]);
                r["رشته تحصیلی"] = LookupName(degreeFieldsTable, "DegreeFieldID", "DegreeFieldName", p["DegreeFieldID"]);
                r["عنوان شغلی اصلی"] = LookupName(chartAffairsTable, "ChartID", "ChartName", p["MainJobTitle"]);
                r["فعالیت فعلی"] = LookupName(chartAffairsTable, "ChartID", "ChartName", p["CurrentActivity"]);
                r["وضعیت حضور"] = LookupName(statusPresenceTable, "StatusID", "StatusName", p["StatusID"]);

                r["مغایرت"] = Convert.ToInt32(p["Inconsistency"]) == 1 ? "دارد" : "ندارد";
                r["توضیحات"] = p["Description"];

                searchTablePersonnel.Rows.Add(r);
            }
        }

        private static string LookupName(DataTable? table, string idColumn, string nameColumn, object? idValue)
        {
            if (table == null || idValue == null || idValue == DBNull.Value)
                return "نامشخص";

            int id = Convert.ToInt32(idValue);
            if (id <= 0)
                return "نامشخص";

            DataRow[] rows = table.Select($"{idColumn} = {id}");
            return rows.Length > 0 ? Convert.ToString(rows[0][nameColumn]) ?? "نامشخص" : "نامشخص";
        }

        private void LoadProvinces()
        {
            clbProvinces.Items.Clear();
            if (provincesTable == null)
                return;

            foreach (DataRow row in provincesTable.Rows)
                clbProvinces.Items.Add(Convert.ToString(row["ProvinceName"]) ?? "", false);
        }

        // ─────────────────────────────────────────────
        // فیلترهای وابسته
        // ─────────────────────────────────────────────
        private void ClbProvinces_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)(() =>
            {
                if (_loading)
                    return;

                clbCities.Items.Clear();
                clbAffairs.Items.Clear();
                clbDepartments.Items.Clear();
                clbDistricts.Items.Clear();
                clbPosts.Items.Clear();

                bool hasSelected = clbProvinces.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                clbCities.Enabled = hasSelected;
                clbAffairs.Enabled = hasSelected;
                clbDepartments.Enabled = false;
                clbDistricts.Enabled = false;
                clbPosts.Enabled = false;

                if (!hasSelected)
                    return;

                List<string> selected = clbProvinces.CheckedItems.Cast<string>().ToList();
                if (e.NewValue == CheckState.Checked)
                    selected.Add(Convert.ToString(clbProvinces.Items[e.Index])!);
                else
                    selected.Remove(Convert.ToString(clbProvinces.Items[e.Index])!);

                foreach (string city in DistinctValues("شهر", selected, "استان"))
                    clbCities.Items.Add(city, false);
                foreach (string affair in DistinctValues("امور", selected, "استان"))
                    clbAffairs.Items.Add(affair, false);
            }));
        }

        private void ClbCities_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)(() =>
            {
                if (_loading)
                    return;

                clbDepartments.Items.Clear();
                clbDistricts.Items.Clear();
                clbPosts.Items.Clear();

                bool hasSelected = clbCities.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                bool hasAffairs = clbAffairs.CheckedItems.Count > 0;
                clbDepartments.Enabled = hasSelected || hasAffairs;
                clbDistricts.Enabled = false;
                clbPosts.Enabled = false;

                if (!clbDepartments.Enabled)
                    return;

                foreach (string dept in DistinctDepts())
                    clbDepartments.Items.Add(dept, false);
            }));
        }

        private void ClbAffairs_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)(() =>
            {
                if (_loading)
                    return;

                clbDepartments.Items.Clear();
                clbDistricts.Items.Clear();
                clbPosts.Items.Clear();

                bool hasSelected = clbAffairs.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                bool hasCities = clbCities.CheckedItems.Count > 0;
                clbDepartments.Enabled = hasSelected || hasCities;
                clbDistricts.Enabled = false;
                clbPosts.Enabled = false;

                if (!clbDepartments.Enabled)
                    return;

                foreach (string dept in DistinctDepts())
                    clbDepartments.Items.Add(dept, false);
            }));
        }

        private void ClbDepartments_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)(() =>
            {
                if (_loading)
                    return;

                clbDistricts.Items.Clear();
                clbPosts.Items.Clear();

                bool hasSelected = clbDepartments.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                clbDistricts.Enabled = hasSelected;
                clbPosts.Enabled = false;

                if (!hasSelected)
                    return;

                List<string> selected = clbDepartments.CheckedItems.Cast<string>().ToList();
                if (e.NewValue == CheckState.Checked)
                    selected.Add(Convert.ToString(clbDepartments.Items[e.Index])!);
                else
                    selected.Remove(Convert.ToString(clbDepartments.Items[e.Index])!);

                foreach (string district in DistinctValues("ناحیه", selected, "اداره"))
                    clbDistricts.Items.Add(district, false);
            }));
        }

        private void ClbDistricts_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)(() =>
            {
                if (_loading)
                    return;

                clbPosts.Items.Clear();

                bool hasSelected = clbDistricts.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                clbPosts.Enabled = hasSelected;

                if (!hasSelected)
                    return;

                List<string> selected = clbDistricts.CheckedItems.Cast<string>().ToList();
                if (e.NewValue == CheckState.Checked)
                    selected.Add(Convert.ToString(clbDistricts.Items[e.Index])!);
                else
                    selected.Remove(Convert.ToString(clbDistricts.Items[e.Index])!);

                foreach (string post in DistinctValues("نام پست", selected, "ناحیه"))
                    clbPosts.Items.Add(post, false);
            }));
        }

        private IEnumerable<string> DistinctValues(string column, IEnumerable<string> parents, string parentColumn)
        {
            if (searchTablePersonnel == null)
                return Array.Empty<string>();

            return searchTablePersonnel.AsEnumerable()
                .Where(r => parents.Contains(Convert.ToString(r[parentColumn])))
                .Select(r => Convert.ToString(r[column]) ?? "")
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(v => v, StringComparer.CurrentCulture);
        }

        private IEnumerable<string> DistinctDepts()
        {
            if (searchTablePersonnel == null)
                return Array.Empty<string>();

            List<string> provinces = clbProvinces.CheckedItems.Cast<string>().ToList();
            List<string> cities = clbCities.CheckedItems.Cast<string>().ToList();
            List<string> affairs = clbAffairs.CheckedItems.Cast<string>().ToList();

            return searchTablePersonnel.AsEnumerable()
                .Where(r => (provinces.Count == 0 || provinces.Contains(Convert.ToString(r["استان"]))) &&
                            (cities.Count == 0 || cities.Contains(Convert.ToString(r["شهر"]))) &&
                            (affairs.Count == 0 || affairs.Contains(Convert.ToString(r["امور"]))))
                .Select(r => Convert.ToString(r["اداره"]) ?? "")
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(v => v, StringComparer.CurrentCulture);
        }

        // ─────────────────────────────────────────────
        // نتایج
        // ─────────────────────────────────────────────
        private void ApplyResults()
        {
            if (searchTablePersonnel == null)
                return;

            string search = txtFreeSearch.Text.Trim();
            List<string> provinces = clbProvinces.CheckedItems.Cast<string>().ToList();
            List<string> cities = clbCities.CheckedItems.Cast<string>().ToList();
            List<string> affairs = clbAffairs.CheckedItems.Cast<string>().ToList();
            List<string> depts = clbDepartments.CheckedItems.Cast<string>().ToList();
            List<string> districts = clbDistricts.CheckedItems.Cast<string>().ToList();
            List<string> posts = clbPosts.CheckedItems.Cast<string>().ToList();

            IEnumerable<DataRow> rows = searchTablePersonnel.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(search))
            {
                rows = rows.Where(r =>
                    (Convert.ToString(r["نام"]) ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase) ||
                    (Convert.ToString(r["نام خانوادگی"]) ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase) ||
                    (Convert.ToString(r["نام پدر"]) ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase) ||
                    (Convert.ToString(r["شماره پرسنلی"]) ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase) ||
                    (Convert.ToString(r["کد ملی"]) ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase) ||
                    (Convert.ToString(r["شماره موبایل"]) ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase));
            }

            if (provinces.Count > 0)
                rows = rows.Where(r => provinces.Contains(Convert.ToString(r["استان"])));
            if (cities.Count > 0)
                rows = rows.Where(r => cities.Contains(Convert.ToString(r["شهر"])));
            if (affairs.Count > 0)
                rows = rows.Where(r => affairs.Contains(Convert.ToString(r["امور"])));
            if (depts.Count > 0)
                rows = rows.Where(r => depts.Contains(Convert.ToString(r["اداره"])));
            if (districts.Count > 0)
                rows = rows.Where(r => districts.Contains(Convert.ToString(r["ناحیه"])));
            if (posts.Count > 0)
                rows = rows.Where(r => posts.Contains(Convert.ToString(r["نام پست"])));

            DataTable filtered = searchTablePersonnel.Clone();
            foreach (DataRow row in rows)
                filtered.ImportRow(row);

            dgvResults.DataSource = filtered;
            foreach (DataGridViewColumn column in dgvResults.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.Automatic;
                column.MinimumWidth = 70;
            }

            lblCount.Text = $"{filtered.Rows.Count:N0} پرسنل";
            lblStatus.Text = $"نمایش {filtered.Rows.Count} از {searchTablePersonnel.Rows.Count} پرسنل  |  " +
                             $"آخرین بروزرسانی: {DateTime.Now:HH:mm:ss}";
        }

        private void ClearFilters()
        {
            _loading = true;
            try
            {
                foreach (var list in new[] { clbProvinces, clbCities, clbAffairs, clbDepartments, clbDistricts, clbPosts })
                {
                    for (int i = 0; i < list.Items.Count; i++)
                        list.SetItemChecked(i, false);
                    list.Items.Clear();
                }

                clbCities.Enabled = false;
                clbAffairs.Enabled = false;
                clbDepartments.Enabled = false;
                clbDistricts.Enabled = false;
                clbPosts.Enabled = false;
            }
            finally
            {
                _loading = false;
            }

            LoadProvinces();
            ApplyResults();
        }

        private void ExportCsv()
        {
            if (dgvResults.DataSource is not DataTable dt || dt.Rows.Count == 0)
            {
                MessageBox.Show(this, "هیچ داده‌ای برای خروجی وجود ندارد.", "هشدار",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var dialog = new SaveFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv",
                Title = "ذخیره نتایج جستجو",
                FileName = $"پرسنل_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
                AddExtension = true
            };
            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            try
            {
                db.ExportToCsv(dt, dialog.FileName);
                MessageBox.Show(this, "✅ فایل CSV با موفقیت ذخیره شد.", "خروجی",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPersonnelSearch.ExportCsv", ex);
                MessageBox.Show(this, $"خطا در ذخیره CSV:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteSelected()
        {
            if (dgvResults.CurrentRow == null || dgvResults.CurrentRow.Cells["شناسه پرسنل"].Value == null)
            {
                MessageBox.Show(this, "ابتدا یک پرسنل را از جدول نتایج انتخاب کنید.", "هشدار",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int personnelId = Convert.ToInt32(dgvResults.CurrentRow.Cells["شناسه پرسنل"].Value);
            string fullName = $"{dgvResults.CurrentRow.Cells["نام"].Value} {dgvResults.CurrentRow.Cells["نام خانوادگی"].Value}".Trim();

            if (MessageBox.Show(this,
                    $"آیا مطمئن هستید که می‌خواهید پرسنل «{fullName}» را حذف کنید؟\n\nاین عملیات قابل بازگشت نیست.",
                    "تأیید حذف", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            try
            {
                int affected = db.ExecuteNonQuery("DELETE FROM Personnel WHERE PersonnelID = ?",
                    new[] { new System.Data.OleDb.OleDbParameter("?", personnelId) });

                if (affected > 0)
                {
                    MessageBox.Show(this, "✅ پرسنل با موفقیت حذف شد.", "موفقیت",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                    CreateSearchTable();
                    LoadProvinces();
                    ApplyResults();
                }
                else
                {
                    MessageBox.Show(this, "هیچ پرسنلی حذف نشد!", "هشدار",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPersonnelSearch.DeleteSelected", ex);
                MessageBox.Show(this, $"خطا در حذف پرسنل:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
