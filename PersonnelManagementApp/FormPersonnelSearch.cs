using System;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Linq;

namespace PersonnelManagementApp
{
    public partial class FormPersonnelSearch : Form
    {
        private DbHelper db = new DbHelper();
        private DataTable personnelTable;
        private DataTable provincesTable;
        private DataTable citiesTable;
        private DataTable transferAffairsTable;
        private DataTable operationDepartmentsTable;
        private DataTable districtsTable;
        private DataTable postsNamesTable;
        private DataTable searchTablePersonnel;
        private TextBox txtFreeSearch;
        private CheckedListBox clbProvinces;
        private CheckedListBox clbCities;
        private CheckedListBox clbAffairs;
        private CheckedListBox clbDepartments;
        private CheckedListBox clbDistricts;
        private CheckedListBox clbPosts;
        private DataGridView dgvResults;
        private Button btnFreeSearch, btnShow, btnExport, btnBack;

        public FormPersonnelSearch()
        {
            InitializeComponent();

            // اعمال فونت‌های تنظیم‌شده
            FontSettings.ApplyFontToForm(this);

            LoadData();
            CreateSearchTable();
            LoadProvinces();
        }

        private void InitializeComponent()
        {
            this.Text = "جستجوی پرسنل";
            this.WindowState = FormWindowState.Maximized;
            this.RightToLeft = RightToLeft.Yes;
            this.BackColor = Color.FromArgb(240, 248, 255);

            // پس‌زمینه گرادیانت
            using (LinearGradientBrush brush = new LinearGradientBrush(this.ClientRectangle, Color.LightBlue, Color.White, LinearGradientMode.Vertical))
            {
                this.BackgroundImage = new Bitmap(this.Width, this.Height);
                using (Graphics g = Graphics.FromImage(this.BackgroundImage))
                {
                    g.FillRectangle(brush, this.ClientRectangle);
                }
            }

            // گرید نمایش
            dgvResults = new DataGridView
            {
                Name = "dgvResults",
                Location = new Point(50, 50),
                Size = new Size(this.ClientSize.Width - 100, 300),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None, // غیرفعال کردن AutoSize برای اسکرول افقی
                ScrollBars = ScrollBars.Both, // فعال کردن اسکرول افقی و عمودی
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                ColumnHeadersHeight = 40,
                RowTemplate = { Height = 30 },
                AllowUserToResizeColumns = true,
                AllowUserToResizeRows = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            ApplyRoundedCorners(dgvResults, 20);
            dgvResults.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) dgvResults.BackColor = Color.LightYellow; else dgvResults.BackColor = Color.White; };

            // جستجوی آزاد
            txtFreeSearch = new TextBox
            {
                Location = new Point(50, 360),
                Size = new Size(300, 40),
                Font = FontSettings.TextBoxFont,
                PlaceholderText = "جستجوی آزاد (نام، نام خانوادگی، شماره پرسنلی، کد ملی)...",
                BorderStyle = BorderStyle.None
            };
            ApplyRoundedCorners(txtFreeSearch, 15);
            txtFreeSearch.BackColor = Color.WhiteSmoke;
            btnFreeSearch = new Button
            {
                Text = "جستجو",
                Location = new Point(360, 360),
                Size = new Size(150, 40),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightBlue,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnFreeSearch, 15);
            btnFreeSearch.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) btnFreeSearch.BackColor = Color.RoyalBlue; else btnFreeSearch.BackColor = Color.LightBlue; };

            // جستجوی انتخابی با لیبل‌ها
            Label lblProvinces = new Label
            {
                Text = "استان‌ها:",
                Location = new Point(50, 410),
                Size = new Size(200, 30),
                Font = FontSettings.SubtitleFont,
                ForeColor = Color.DarkBlue
            };
            clbProvinces = new CheckedListBox { Location = new Point(50, 440), Size = new Size(200, 150), BorderStyle = BorderStyle.None };
            Label lblCities = new Label
            {
                Text = "شهرها:",
                Location = new Point(260, 410),
                Size = new Size(200, 30),
                Font = FontSettings.SubtitleFont,
                ForeColor = Color.DarkBlue
            };
            clbCities = new CheckedListBox { Location = new Point(260, 440), Size = new Size(200, 150), BorderStyle = BorderStyle.None, Enabled = false };
            Label lblAffairs = new Label
            {
                Text = "امور:",
                Location = new Point(470, 410),
                Size = new Size(200, 30),
                Font = FontSettings.SubtitleFont,
                ForeColor = Color.DarkBlue
            };
            clbAffairs = new CheckedListBox { Location = new Point(470, 440), Size = new Size(200, 150), BorderStyle = BorderStyle.None, Enabled = false };
            Label lblDepartments = new Label
            {
                Text = "ادارات:",
                Location = new Point(680, 410),
                Size = new Size(200, 30),
                Font = FontSettings.SubtitleFont,
                ForeColor = Color.DarkBlue
            };
            clbDepartments = new CheckedListBox { Location = new Point(680, 440), Size = new Size(200, 150), BorderStyle = BorderStyle.None, Enabled = false };
            Label lblDistricts = new Label
            {
                Text = "نواحی:",
                Location = new Point(890, 410),
                Size = new Size(200, 30),
                Font = FontSettings.SubtitleFont,
                ForeColor = Color.DarkBlue
            };
            clbDistricts = new CheckedListBox { Location = new Point(890, 440), Size = new Size(200, 150), BorderStyle = BorderStyle.None, Enabled = false };
            Label lblPosts = new Label
            {
                Text = "پست‌ها:",
                Location = new Point(1100, 410),
                Size = new Size(200, 30),
                Font = FontSettings.SubtitleFont,
                ForeColor = Color.DarkBlue
            };
            clbPosts = new CheckedListBox { Location = new Point(1100, 440), Size = new Size(200, 150), BorderStyle = BorderStyle.None, Enabled = false };

            ApplyRoundedCorners(clbProvinces, 15);
            ApplyRoundedCorners(clbCities, 15);
            ApplyRoundedCorners(clbAffairs, 15);
            ApplyRoundedCorners(clbDepartments, 15);
            ApplyRoundedCorners(clbDistricts, 15);
            ApplyRoundedCorners(clbPosts, 15);
            clbProvinces.BackColor = Color.WhiteSmoke;
            clbCities.BackColor = Color.WhiteSmoke;
            clbAffairs.BackColor = Color.WhiteSmoke;
            clbDepartments.BackColor = Color.WhiteSmoke;
            clbDistricts.BackColor = Color.WhiteSmoke;
            clbPosts.BackColor = Color.WhiteSmoke;

            // دکمه‌ها
            btnShow = new Button
            {
                Text = "نمایش نتایج",
                Location = new Point(this.ClientSize.Width - 250, 600),
                Size = new Size(200, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightGreen,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnShow, 20);
            btnShow.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) btnShow.BackColor = Color.ForestGreen; else btnShow.BackColor = Color.LightGreen; };

            btnExport = new Button
            {
                Text = "اکسپورت به CSV",
                Location = new Point(this.ClientSize.Width - 460, 600),
                Size = new Size(200, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightBlue,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnExport, 20);
            btnExport.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) btnExport.BackColor = Color.RoyalBlue; else btnExport.BackColor = Color.LightBlue; };

            btnBack = new Button
            {
                Text = "برگشت به صفحه اصلی",
                Location = new Point(50, 600),
                Size = new Size(200, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightCoral,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnBack, 20);
            btnBack.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) btnBack.BackColor = Color.Red; else btnBack.BackColor = Color.LightCoral; };
            btnBack.Click += (s, e) => { this.Close(); };

            this.Controls.Add(dgvResults);
            this.Controls.Add(txtFreeSearch);
            this.Controls.Add(btnFreeSearch);
            this.Controls.Add(lblProvinces);
            this.Controls.Add(clbProvinces);
            this.Controls.Add(lblCities);
            this.Controls.Add(clbCities);
            this.Controls.Add(lblAffairs);
            this.Controls.Add(clbAffairs);
            this.Controls.Add(lblDepartments);
            this.Controls.Add(clbDepartments);
            this.Controls.Add(lblDistricts);
            this.Controls.Add(clbDistricts);
            this.Controls.Add(lblPosts);
            this.Controls.Add(clbPosts);
            this.Controls.Add(btnShow);
            this.Controls.Add(btnExport);
            this.Controls.Add(btnBack);

            btnFreeSearch.Click += BtnFreeSearch_Click;
            btnShow.Click += BtnShow_Click;
            btnExport.Click += BtnExport_Click;
            clbProvinces.ItemCheck += ClbProvinces_ItemCheck;
            clbCities.ItemCheck += ClbCities_ItemCheck;
            clbAffairs.ItemCheck += ClbAffairs_ItemCheck;
            clbDepartments.ItemCheck += ClbDepartments_ItemCheck;
            clbDistricts.ItemCheck += ClbDistricts_ItemCheck;
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

        private void LoadData()
        {
            try
            {
                provincesTable = db.GetProvinces();
                citiesTable = db.GetCitiesByProvince(0); // دریافت همه شهرها
                transferAffairsTable = db.GetAffairsByProvince(0); // دریافت همه امور
                operationDepartmentsTable = db.GetDeptsByAffair(0); // دریافت همه ادارات
                districtsTable = db.GetDistrictsByDept(0); // دریافت همه نواحی
                postsNamesTable = db.GetPostNamesByDistrict(0); // دریافت همه نام‌های پست
                personnelTable = db.ExecuteQuery("SELECT PersonnelID, ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, VoltageID, " +
                                                "WorkShiftID, GenderID, FirstName, LastName, FatherName, PersonnelNumber, NationalID, " +
                                                "MobileNumber, BirthDate, HireDate, StartDateOperation, ContractTypeID, JobLevelID, " +
                                                "CompanyID, DegreeID, DegreeFieldID, MainJobTitle, CurrentActivity, Inconsistency, " +
                                                "Description, StatusID FROM Personnel");

                MessageBox.Show($"Personnel: {personnelTable.Rows.Count} ردیف\nProvinces: {provincesTable.Rows.Count} ردیف");
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در بارگذاری داده‌ها: " + ex.Message);
            }
        }

        private void CreateSearchTable()
        {
            searchTablePersonnel = new DataTable("SearchTablePersonnel");
            searchTablePersonnel.Columns.Add("PersonnelID", typeof(int)).ColumnName = "شناسه پرسنل";
            searchTablePersonnel.Columns.Add("FirstName", typeof(string)).ColumnName = "نام";
            searchTablePersonnel.Columns.Add("LastName", typeof(string)).ColumnName = "نام خانوادگی";
            searchTablePersonnel.Columns.Add("FatherName", typeof(string)).ColumnName = "نام پدر";
            searchTablePersonnel.Columns.Add("PersonnelNumber", typeof(string)).ColumnName = "شماره پرسنلی";
            searchTablePersonnel.Columns.Add("NationalID", typeof(string)).ColumnName = "کد ملی";
            searchTablePersonnel.Columns.Add("MobileNumber", typeof(string)).ColumnName = "شماره موبایل";
            searchTablePersonnel.Columns.Add("ProvinceName", typeof(string)).ColumnName = "استان";
            searchTablePersonnel.Columns.Add("CityName", typeof(string)).ColumnName = "شهر";
            searchTablePersonnel.Columns.Add("AffairName", typeof(string)).ColumnName = "امور";
            searchTablePersonnel.Columns.Add("DeptName", typeof(string)).ColumnName = "اداره";
            searchTablePersonnel.Columns.Add("DistrictName", typeof(string)).ColumnName = "ناحیه";
            searchTablePersonnel.Columns.Add("PostName", typeof(string)).ColumnName = "نام پست";
            searchTablePersonnel.Columns.Add("VoltageName", typeof(string)).ColumnName = "سطح ولتاژ";
            searchTablePersonnel.Columns.Add("WorkShiftName", typeof(string)).ColumnName = "شیفت کاری";
            searchTablePersonnel.Columns.Add("GenderName", typeof(string)).ColumnName = "جنسیت";
            searchTablePersonnel.Columns.Add("ContractTypeName", typeof(string)).ColumnName = "نوع قرارداد";
            searchTablePersonnel.Columns.Add("JobLevelName", typeof(string)).ColumnName = "سطح شغل";
            searchTablePersonnel.Columns.Add("CompanyName", typeof(string)).ColumnName = "شرکت";
            searchTablePersonnel.Columns.Add("DegreeName", typeof(string)).ColumnName = "مدرک تحصیلی";
            searchTablePersonnel.Columns.Add("DegreeFieldName", typeof(string)).ColumnName = "رشته تحصیلی";
            searchTablePersonnel.Columns.Add("MainJobTitle", typeof(string)).ColumnName = "عنوان شغلی اصلی";
            searchTablePersonnel.Columns.Add("CurrentActivity", typeof(string)).ColumnName = "فعالیت فعلی";
            searchTablePersonnel.Columns.Add("Inconsistency", typeof(bool)).ColumnName = "مغایرت";
            searchTablePersonnel.Columns.Add("Description", typeof(string)).ColumnName = "توضیحات";
            searchTablePersonnel.Columns.Add("StatusName", typeof(string)).ColumnName = "وضعیت حضور";
            searchTablePersonnel.Columns.Add("BirthDate", typeof(string)).ColumnName = "تاریخ تولد";
            searchTablePersonnel.Columns.Add("HireDate", typeof(string)).ColumnName = "تاریخ استخدام";
            searchTablePersonnel.Columns.Add("StartDateOperation", typeof(string)).ColumnName = "تاریخ شروع بکار";

            var voltageLevelsTable = db.GetVoltageLevels();
            var workShiftsTable = db.GetWorkShifts();
            var gendersTable = db.GetGenders();
            var contractTypesTable = db.GetContractTypes();
            var jobLevelsTable = db.GetJobLevels();
            var companiesTable = db.GetCompanies();
            var degreesTable = db.GetDegrees();
            var degreeFieldsTable = db.GetDegreeFields();
            var chartAffairsTable = db.GetChartAffairs();
            var statusPresenceTable = db.GetStatusPresence();

            foreach (DataRow personnelRow in personnelTable.Rows)
            {
                DataRow newRow = searchTablePersonnel.NewRow();
                newRow["شناسه پرسنل"] = personnelRow["PersonnelID"];
                newRow["نام"] = personnelRow["FirstName"];
                newRow["نام خانوادگی"] = personnelRow["LastName"];
                newRow["نام پدر"] = personnelRow["FatherName"];
                newRow["شماره پرسنلی"] = personnelRow["PersonnelNumber"];
                newRow["کد ملی"] = personnelRow["NationalID"];
                newRow["شماره موبایل"] = personnelRow["MobileNumber"];
                newRow["تاریخ تولد"] = personnelRow["BirthDate"];
                newRow["تاریخ استخدام"] = personnelRow["HireDate"];
                newRow["تاریخ شروع بکار"] = personnelRow["StartDateOperation"];

                int? provinceId = personnelRow["ProvinceID"] as int?;
                if (provinceId.HasValue)
                {
                    DataRow[] provinceRows = provincesTable.Select($"ProvinceID = {provinceId.Value}");
                    newRow["استان"] = provinceRows.Length > 0 ? provinceRows[0]["ProvinceName"] : "نامشخص";
                }

                int? cityId = personnelRow["CityID"] as int?;
                if (cityId.HasValue)
                {
                    DataRow[] cityRows = citiesTable.Select($"CityID = {cityId.Value}");
                    newRow["شهر"] = cityRows.Length > 0 ? cityRows[0]["CityName"] : "نامشخص";
                }

                int? affairId = personnelRow["AffairID"] as int?;
                if (affairId.HasValue)
                {
                    DataRow[] affairRows = transferAffairsTable.Select($"AffairID = {affairId.Value}");
                    newRow["امور"] = affairRows.Length > 0 ? affairRows[0]["AffairName"] : "نامشخص";
                }

                int? deptId = personnelRow["DeptID"] as int?;
                if (deptId.HasValue)
                {
                    DataRow[] deptRows = operationDepartmentsTable.Select($"DeptID = {deptId.Value}");
                    newRow["اداره"] = deptRows.Length > 0 ? deptRows[0]["DeptName"] : "نامشخص";
                }

                int? districtId = personnelRow["DistrictID"] as int?;
                if (districtId.HasValue)
                {
                    DataRow[] districtRows = districtsTable.Select($"DistrictID = {districtId.Value}");
                    newRow["ناحیه"] = districtRows.Length > 0 ? districtRows[0]["DistrictName"] : "نامشخص";
                }

                int? postNameId = personnelRow["PostNameID"] as int?;
                if (postNameId.HasValue)
                {
                    DataRow[] postNameRows = postsNamesTable.Select($"PostNameID = {postNameId.Value}");
                    newRow["نام پست"] = postNameRows.Length > 0 ? postNameRows[0]["PostName"] : "نامشخص";
                }

                int? voltageId = personnelRow["VoltageID"] as int?;
                if (voltageId.HasValue)
                {
                    DataRow[] voltageRows = voltageLevelsTable.Select($"VoltageID = {voltageId.Value}");
                    newRow["سطح ولتاژ"] = voltageRows.Length > 0 ? voltageRows[0]["VoltageName"] : "نامشخص";
                }

                int? workShiftId = personnelRow["WorkShiftID"] as int?;
                if (workShiftId.HasValue)
                {
                    DataRow[] workShiftRows = workShiftsTable.Select($"WorkShiftID = {workShiftId.Value}");
                    newRow["شیفت کاری"] = workShiftRows.Length > 0 ? workShiftRows[0]["WorkShiftName"] : "نامشخص";
                }

                int? genderId = personnelRow["GenderID"] as int?;
                if (genderId.HasValue)
                {
                    DataRow[] genderRows = gendersTable.Select($"GenderID = {genderId.Value}");
                    newRow["جنسیت"] = genderRows.Length > 0 ? genderRows[0]["GenderName"] : "نامشخص";
                }

                int? contractTypeId = personnelRow["ContractTypeID"] as int?;
                if (contractTypeId.HasValue)
                {
                    DataRow[] contractTypeRows = contractTypesTable.Select($"ContractTypeID = {contractTypeId.Value}");
                    newRow["نوع قرارداد"] = contractTypeRows.Length > 0 ? contractTypeRows[0]["ContractTypeName"] : "نامشخص";
                }

                int? jobLevelId = personnelRow["JobLevelID"] as int?;
                if (jobLevelId.HasValue)
                {
                    DataRow[] jobLevelRows = jobLevelsTable.Select($"JobLevelID = {jobLevelId.Value}");
                    newRow["سطح شغل"] = jobLevelRows.Length > 0 ? jobLevelRows[0]["JobLevelName"] : "نامشخص";
                }

                int? companyId = personnelRow["CompanyID"] as int?;
                if (companyId.HasValue)
                {
                    DataRow[] companyRows = companiesTable.Select($"CompanyID = {companyId.Value}");
                    newRow["شرکت"] = companyRows.Length > 0 ? companyRows[0]["CompanyName"] : "نامشخص";
                }

                int? degreeId = personnelRow["DegreeID"] as int?;
                if (degreeId.HasValue)
                {
                    DataRow[] degreeRows = degreesTable.Select($"DegreeID = {degreeId.Value}");
                    newRow["مدرک تحصیلی"] = degreeRows.Length > 0 ? degreeRows[0]["DegreeName"] : "نامشخص";
                }

                int? degreeFieldId = personnelRow["DegreeFieldID"] as int?;
                if (degreeFieldId.HasValue)
                {
                    DataRow[] degreeFieldRows = degreeFieldsTable.Select($"DegreeFieldID = {degreeFieldId.Value}");
                    newRow["رشته تحصیلی"] = degreeFieldRows.Length > 0 ? degreeFieldRows[0]["DegreeFieldName"] : "نامشخص";
                }

                int? mainJobTitleId = personnelRow["MainJobTitle"] as int?;
                if (mainJobTitleId.HasValue)
                {
                    DataRow[] mainJobTitleRows = chartAffairsTable.Select($"ChartID = {mainJobTitleId.Value}");
                    newRow["عنوان شغلی اصلی"] = mainJobTitleRows.Length > 0 ? mainJobTitleRows[0]["ChartName"] : "نامشخص";
                }

                int? currentActivityId = personnelRow["CurrentActivity"] as int?;
                if (currentActivityId.HasValue)
                {
                    DataRow[] currentActivityRows = chartAffairsTable.Select($"ChartID = {currentActivityId.Value}");
                    newRow["فعالیت فعلی"] = currentActivityRows.Length > 0 ? currentActivityRows[0]["ChartName"] : "نامشخص";
                }

                newRow["مغایرت"] = personnelRow["Inconsistency"];
                newRow["توضیحات"] = personnelRow["Description"];
                int? statusId = personnelRow["StatusID"] as int?;
                if (statusId.HasValue)
                {
                    DataRow[] statusRows = statusPresenceTable.Select($"StatusID = {statusId.Value}");
                    newRow["وضعیت حضور"] = statusRows.Length > 0 ? statusRows[0]["StatusName"] : "نامشخص";
                }

                searchTablePersonnel.Rows.Add(newRow);
            }

            MessageBox.Show($"SearchTablePersonnel ساخته شد با {searchTablePersonnel.Rows.Count} ردیف.");
        }

        private void LoadProvinces()
        {
            clbProvinces.Items.Clear();
            foreach (DataRow row in provincesTable.Rows)
                clbProvinces.Items.Add(row["ProvinceName"].ToString(), false);
        }

        private void ClbProvinces_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                // پاک کردن آیتم‌های وابسته
                clbCities.Items.Clear();
                clbAffairs.Items.Clear();
                clbDepartments.Items.Clear();
                clbDistricts.Items.Clear();
                clbPosts.Items.Clear();

                // فعال کردن شهرها و امور
                bool hasSelectedProvinces = clbProvinces.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                clbCities.Enabled = hasSelectedProvinces;
                clbAffairs.Enabled = hasSelectedProvinces;
                clbDepartments.Enabled = false;
                clbDistricts.Enabled = false;
                clbPosts.Enabled = false;

                if (hasSelectedProvinces)
                {
                    var selectedProvinces = clbProvinces.CheckedItems.Cast<string>().ToList();
                    if (e.NewValue == CheckState.Checked)
                        selectedProvinces.Add(clbProvinces.Items[e.Index].ToString());
                    else
                        selectedProvinces.Remove(clbProvinces.Items[e.Index].ToString());

                    // بارگذاری شهرها
                    var cities = searchTablePersonnel.AsEnumerable()
                        .Where(r => selectedProvinces.Contains(r["استان"].ToString()))
                        .Select(r => r["شهر"].ToString())
                        .Distinct()
                        .OrderBy(c => c);
                    foreach (var city in cities)
                        clbCities.Items.Add(city, false);

                    // بارگذاری امور
                    var affairs = searchTablePersonnel.AsEnumerable()
                        .Where(r => selectedProvinces.Contains(r["استان"].ToString()))
                        .Select(r => r["امور"].ToString())
                        .Distinct()
                        .OrderBy(a => a);
                    foreach (var affair in affairs)
                        clbAffairs.Items.Add(affair, false);
                }
            });
        }

        private void ClbCities_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                // پاک کردن آیتم‌های وابسته
                clbDepartments.Items.Clear();
                clbDistricts.Items.Clear();
                clbPosts.Items.Clear();

                // فعال کردن ادارات فقط اگر شهر یا امور انتخاب شده باشد
                bool hasSelectedCities = clbCities.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                bool hasSelectedAffairs = clbAffairs.CheckedItems.Count > 0;
                clbDepartments.Enabled = hasSelectedCities || hasSelectedAffairs;
                clbDistricts.Enabled = false;
                clbPosts.Enabled = false;

                if (hasSelectedCities)
                {
                    var selectedProvinces = clbProvinces.CheckedItems.Cast<string>().ToList();
                    var selectedCities = clbCities.CheckedItems.Cast<string>().ToList();
                    var selectedAffairs = clbAffairs.CheckedItems.Cast<string>().ToList();
                    if (e.NewValue == CheckState.Checked)
                        selectedCities.Add(clbCities.Items[e.Index].ToString());
                    else
                        selectedCities.Remove(clbCities.Items[e.Index].ToString());

                    // بارگذاری ادارات با در نظر گرفتن شهرها و امور انتخاب‌شده
                    var depts = searchTablePersonnel.AsEnumerable()
                        .Where(r => selectedProvinces.Contains(r["استان"].ToString()) &&
                                    (selectedCities.Count == 0 || selectedCities.Contains(r["شهر"].ToString())) &&
                                    (selectedAffairs.Count == 0 || selectedAffairs.Contains(r["امور"].ToString())))
                        .Select(r => r["اداره"].ToString())
                        .Distinct()
                        .OrderBy(d => d);
                    foreach (var dept in depts)
                        clbDepartments.Items.Add(dept, false);
                }
            });
        }

        private void ClbAffairs_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                // پاک کردن آیتم‌های وابسته
                clbDepartments.Items.Clear();
                clbDistricts.Items.Clear();
                clbPosts.Items.Clear();

                // فعال کردن ادارات فقط اگر شهر یا امور انتخاب شده باشد
                bool hasSelectedCities = clbCities.CheckedItems.Count > 0;
                bool hasSelectedAffairs = clbAffairs.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                clbDepartments.Enabled = hasSelectedCities || hasSelectedAffairs;
                clbDistricts.Enabled = false;
                clbPosts.Enabled = false;

                if (hasSelectedAffairs)
                {
                    var selectedProvinces = clbProvinces.CheckedItems.Cast<string>().ToList();
                    var selectedCities = clbCities.CheckedItems.Cast<string>().ToList();
                    var selectedAffairs = clbAffairs.CheckedItems.Cast<string>().ToList();
                    if (e.NewValue == CheckState.Checked)
                        selectedAffairs.Add(clbAffairs.Items[e.Index].ToString());
                    else
                        selectedAffairs.Remove(clbAffairs.Items[e.Index].ToString());

                    // بارگذاری ادارات با در نظر گرفتن شهرها و امور انتخاب‌شده
                    var depts = searchTablePersonnel.AsEnumerable()
                        .Where(r => selectedProvinces.Contains(r["استان"].ToString()) &&
                                    (selectedCities.Count == 0 || selectedCities.Contains(r["شهر"].ToString())) &&
                                    (selectedAffairs.Count == 0 || selectedAffairs.Contains(r["امور"].ToString())))
                        .Select(r => r["اداره"].ToString())
                        .Distinct()
                        .OrderBy(d => d);
                    foreach (var dept in depts)
                        clbDepartments.Items.Add(dept, false);
                }
            });
        }

        private void ClbDepartments_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                // پاک کردن آیتم‌های وابسته
                clbDistricts.Items.Clear();
                clbPosts.Items.Clear();

                // فعال کردن نواحی
                bool hasSelectedDepartments = clbDepartments.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                clbDistricts.Enabled = hasSelectedDepartments;
                clbPosts.Enabled = false;

                if (hasSelectedDepartments)
                {
                    var selectedProvinces = clbProvinces.CheckedItems.Cast<string>().ToList();
                    var selectedCities = clbCities.CheckedItems.Cast<string>().ToList();
                    var selectedAffairs = clbAffairs.CheckedItems.Cast<string>().ToList();
                    var selectedDepts = clbDepartments.CheckedItems.Cast<string>().ToList();
                    if (e.NewValue == CheckState.Checked)
                        selectedDepts.Add(clbDepartments.Items[e.Index].ToString());
                    else
                        selectedDepts.Remove(clbDepartments.Items[e.Index].ToString());

                    // بارگذاری نواحی
                    var districts = searchTablePersonnel.AsEnumerable()
                        .Where(r => selectedProvinces.Contains(r["استان"].ToString()) &&
                                    (selectedCities.Count == 0 || selectedCities.Contains(r["شهر"].ToString())) &&
                                    (selectedAffairs.Count == 0 || selectedAffairs.Contains(r["امور"].ToString())) &&
                                    selectedDepts.Contains(r["اداره"].ToString()))
                        .Select(r => r["ناحیه"].ToString())
                        .Distinct()
                        .OrderBy(d => d);
                    foreach (var district in districts)
                        clbDistricts.Items.Add(district, false);
                }
            });
        }

        private void ClbDistricts_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                // پاک کردن آیتم‌های وابسته
                clbPosts.Items.Clear();

                // فعال کردن پست‌ها
                bool hasSelectedDistricts = clbDistricts.CheckedItems.Count > 0 || e.NewValue == CheckState.Checked;
                clbPosts.Enabled = hasSelectedDistricts;

                if (hasSelectedDistricts)
                {
                    var selectedProvinces = clbProvinces.CheckedItems.Cast<string>().ToList();
                    var selectedCities = clbCities.CheckedItems.Cast<string>().ToList();
                    var selectedAffairs = clbAffairs.CheckedItems.Cast<string>().ToList();
                    var selectedDepts = clbDepartments.CheckedItems.Cast<string>().ToList();
                    var selectedDistricts = clbDistricts.CheckedItems.Cast<string>().ToList();
                    if (e.NewValue == CheckState.Checked)
                        selectedDistricts.Add(clbDistricts.Items[e.Index].ToString());
                    else
                        selectedDistricts.Remove(clbDistricts.Items[e.Index].ToString());

                    // بارگذاری پست‌ها
                    var posts = searchTablePersonnel.AsEnumerable()
                        .Where(r => selectedProvinces.Contains(r["استان"].ToString()) &&
                                    (selectedCities.Count == 0 || selectedCities.Contains(r["شهر"].ToString())) &&
                                    (selectedAffairs.Count == 0 || selectedAffairs.Contains(r["امور"].ToString())) &&
                                    selectedDepts.Contains(r["اداره"].ToString()) &&
                                    selectedDistricts.Contains(r["ناحیه"].ToString()))
                        .Select(r => r["نام پست"].ToString())
                        .Distinct()
                        .OrderBy(p => p);
                    foreach (var post in posts)
                        clbPosts.Items.Add(post, false);
                }
            });
        }

        private void BtnFreeSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string searchText = txtFreeSearch.Text.Trim().ToLower();
                var filteredTable = searchTablePersonnel.Clone();
                foreach (DataRow row in searchTablePersonnel.Rows)
                {
                    if (row["نام"].ToString().ToLower().Contains(searchText) ||
                        row["نام خانوادگی"].ToString().ToLower().Contains(searchText) ||
                        row["شماره پرسنلی"].ToString().ToLower().Contains(searchText) ||
                        row["کد ملی"].ToString().ToLower().Contains(searchText))
                    {
                        filteredTable.ImportRow(row);
                    }
                }

                dgvResults.DataSource = filteredTable;
                foreach (DataGridViewColumn column in dgvResults.Columns)
                {
                    column.Width = 150; // عرض ثابت برای ایجاد اسکرول افقی
                    column.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در جستجوی آزاد: " + ex.Message);
            }
        }

        private void BtnShow_Click(object sender, EventArgs e)
        {
            try
            {
                var selectedProvinces = clbProvinces.CheckedItems.Cast<string>().ToList();
                var selectedCities = clbCities.CheckedItems.Cast<string>().ToList();
                var selectedAffairs = clbAffairs.CheckedItems.Cast<string>().ToList();
                var selectedDepts = clbDepartments.CheckedItems.Cast<string>().ToList();
                var selectedDistricts = clbDistricts.CheckedItems.Cast<string>().ToList();
                var selectedPosts = clbPosts.CheckedItems.Cast<string>().ToList();

                var filteredTable = searchTablePersonnel.Clone();
                var filteredRows = searchTablePersonnel.AsEnumerable();

                if (selectedProvinces.Any())
                    filteredRows = filteredRows.Where(r => selectedProvinces.Contains(r["استان"].ToString()));
                if (selectedCities.Any())
                    filteredRows = filteredRows.Where(r => selectedCities.Contains(r["شهر"].ToString()));
                if (selectedAffairs.Any())
                    filteredRows = filteredRows.Where(r => selectedAffairs.Contains(r["امور"].ToString()));
                if (selectedDepts.Any())
                    filteredRows = filteredRows.Where(r => selectedDepts.Contains(r["اداره"].ToString()));
                if (selectedDistricts.Any())
                    filteredRows = filteredRows.Where(r => selectedDistricts.Contains(r["ناحیه"].ToString()));
                if (selectedPosts.Any())
                    filteredRows = filteredRows.Where(r => selectedPosts.Contains(r["نام پست"].ToString()));

                foreach (var row in filteredRows)
                    filteredTable.ImportRow(row);

                dgvResults.DataSource = filteredTable;
                foreach (DataGridViewColumn column in dgvResults.Columns)
                {
                    column.Width = 150; // عرض ثابت برای ایجاد اسکرول افقی
                    column.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در نمایش نتایج: " + ex.Message);
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (dgvResults.DataSource is DataTable dt && dt.Rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "CSV Files (*.csv)|*.csv",
                    FileName = "PersonnelSearchResults.csv"
                };
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    db.ExportToCsv(dt, sfd.FileName);
                }
            }
            else
            {
                MessageBox.Show("هیچ داده‌ای برای اکسپورت وجود ندارد!", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}