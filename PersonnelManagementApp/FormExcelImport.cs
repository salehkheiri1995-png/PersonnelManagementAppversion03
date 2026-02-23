using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    public partial class FormImportExcel : Form
    {
        private DbHelper db = new DbHelper();
        private TextBox txtFilePath;
        private DataGridView dgvPreview;
        private readonly int expectedColumnCount = 26;

        // جداول lookup
        private DataTable provinces;
        private DataTable cities;
        private DataTable affairs;       // امور (TransferAffairs)
        private DataTable depts;         // ناحیه بهره‌برداری (OperationDepartments)
        private DataTable districts;
        private DataTable postNames;
        private DataTable voltages;
        private DataTable workShifts;
        private DataTable genders;
        private DataTable contractTypes;
        private DataTable jobLevels;
        private DataTable companies;
        private DataTable degrees;
        private DataTable degreeFields;
        private DataTable chartAffairs;

        // رویداد برای اطلاع‌رسانی به فرم اصلی
        public event EventHandler ImportCompleted;

        public FormImportExcel()
        {
            LoadTables();
            InitializeComponent();
        }

        /// <summary>
        /// لود تمام جداول lookup از دیتابیس
        /// </summary>
        private void LoadTables()
        {
            try
            {
                provinces = db.GetProvinces();
                cities = db.ExecuteQuery("SELECT CityID, CityName FROM Cities");
                affairs = db.ExecuteQuery("SELECT AffairID, AffairName FROM TransferAffairs");
                depts = db.ExecuteQuery("SELECT DeptID, DeptName FROM OperationDepartments");
                districts = db.ExecuteQuery("SELECT DistrictID, DistrictName FROM Districts");
                postNames = db.ExecuteQuery("SELECT PostNameID, PostName FROM PostsNames");
                voltages = db.GetVoltageLevels();
                workShifts = db.GetWorkShifts();
                genders = db.GetGenders();
                contractTypes = db.GetContractTypes();
                jobLevels = db.GetJobLevels();
                companies = db.GetCompanies();
                degrees = db.GetDegrees();
                degreeFields = db.GetDegreeFields();
                chartAffairs = db.GetChartAffairs();
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در لود جداول پایه: " + ex.Message, "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                File.AppendAllText("import_errors.log",
                    $"[{DateTime.Now}] خطا در LoadTables: {ex.Message}\n");
            }
        }

        private void InitializeComponent()
        {
            this.Text = "وارد کردن اطلاعات از اکسل";
            this.Size = new Size(1200, 900);
            this.WindowState = FormWindowState.Maximized;
            this.AutoScroll = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.RightToLeft = RightToLeft.Yes;
            this.BackColor = Color.FromArgb(240, 248, 255);

            int y = 20;
            int controlWidth = 700;
            int controlHeight = 40;
            int yStep = 60;
            int formWidth = this.ClientSize.Width;
            int labelWidth = 150;
            int totalControlWidth = labelWidth + controlWidth + 10;
            int xControl = (formWidth - totalControlWidth) / 2;
            int xLabel = xControl + controlWidth + 10;

            Label lblFilePath = new Label
            {
                Text = "مسیر فایل اکسل:",
                Location = new Point(xLabel, y),
                Size = new Size(labelWidth, 30),
                Font = new Font("Tahoma", 12)
            };
            txtFilePath = new TextBox
            {
                Location = new Point(xControl, y),
                Size = new Size(controlWidth - 100, controlHeight),
                Font = new Font("Tahoma", 12),
                ReadOnly = true
            };
            this.Controls.Add(lblFilePath);
            this.Controls.Add(txtFilePath);

            Button btnBrowse = new Button
            {
                Text = "انتخاب فایل",
                Location = new Point(xControl + controlWidth - 90, y),
                Size = new Size(90, controlHeight),
                Font = new Font("Tahoma", 12),
                BackColor = Color.SteelBlue,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(btnBrowse);
            y += yStep;

            Label lblPreview = new Label
            {
                Text = "پیش‌نمایش داده‌ها:",
                Location = new Point(xLabel, y),
                Size = new Size(labelWidth, 30),
                Font = new Font("Tahoma", 12)
            };
            this.Controls.Add(lblPreview);
            y += 40;

            dgvPreview = new DataGridView
            {
                Location = new Point(50, y),
                Size = new Size(formWidth - 100, this.ClientSize.Height - y - 120),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                ScrollBars = ScrollBars.Both,
                AllowUserToAddRows = false,
                ReadOnly = true,
                Font = new Font("Tahoma", 10),
                ColumnHeadersHeight = 40,
                RowTemplate = { Height = 30 },
                AllowUserToResizeColumns = true,
                RightToLeft = RightToLeft.Yes,
                Anchor = AnchorStyles.Top | AnchorStyles.Left |
                         AnchorStyles.Right | AnchorStyles.Bottom
            };
            this.Controls.Add(dgvPreview);
            y += this.ClientSize.Height - y - 100;

            int buttonWidth = 180;
            int buttonSpace = 10;
            int totalButtonWidth = (buttonWidth * 2) + buttonSpace;
            int xButtonStart = (formWidth - totalButtonWidth) / 2;

            Button btnImport = new Button
            {
                Text = "وارد کردن به دیتابیس",
                Location = new Point(xButtonStart, y),
                Size = new Size(buttonWidth, 50),
                Font = new Font("Tahoma", 12),
                BackColor = Color.SeaGreen,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnImport.Click += BtnImport_Click;
            this.Controls.Add(btnImport);

            Button btnCancel = new Button
            {
                Text = "لغو",
                Location = new Point(xButtonStart + buttonWidth + buttonSpace, y),
                Size = new Size(buttonWidth, 50),
                Font = new Font("Tahoma", 12),
                BackColor = Color.IndianRed,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                Title = "انتخاب فایل اکسل"
            };
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = ofd.FileName;
                LoadExcelPreview(ofd.FileName);
            }
        }

        private void LoadExcelPreview(string filePath)
        {
            try
            {
                string cs = BuildOleDbConnectionString(filePath);
                using (OleDbConnection conn = new OleDbConnection(cs))
                {
                    conn.Open();
                    OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dt.Columns.Count < expectedColumnCount)
                    {
                        MessageBox.Show(
                            $"فایل اکسل باید حداقل {expectedColumnCount} ستون داشته باشد." +
                            $"\nتعداد ستون‌های فعلی: {dt.Columns.Count}",
                            "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    dgvPreview.DataSource = dt;
                    foreach (DataGridViewColumn col in dgvPreview.Columns)
                    {
                        col.Width = 180;
                        col.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در خواندن فایل اکسل: " + ex.Message,
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                File.AppendAllText("import_errors.log",
                    $"[{DateTime.Now}] خطا در LoadExcelPreview: {ex.Message}\n");
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text) || !File.Exists(txtFilePath.Text))
            {
                MessageBox.Show("لطفاً فایل اکسل معتبر انتخاب کنید.",
                    "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string cs = BuildOleDbConnectionString(txtFilePath.Text);
                using (OleDbConnection conn = new OleDbConnection(cs))
                {
                    conn.Open();
                    OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", conn);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dt.Columns.Count < expectedColumnCount)
                    {
                        MessageBox.Show(
                            $"فایل اکسل باید حداقل {expectedColumnCount} ستون داشته باشد." +
                            $"\nتعداد ستون‌های فعلی: {dt.Columns.Count}",
                            "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    int successCount = 0;
                    int failCount = 0;
                    List<string> errors = new List<string>();

                    foreach (DataRow row in dt.Rows)
                    {
                        // رد کردن ردیف‌های خالی
                        if (IsRowEmpty(row)) continue;

                        string errorMsg;
                        if (ImportRowToDatabase(row, out errorMsg))
                            successCount++;
                        else
                        {
                            failCount++;
                            if (!string.IsNullOrEmpty(errorMsg))
                                errors.Add(errorMsg);
                        }
                    }

                    string resultMsg = $"✅ {successCount} رکورد با موفقیت وارد شد.";
                    if (failCount > 0)
                        resultMsg += $"\n⚠️ {failCount} رکورد با خطا مواجه شد.";
                    if (errors.Count > 0)
                        resultMsg += "\n\nخطاها:\n" + string.Join("\n", errors.Take(5));

                    MessageBox.Show(resultMsg, "نتیجه وارد کردن",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // *** اینجاست که مشکل لود نشدن حل می‌شه ***
                    // به‌جای بستن فرم، رویداد ImportCompleted را fire می‌کنیم
                    // تا فرم والد بتواند داده‌ها را دوباره لود کند
                    ImportCompleted?.Invoke(this, EventArgs.Empty);

                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در وارد کردن داده‌ها: " + ex.Message,
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                File.AppendAllText("import_errors.log",
                    $"[{DateTime.Now}] خطا در BtnImport_Click: {ex.Message}\n");
            }
        }

        private bool IsRowEmpty(DataRow row)
        {
            return row.ItemArray.All(field =>
                field == null || field == DBNull.Value ||
                string.IsNullOrWhiteSpace(field.ToString()));
        }

        private bool ImportRowToDatabase(DataRow row, out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {
                string provinceName = GetSafeString(row, 0);
                string cityName = GetSafeString(row, 1);
                string affairName = GetSafeString(row, 2);   // امور
                string deptName = GetSafeString(row, 3);   // ناحیه بهره‌برداری
                string districtName = GetSafeString(row, 4);
                string postName = GetSafeString(row, 5);
                string voltageName = GetSafeString(row, 6);
                string workShiftName = GetSafeString(row, 7);
                string genderName = GetSafeString(row, 8);
                string firstName = GetSafeString(row, 9);
                string lastName = GetSafeString(row, 10);
                string fatherName = GetSafeString(row, 11);
                string personnelNumber = GetSafeString(row, 12);
                string nationalID = GetSafeString(row, 13);
                string mobileNumber = GetSafeString(row, 14);
                DateTime? birthDate = ParseExcelDate(GetSafeString(row, 15));
                DateTime? hireDate = ParseExcelDate(GetSafeString(row, 16));
                DateTime? startDateOp = ParseExcelDate(GetSafeString(row, 17));
                string contractTypeName = GetSafeString(row, 18);
                string jobLevelName = GetSafeString(row, 19);
                string companyName = GetSafeString(row, 20);
                string degreeName = GetSafeString(row, 21);
                string degreeFieldName = GetSafeString(row, 22);
                string mainJobTitleName = GetSafeString(row, 23);
                string currentActName = GetSafeString(row, 24);
                string inconsistencyStr = GetSafeString(row, 25);
                string description = row.Table.Columns.Count > 26
                                           ? GetSafeString(row, 26)
                                           : string.Empty;

                // اعتبارسنجی فیلدهای ضروری
                if (string.IsNullOrEmpty(firstName) || string.IsNullOrEmpty(lastName) ||
                    string.IsNullOrEmpty(nationalID))
                {
                    errorMessage = $"نام/نام‌خانوادگی/کدملی خالی - ردیف نادیده گرفته شد";
                    return false;
                }

                if (!birthDate.HasValue || !hireDate.HasValue || !startDateOp.HasValue)
                {
                    errorMessage = $"تاریخ نامعتبر برای: {firstName} {lastName}";
                    return false;
                }

                // *** resolve کردن ID ها با آستانه مناسب‌تر ***
                object provinceID = FindClosestID(provinces, "ProvinceName", "ProvinceID", provinceName, maxDistance: 4);
                object cityID = FindClosestID(cities, "CityName", "CityID", cityName, maxDistance: 4);
                object affairID = FindClosestID(affairs, "AffairName", "AffairID", affairName, maxDistance: 6);
                object deptID = FindClosestID(depts, "DeptName", "DeptID", deptName, maxDistance: 6);
                object districtID = FindClosestID(districts, "DistrictName", "DistrictID", districtName, maxDistance: 4);
                object postNameID = FindClosestID(postNames, "PostName", "PostNameID", postName, maxDistance: 5);
                object voltageID = FindClosestID(voltages, "VoltageName", "VoltageID", voltageName, maxDistance: 3);
                object workShiftID = FindClosestID(workShifts, "WorkShiftName", "WorkShiftID", workShiftName, maxDistance: 3);
                object genderID = FindClosestID(genders, "GenderName", "GenderID", genderName, maxDistance: 2);
                object contractTypeID = FindClosestID(contractTypes, "ContractTypeName", "ContractTypeID", contractTypeName, maxDistance: 5);
                object jobLevelID = FindClosestID(jobLevels, "JobLevelName", "JobLevelID", jobLevelName, maxDistance: 4);
                object companyID = FindClosestID(companies, "CompanyName", "CompanyID", companyName, maxDistance: 6);
                object degreeID = FindClosestID(degrees, "DegreeName", "DegreeID", degreeName, maxDistance: 3);
                object degreeFieldID = FindClosestID(degreeFields, "DegreeFieldName", "DegreeFieldID", degreeFieldName, maxDistance: 5);
                object mainJobTitleID = FindClosestID(chartAffairs, "ChartName", "ChartID", mainJobTitleName, maxDistance: 6);
                object currentActID = FindClosestID(chartAffairs, "ChartName", "ChartID", currentActName, maxDistance: 6);

                bool inconsistency = !string.IsNullOrEmpty(inconsistencyStr)
                    ? (inconsistencyStr.Trim() == "دارد")
                    : (mainJobTitleName != currentActName);

                string query =
                    "INSERT INTO Personnel " +
                    "(ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, " +
                    " VoltageID, WorkShiftID, GenderID, FirstName, LastName, FatherName, " +
                    " PersonnelNumber, NationalID, MobileNumber, BirthDate, HireDate, " +
                    " StartDateOperation, ContractTypeID, JobLevelID, CompanyID, " +
                    " DegreeID, DegreeFieldID, MainJobTitle, CurrentActivity, " +
                    " Inconsistency, Description) " +
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

                OleDbParameter[] parameters = new OleDbParameter[]
                {
                    new OleDbParameter("?", provinceID     ?? DBNull.Value),
                    new OleDbParameter("?", cityID         ?? DBNull.Value),
                    new OleDbParameter("?", affairID       ?? DBNull.Value),
                    new OleDbParameter("?", deptID         ?? DBNull.Value),
                    new OleDbParameter("?", districtID     ?? DBNull.Value),
                    new OleDbParameter("?", postNameID     ?? DBNull.Value),
                    new OleDbParameter("?", voltageID      ?? DBNull.Value),
                    new OleDbParameter("?", workShiftID    ?? DBNull.Value),
                    new OleDbParameter("?", genderID       ?? DBNull.Value),
                    new OleDbParameter("?", firstName),
                    new OleDbParameter("?", lastName),
                    new OleDbParameter("?", fatherName),
                    new OleDbParameter("?", personnelNumber),
                    new OleDbParameter("?", nationalID),
                    new OleDbParameter("?", mobileNumber),
                    new OleDbParameter("?", birthDate.Value.ToString("yyyy-MM-dd")),
                    new OleDbParameter("?", hireDate.Value.ToString("yyyy-MM-dd")),
                    new OleDbParameter("?", startDateOp.Value.ToString("yyyy-MM-dd")),
                    new OleDbParameter("?", contractTypeID ?? DBNull.Value),
                    new OleDbParameter("?", jobLevelID     ?? DBNull.Value),
                    new OleDbParameter("?", companyID      ?? DBNull.Value),
                    new OleDbParameter("?", degreeID       ?? DBNull.Value),
                    new OleDbParameter("?", degreeFieldID  ?? DBNull.Value),
                    new OleDbParameter("?", mainJobTitleID ?? DBNull.Value),
                    new OleDbParameter("?", currentActID   ?? DBNull.Value),
                    new OleDbParameter("?", inconsistency ? 1 : 0),
                    new OleDbParameter("?", string.IsNullOrEmpty(description)
                                            ? (object)DBNull.Value : description)
                };

                db.ExecuteNonQuery(query, parameters);
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                File.AppendAllText("import_errors.log",
                    $"[{DateTime.Now}] خطا در ImportRowToDatabase: {ex.Message}\n");
                return false;
            }
        }

        private string GetSafeString(DataRow row, int colIndex)
        {
            if (colIndex < row.Table.Columns.Count && row[colIndex] != DBNull.Value)
                return row[colIndex].ToString().Trim();
            return string.Empty;
        }

        private DateTime? ParseExcelDate(string dateValue)
        {
            if (string.IsNullOrEmpty(dateValue)) return null;

            // عدد سریال اکسل
            if (double.TryParse(dateValue, out double serialDate))
            {
                try { return DateTime.FromOADate(serialDate); }
                catch { /* ادامه */ }
            }

            string[] formats = {
                "yyyy/MM/dd", "yyyy-MM-dd", "dd/MM/yyyy",
                "MM/dd/yyyy", "yyyy/M/d", "d/M/yyyy",
                "yy/MM/dd", "dd-MM-yyyy"
            };

            if (DateTime.TryParseExact(dateValue, formats,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out DateTime result))
                return result;

            if (DateTime.TryParse(dateValue, out result))
                return result;

            return null;
        }

        private string Normalize(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            // حذف zero-width non-joiner و فضاهای اضافه، یکسان‌سازی
            return s.Replace("\u200C", "")
                    .Replace("\u200B", "")
                    .Replace(" ", "")
                    .Replace("ي", "ی")
                    .Replace("ك", "ک")
                    .ToLowerInvariant();
        }

        private int LevenshteinDistance(string s1, string s2)
        {
            int[,] m = new int[s1.Length + 1, s2.Length + 1];
            for (int i = 0; i <= s1.Length; i++) m[i, 0] = i;
            for (int j = 0; j <= s2.Length; j++) m[0, j] = j;
            for (int i = 1; i <= s1.Length; i++)
                for (int j = 1; j <= s2.Length; j++)
                {
                    int cost = s1[i - 1] == s2[j - 1] ? 0 : 1;
                    m[i, j] = Math.Min(
                        Math.Min(m[i - 1, j] + 1, m[i, j - 1] + 1),
                        m[i - 1, j - 1] + cost);
                }
            return m[s1.Length, s2.Length];
        }

        /// <param name="maxDistance">آستانه قابل تنظیم برای هر جدول</param>
        private object FindClosestID(DataTable table, string nameCol, string idCol,
                                     string searchValue, int maxDistance = 4)
        {
            string normalizedSearch = Normalize(searchValue);
            if (string.IsNullOrEmpty(normalizedSearch)) return DBNull.Value;

            // ابتدا تطابق دقیق را بررسی می‌کنیم
            foreach (DataRow r in table.Rows)
            {
                if (Normalize(r[nameCol]?.ToString() ?? "") == normalizedSearch)
                    return r[idCol];
            }

            // سپس جستجوی Fuzzy
            int minDistance = int.MaxValue;
            object closestID = DBNull.Value;
            foreach (DataRow r in table.Rows)
            {
                string val = Normalize(r[nameCol]?.ToString() ?? "");
                int dist = LevenshteinDistance(normalizedSearch, val);
                if (dist < minDistance)
                {
                    minDistance = dist;
                    closestID = r[idCol];
                }
            }

            return minDistance <= maxDistance ? closestID : DBNull.Value;
        }

        private string BuildOleDbConnectionString(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLower();
            string props = ext == ".xls"
                ? "Excel 8.0;HDR=YES"
                : "Excel 12.0 Xml;HDR=YES";
            return $"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   $"Data Source={filePath};" +
                   $"Extended Properties='{props}';";
        }
    }
}
