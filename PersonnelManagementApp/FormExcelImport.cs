using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace PersonnelManagementApp
{
    // ✅ فرم وارد کردن اطلاعات پرسنلی از فایل اکسل
    public class FormExcelImport : Form
    {
        private readonly DbHelper _db;
        private string _selectedFilePath;
        private static readonly PersianCalendar _pc = new PersianCalendar();

        // UI controls
        private Label lblFile;
        private Button btnImport;
        private ProgressBar progressBar;
        private Label lblProgress;
        private DataGridView dgvPreview;
        private Label lblInfo;

        // ✅ کش مقادیر جدول‌های پشتیبان: نام → ID
        // OrdinalIgnoreCase برای جلوگیری از مشکل حروف بزرگ/کوچک
        private readonly Dictionary<string, int> _cProvinces    = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cCities       = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cAffairs      = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cDepts        = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cDistricts    = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cPostNames    = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cVoltages     = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cWorkShifts   = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cGenders      = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cContracts    = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cJobLevels    = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cCompanies    = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cDegrees      = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cDegreeFields = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cChartAffairs = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, int> _cStatuses     = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // ایندکس ستون‌ها در اکسل
        private const int COL_PROVINCE        = 0;   // استان
        private const int COL_CITY            = 1;   // شهر
        private const int COL_AFFAIR          = 2;   // امور انتقال
        private const int COL_DEPT            = 3;   // اداره
        private const int COL_DISTRICT        = 4;   // ناحیه بهره‌برداری
        private const int COL_POST_NAME       = 5;   // نام پست
        private const int COL_VOLTAGE         = 6;   // سطح ولتاژ
        private const int COL_WORKSHIFT       = 7;   // روزکار/نوبتکار
        private const int COL_GENDER          = 8;   // جنسیت
        private const int COL_FIRSTNAME       = 9;   // نام
        private const int COL_LASTNAME        = 10;  // نام خانوادگی
        private const int COL_FATHERNAME      = 11;  // نام پدر
        private const int COL_PERSONNELNUMBER = 12;  // شماره پرسنلی
        private const int COL_NATIONALID      = 13;  // کدملی
        private const int COL_MOBILE          = 14;  // موبایل
        private const int COL_BIRTHDATE       = 15;  // تاریخ تولد
        private const int COL_HIREDATE        = 16;  // تاریخ استخدام
        private const int COL_STARTDATE       = 17;  // تاریخ شروع بکار
        private const int COL_CONTRACTTYPE    = 18;  // نوع قرارداد
        private const int COL_JOBLEVEL        = 19;  // سطح شغل
        private const int COL_COMPANY         = 20;  // شرکت
        private const int COL_DEGREE          = 21;  // مدرک
        private const int COL_DEGREEFIELD     = 22;  // رشته تحصیلی
        private const int COL_MAINJOB         = 23;  // عنوان شغلی اصلی
        private const int COL_CURRENTACTIVITY = 24;  // فعالیت فعلی
        private const int COL_STATUS          = 25;  // وضعیت

        public FormExcelImport(DbHelper db)
        {
            _db = db;
            BuildUI();
        }

        // ───────────────────────────────────────────────────────
        // ساخت UI
        // ───────────────────────────────────────────────────────
        private void BuildUI()
        {
            this.Text = "📥  وارد کردن اطلاعات از فایل اکسل";
            this.Size = new Size(1280, 780);
            this.MinimumSize = new Size(900, 600);
            this.RightToLeft = RightToLeft.Yes;
            this.RightToLeftLayout = true;
            this.StartPosition = FormStartPosition.CenterParent;

            // ―― پنل بالا
            Panel topPanel = new Panel { Dock = DockStyle.Top, Height = 75, BackColor = Color.FromArgb(245, 248, 255) };

            Button btnBrowse = new Button
            {
                Text = "📂  انتخاب فایل اکسل",
                Location = new Point(12, 20),
                Size = new Size(190, 38),
                BackColor = Color.SteelBlue,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("B Nazanin", 10F)
            };
            btnBrowse.FlatAppearance.BorderSize = 0;
            btnBrowse.Click += BtnBrowse_Click;

            lblFile = new Label
            {
                Text = "هنوز فایلی انتخاب نشده است...",
                Location = new Point(215, 28),
                Size = new Size(640, 22),
                ForeColor = Color.Gray,
                Font = new Font("B Nazanin", 9F)
            };

            btnImport = new Button
            {
                Text = "✅  وارد کردن به دیتابیس",
                Location = new Point(870, 20),
                Size = new Size(220, 38),
                BackColor = Color.ForestGreen,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("B Nazanin", 10F),
                Enabled = false
            };
            btnImport.FlatAppearance.BorderSize = 0;
            btnImport.Click += BtnImport_Click;

            Button btnClose = new Button
            {
                Text = "✖  بستن",
                Location = new Point(1100, 20),
                Size = new Size(120, 38),
                BackColor = Color.Crimson,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("B Nazanin", 10F)
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => this.Close();

            topPanel.Controls.AddRange(new Control[] { btnBrowse, lblFile, btnImport, btnClose });

            // ―― پنل اطلاعات
            lblInfo = new Label
            {
                Dock = DockStyle.Top,
                Height = 48,
                Text = "ℹ️  ستون‌های مورد انتظار: استان | شهر | امور | اداره | ناحیه | نام پست | ولتاژ | شیفت | جنسیت | نام | نام خانوادگی | نام پدر | ش.پرسنلی | کدملی | موبایل | تاریخ تولد | تاریخ استخدام | تاریخ شروع | نوع قرارداد | سطح شغل | شرکت | مدرک | رشته | عنوان شغلی اصلی | فعالیت فعلی | وضعیت",
                ForeColor = Color.FromArgb(50, 80, 160),
                BackColor = Color.FromArgb(230, 240, 255),
                Padding = new Padding(10, 5, 10, 5),
                Font = new Font("B Nazanin", 8.5F)
            };

            // ―― پنل پیشرفت
            Panel progressPanel = new Panel { Dock = DockStyle.Top, Height = 38, BackColor = Color.White };
            progressBar = new ProgressBar { Location = new Point(10, 7), Size = new Size(850, 24), Visible = false };
            lblProgress = new Label { Location = new Point(875, 9), Size = new Size(350, 22), Text = "", Font = new Font("B Nazanin", 9.5F), ForeColor = Color.DarkSlateGray };
            progressPanel.Controls.AddRange(new Control[] { progressBar, lblProgress });

            // ―― DataGridView پیش‌نمایی
            dgvPreview = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                ReadOnly = true,
                AllowUserToAddRows = false,
                RightToLeft = RightToLeft.Yes,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                Font = new Font("B Nazanin", 8.5F),
                RowHeadersVisible = false,
                BorderStyle = BorderStyle.None,
                GridColor = Color.LightGray
            };

            this.Controls.Add(dgvPreview);
            this.Controls.Add(progressPanel);
            this.Controls.Add(lblInfo);
            this.Controls.Add(topPanel);
        }

        // ───────────────────────────────────────────────────────
        // انتخاب فایل و نمایش پیش‌نمایش
        // ───────────────────────────────────────────────────────
        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "فایل اکسل (*.xlsx)|*.xlsx|فایل اکسل قدیمی (*.xls)|*.xls";
                ofd.Title = "انتخاب فایل اکسل پرسنلی";
                if (ofd.ShowDialog() == DialogResult.OK)
                    LoadPreview(ofd.FileName);
            }
        }

        private void LoadPreview(string filePath)
        {
            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("PersonnelManagementApp");
                using (var pkg = new ExcelPackage(new FileInfo(filePath)))
                {
                    var ws = pkg.Workbook.Worksheets[0];
                    if (ws.Dimension == null)
                    {
                        MessageBox.Show("فایل اکسل خالی است!", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    int rows = ws.Dimension.Rows;
                    int cols = ws.Dimension.Columns;
                    DataTable dt = new DataTable();

                    for (int c = 1; c <= cols; c++)
                        dt.Columns.Add(ws.Cells[1, c].Text?.Trim() ?? $"ستون{c}");

                    int preview = Math.Min(rows, 51);
                    for (int r = 2; r <= preview; r++)
                    {
                        var dr = dt.NewRow();
                        for (int c = 1; c <= cols; c++)
                            dr[c - 1] = ws.Cells[r, c].Text?.Trim() ?? "";
                        dt.Rows.Add(dr);
                    }

                    dgvPreview.DataSource = dt;
                    _selectedFilePath = filePath;
                    lblFile.Text = $"✅  {Path.GetFileName(filePath)}   —   مجموع {rows - 1} ردیف داده";
                    lblFile.ForeColor = Color.DarkGreen;
                    btnImport.Enabled = true;

                    if (rows - 1 > 50)
                        lblProgress.Text = $"بیش از 50 ردیف — فقط 50 اول نمایش داده می‌شود";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در خواندن فایل اکسل:\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ───────────────────────────────────────────────────────
        // شروع وارد کردن
        // ───────────────────────────────────────────────────────
        private void BtnImport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_selectedFilePath)) return;

            var res = MessageBox.Show(
                "⚠️ آیا مطمئن هستید که می‌خواهید اطلاعات اکسل را به دیتابیس وارد کنید?\n\n"
                + "• ردیف‌هایی که کدملی تکراری دارند نادیده گرفته می‌شوند.\n"
                + "• تاریخ‌ها به شمسی تبدیل می‌شوند.\n"
                + "• مقادیر خالی با مقادیر پیش‌فرض جایگزینی می‌شوند.",
                "تأیید وارد کردن", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (res != DialogResult.Yes) return;

            btnImport.Enabled = false;
            progressBar.Visible = true;
            lblProgress.Text = "⏳ در حال آماده‌سازی...";
            Application.DoEvents();

            try
            {
                LoadAllCaches();
                DoImport(_selectedFilePath);
            }
            finally
            {
                btnImport.Enabled = true;
                progressBar.Visible = false;
            }
        }

        private void DoImport(string filePath)
        {
            int success = 0, skipped = 0, failed = 0;
            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("PersonnelManagementApp");
                using (var pkg = new ExcelPackage(new FileInfo(filePath)))
                {
                    var ws = pkg.Workbook.Worksheets[0];
                    int totalRows = ws.Dimension.Rows - 1;
                    progressBar.Maximum = Math.Max(totalRows, 1);
                    progressBar.Value = 0;

                    var existingNIDs = GetExistingNationalIDs();

                    for (int r = 2; r <= ws.Dimension.Rows; r++)
                    {
                        try
                        {
                            int totalCols = ws.Dimension.Columns;
                            string[] cells = new string[Math.Max(totalCols, 27)];
                            for (int c = 1; c <= totalCols; c++)
                                cells[c - 1] = ws.Cells[r, c].Text?.Trim() ?? "";

                            // رد کردن ردیف‌های کاملاً خالی
                            if (string.IsNullOrWhiteSpace(cells[COL_FIRSTNAME]) &&
                                string.IsNullOrWhiteSpace(cells[COL_NATIONALID])) continue;

                            // رد کردن کدملی تکراری
                            string nid = cells[COL_NATIONALID]?.Trim();
                            if (!string.IsNullOrWhiteSpace(nid) && existingNIDs.Contains(nid))
                            { skipped++; continue; }

                            InsertRecord(cells);
                            success++;
                            if (!string.IsNullOrWhiteSpace(nid)) existingNIDs.Add(nid);
                        }
                        catch { failed++; }

                        progressBar.Value = Math.Min(r - 1, progressBar.Maximum);
                        lblProgress.Text = $"✅ موفق: {success}  |  ⏭ تکراری: {skipped}  |  ❌ ناموفق: {failed}";
                        Application.DoEvents();
                    }
                }

                progressBar.Value = progressBar.Maximum;
                MessageBox.Show(
                    $"✅ وارد کردن اطلاعات به پایان رسید!\n\n"
                    + $"🟢  موفق: {success} ردیف\n"
                    + $"⏭️  تکراری (نادیده گرفته): {skipped} ردیف\n"
                    + $"🔴  ناموفق: {failed} ردیف",
                    "نتیجه وارد کردن", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در وارد کردن:\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ───────────────────────────────────────────────────────
        // ✅ درج رکورد — نسخه اصلاح‌شده
        // ───────────────────────────────────────────────────────
        private void InsertRecord(string[] cells)
        {
            // ── فیلدهای عمومی با fuzzy matching
            int provinceId    = GetOrCreateFuzzy(_cProvinces,    GetCell(cells, COL_PROVINCE),              "Provinces",            "ProvinceID",    "ProvinceName");
            int cityId        = GetOrCreateFuzzy(_cCities,       GetCell(cells, COL_CITY),                  "Cities",               "CityID",        "CityName");
            int affairId      = GetOrCreateFuzzy(_cAffairs,      GetCell(cells, COL_AFFAIR),                "TransferAffairs",      "AffairID",      "AffairName");

            // ✅ اداره عملیاتی: با AffairID ذخیره می‌شود (ForeignKey حفظ می‌شود)
            int deptId        = GetOrCreateWithAffair(_cDepts,    GetCell(cells, COL_DEPT),                 "OperationDepartments", "DeptID",        "DeptName",   affairId);

            // ✅ ناحیه بهره‌برداری: با AffairID ذخیره می‌شود ← این باگ اصلی بود
            int districtId    = GetOrCreateWithAffair(_cDistricts, GetCell(cells, COL_DISTRICT),            "Districts",            "DistrictID",    "DistrictName", affairId);

            int postNameId    = GetOrCreateFuzzy(_cPostNames,    GetCell(cells, COL_POST_NAME),             "PostsNames",           "PostNameID",    "PostName");
            int voltageId     = GetOrCreateFuzzy(_cVoltages,     GetCell(cells, COL_VOLTAGE),               "VoltageLevels",        "VoltageID",     "VoltageName");
            int workShiftId   = GetOrCreateFuzzy(_cWorkShifts,   GetCell(cells, COL_WORKSHIFT),             "WorkShift",            "WorkShiftID",   "WorkShiftName");
            int genderId      = GetOrCreateFuzzy(_cGenders,      GetCell(cells, COL_GENDER),                "Gender",               "GenderID",      "GenderName");
            int contractId    = GetOrCreateFuzzy(_cContracts,    GetCell(cells, COL_CONTRACTTYPE),          "ContractType",         "ContractTypeID","ContractTypeName");
            int jobLevelId    = GetOrCreateFuzzy(_cJobLevels,    GetCell(cells, COL_JOBLEVEL),              "JobLevel",             "JobLevelID",    "JobLevelName");
            int companyId     = GetOrCreateFuzzy(_cCompanies,    GetCell(cells, COL_COMPANY),               "Company",              "CompanyID",     "CompanyName");
            int degreeId      = GetOrCreateFuzzy(_cDegrees,      GetCell(cells, COL_DEGREE),                "Degree",               "DegreeID",      "DegreeName");
            int degreeFieldId = GetOrCreateFuzzy(_cDegreeFields, GetJobCell(cells, COL_DEGREEFIELD),        "DegreeField",          "DegreeFieldID", "DegreeFieldName");
            int mainJobId     = GetOrCreateChart(GetJobCell(cells, COL_MAINJOB),                            affairId);
            int currentActId  = GetOrCreateChart(GetJobCell(cells, COL_CURRENTACTIVITY),                    affairId);
            int statusId      = GetOrCreateFuzzy(_cStatuses,     GetCell(cells, COL_STATUS, "حاضر"),       "StatusPresence",       "StatusID",      "StatusName");

            string birthDate  = ParseDate(GetCell(cells, COL_BIRTHDATE,  ""));
            string hireDate   = ParseDate(GetCell(cells, COL_HIREDATE,   ""));
            string startDate  = ParseDate(GetCell(cells, COL_STARTDATE,  ""));

            string sql = @"INSERT INTO Personnel
                (ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, VoltageID, WorkShiftID, GenderID,
                 FirstName, LastName, FatherName, PersonnelNumber, NationalID, MobileNumber,
                 BirthDate, HireDate, StartDateOperation, ContractTypeID, JobLevelID, CompanyID,
                 DegreeID, DegreeFieldID, MainJobTitle, CurrentActivity, StatusID)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

            var ps = new OleDbParameter[]
            {
                new OleDbParameter("?", provinceId),
                new OleDbParameter("?", cityId),
                new OleDbParameter("?", affairId),
                new OleDbParameter("?", deptId),
                new OleDbParameter("?", districtId),
                new OleDbParameter("?", postNameId),
                new OleDbParameter("?", voltageId),
                new OleDbParameter("?", workShiftId),
                new OleDbParameter("?", genderId),
                new OleDbParameter("?", GetCell(cells, COL_FIRSTNAME)),
                new OleDbParameter("?", GetCell(cells, COL_LASTNAME)),
                new OleDbParameter("?", GetCell(cells, COL_FATHERNAME)),
                new OleDbParameter("?", GetCell(cells, COL_PERSONNELNUMBER, "داده‌ای وجود ندارد")),
                new OleDbParameter("?", GetCell(cells, COL_NATIONALID,     "داده‌ای وجود ندارد")),
                new OleDbParameter("?", GetCell(cells, COL_MOBILE,         "داده‌ای وجود ندارد")),
                new OleDbParameter("?", birthDate),
                new OleDbParameter("?", hireDate),
                new OleDbParameter("?", startDate),
                new OleDbParameter("?", contractId),
                new OleDbParameter("?", jobLevelId),
                new OleDbParameter("?", companyId),
                new OleDbParameter("?", degreeId),
                new OleDbParameter("?", degreeFieldId),
                new OleDbParameter("?", mainJobId),
                new OleDbParameter("?", currentActId),
                new OleDbParameter("?", statusId)
            };

            _db.ExecuteNonQuery(sql, ps);
        }

        // ───────────────────────────────────────────────────────
        // ✅ Fuzzy Matching — برگرفته از کد قدیمی PostDatabaseManager
        // ───────────────────────────────────────────────────────

        /// نرمال‌سازی متن فارسی/عربی: حذف نیم‌فاصله، یکسان‌سازی ی/ک عربی
        private string Normalize(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s
                .Replace("\u200C", "")   // حذف نیم‌فاصله ZWJ
                .Replace("\u200F", "")   // حذف RLM
                .Replace("ي", "ی")        // ی عربی → فارسی
                .Replace("ك", "ک")        // ک عربی → فارسی
                .Replace("  ", " ")       // فاصله‌های دوگانه
                .Trim();
        }

        /// فاصله Levenshtein — برای تطبیق متن‌های مشابه
        private int LevenshteinDistance(string s1, string s2)
        {
            int[,] matrix = new int[s1.Length + 1, s2.Length + 1];
            for (int i = 0; i <= s1.Length; i++) matrix[i, 0] = i;
            for (int j = 0; j <= s2.Length; j++) matrix[0, j] = j;
            for (int i = 1; i <= s1.Length; i++)
                for (int j = 1; j <= s2.Length; j++)
                {
                    int cost = (s1[i - 1] == s2[j - 1]) ? 0 : 1;
                    matrix[i, j] = Math.Min(
                        Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                        matrix[i - 1, j - 1] + cost);
                }
            return matrix[s1.Length, s2.Length];
        }

        /// جستجوی fuzzy در cache — اول دقیق، سپس Levenshtein
        private int FindFuzzyInCache(Dictionary<string, int> cache, string searchName, int maxDistance = 3)
        {
            string normalizedSearch = Normalize(searchName);
            if (string.IsNullOrEmpty(normalizedSearch)) return 0;

            // 1) مطابقت دقیق پس از نرمال‌سازی
            foreach (var kv in cache)
                if (Normalize(kv.Key) == normalizedSearch) return kv.Value;

            // 2) مطابقت fuzzy با الگوریتم Levenshtein
            int minDist = int.MaxValue;
            int closestId = 0;
            foreach (var kv in cache)
            {
                int dist = LevenshteinDistance(normalizedSearch, Normalize(kv.Key));
                if (dist < minDist) { minDist = dist; closestId = kv.Value; }
            }
            return (minDist <= maxDistance) ? closestId : 0;
        }

        /// دریافت ID با fuzzy matching — اگر پیدا نشد INSERT ساده (بدون FK)
        private int GetOrCreateFuzzy(Dictionary<string, int> cache, string name,
                                      string table, string idCol, string nameCol)
        {
            int found = FindFuzzyInCache(cache, name);
            if (found > 0) return found;

            int newId = ExecuteInsertGetId(
                $"INSERT INTO {table} ({nameCol}) VALUES (?)",
                new OleDbParameter[] { new OleDbParameter("?", name) });

            if (newId > 0) cache[name] = newId;
            return newId;
        }

        /// ✅ دریافت ID با fuzzy matching — اگر پیدا نشد INSERT با AffairID
        /// این متد برای Districts و OperationDepartments استفاده می‌شود
        /// تا رابطه FK با جدول TransferAffairs حفظ شود
        private int GetOrCreateWithAffair(Dictionary<string, int> cache, string name,
                                           string table, string idCol, string nameCol,
                                           int affairId)
        {
            int found = FindFuzzyInCache(cache, name);
            if (found > 0) return found;

            // INSERT با AffairID تا ناحیه/اداره به امور مرتبط شود
            int newId = ExecuteInsertGetId(
                $"INSERT INTO {table} ({nameCol}, AffairID) VALUES (?, ?)",
                new OleDbParameter[]
                {
                    new OleDbParameter("?", name),
                    new OleDbParameter("?", affairId)
                });

            if (newId > 0) cache[name] = newId;
            return newId;
        }

        /// خاص ChartAffairs1 که نیاز به AffairID دارد
        private int GetOrCreateChart(string chartName, int affairId)
        {
            if (_cChartAffairs.TryGetValue(chartName, out int id)) return id;

            int found = FindFuzzyInCache(_cChartAffairs, chartName);
            if (found > 0) return found;

            int newId = ExecuteInsertGetId(
                "INSERT INTO ChartAffairs1 (AffairID, ChartName) VALUES (?, ?)",
                new OleDbParameter[]
                {
                    new OleDbParameter("?", affairId),
                    new OleDbParameter("?", chartName)
                });

            if (newId > 0) _cChartAffairs[chartName] = newId;
            return newId;
        }

        /// اجرای INSERT و دریافت @@IDENTITY در یک اتصال (ضروری برای Access)
        private int ExecuteInsertGetId(string insertSql, OleDbParameter[] ps)
        {
            using (var conn = new OleDbConnection(_db.GetConnectionString_Public()))
            {
                conn.Open();
                using (var cmd = new OleDbCommand(insertSql, conn))
                {
                    if (ps != null) cmd.Parameters.AddRange(ps);
                    cmd.ExecuteNonQuery();
                }
                using (var cmd2 = new OleDbCommand("SELECT @@IDENTITY", conn))
                {
                    object result = cmd2.ExecuteScalar();
                    return (result != null && result != DBNull.Value) ? Convert.ToInt32(result) : 0;
                }
            }
        }

        // ───────────────────────────────────────────────────────
        // Cache Loader
        // ───────────────────────────────────────────────────────
        private void LoadAllCaches()
        {
            _cProvinces.Clear();
            _cCities.Clear();
            _cAffairs.Clear();
            _cDepts.Clear();
            _cDistricts.Clear();
            _cPostNames.Clear();
            _cVoltages.Clear();
            _cWorkShifts.Clear();
            _cGenders.Clear();
            _cContracts.Clear();
            _cJobLevels.Clear();
            _cCompanies.Clear();
            _cDegrees.Clear();
            _cDegreeFields.Clear();
            _cChartAffairs.Clear();
            _cStatuses.Clear();

            FillCache("SELECT ProvinceID, ProvinceName FROM Provinces",              "ProvinceID",    "ProvinceName",    _cProvinces);
            FillCache("SELECT CityID, CityName FROM Cities",                         "CityID",        "CityName",        _cCities);
            FillCache("SELECT AffairID, AffairName FROM TransferAffairs",            "AffairID",      "AffairName",      _cAffairs);
            FillCache("SELECT DeptID, DeptName FROM OperationDepartments",           "DeptID",        "DeptName",        _cDepts);
            FillCache("SELECT DistrictID, DistrictName FROM Districts",              "DistrictID",    "DistrictName",    _cDistricts);
            FillCache("SELECT PostNameID, PostName FROM PostsNames",                 "PostNameID",    "PostName",        _cPostNames);
            FillCache("SELECT VoltageID, VoltageName FROM VoltageLevels",            "VoltageID",     "VoltageName",     _cVoltages);
            FillCache("SELECT WorkShiftID, WorkShiftName FROM WorkShift",            "WorkShiftID",   "WorkShiftName",   _cWorkShifts);
            FillCache("SELECT GenderID, GenderName FROM Gender",                     "GenderID",      "GenderName",      _cGenders);
            FillCache("SELECT ContractTypeID, ContractTypeName FROM ContractType",   "ContractTypeID","ContractTypeName",_cContracts);
            FillCache("SELECT JobLevelID, JobLevelName FROM JobLevel",               "JobLevelID",    "JobLevelName",    _cJobLevels);
            FillCache("SELECT CompanyID, CompanyName FROM Company",                  "CompanyID",     "CompanyName",     _cCompanies);
            FillCache("SELECT DegreeID, DegreeName FROM Degree",                     "DegreeID",      "DegreeName",      _cDegrees);
            FillCache("SELECT DegreeFieldID, DegreeFieldName FROM DegreeField",      "DegreeFieldID", "DegreeFieldName", _cDegreeFields);
            FillCache("SELECT ChartID, ChartName FROM ChartAffairs1",                "ChartID",       "ChartName",       _cChartAffairs);
            FillCache("SELECT StatusID, StatusName FROM StatusPresence",             "StatusID",      "StatusName",      _cStatuses);
        }

        private void FillCache(string query, string idCol, string nameCol, Dictionary<string, int> cache)
        {
            try
            {
                var dt = _db.ExecuteQuery(query);
                if (dt == null) return;
                foreach (DataRow row in dt.Rows)
                {
                    string name = row[nameCol]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(name) && !cache.ContainsKey(name))
                        cache[name] = Convert.ToInt32(row[idCol]);
                }
            }
            catch { }
        }

        // ───────────────────────────────────────────────────────
        // متدهای کمکی
        // ───────────────────────────────────────────────────────
        /// دریافت مقدار سلول — اگر خالی بود مقدار پیش‌فرض برگردان
        private string GetCell(string[] cells, int idx, string def = "داده‌ای وجود ندارد")
        {
            if (idx >= cells.Length) return def;
            string v = cells[idx]?.Trim();
            return string.IsNullOrWhiteSpace(v) ? def : v;
        }

        /// برای ستون‌های شغلی — مقدار پیش‌فرض غیرمرتبط
        private string GetJobCell(string[] cells, int idx)
        {
            string v = GetCell(cells, idx, "");
            return string.IsNullOrWhiteSpace(v) ? "غیرمرتبط" : v;
        }

        /// تبدیل تاریخ میلادی/شمسی به فرمت ذخیره‌سازی شمسی
        private string ParseDate(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw) || raw == "داده‌ای وجود ندارد")
                return "1300/01/01";
            try
            {
                string[] p = raw.Trim().Split(new char[] { '/', '-', '.' }, StringSplitOptions.RemoveEmptyEntries);
                if (p.Length == 3)
                {
                    if (int.TryParse(p[0], out int y) &&
                        int.TryParse(p[1], out int m) &&
                        int.TryParse(p[2], out int d))
                    {
                        // میلادی (1800+) → تبدیل به شمسی
                        if (y >= 1800 && m >= 1 && m <= 12 && d >= 1 && d <= 31)
                        {
                            var dt = new DateTime(y, m, d);
                            return $"{_pc.GetYear(dt):0000}/{_pc.GetMonth(dt):00}/{_pc.GetDayOfMonth(dt):00}";
                        }
                        // از قبل شمسی
                        if (y >= 1300 && y <= 1500)
                            return $"{y:0000}/{m:00}/{d:00}";
                        // سال کوتاه مثل 71 → 1371
                        if (y >= 1 && y <= 99)
                            return $"{y + 1300:0000}/{m:00}/{d:00}";
                    }
                    // حالت dd/MM/yy مثل 26/11/71
                    if (int.TryParse(p[0], out int d2) &&
                        int.TryParse(p[1], out int m2) &&
                        int.TryParse(p[2], out int y2) &&
                        d2 > 12 && m2 <= 12)
                    {
                        int fy = y2 < 100 ? y2 + 1300 : y2;
                        return $"{fy:0000}/{m2:00}/{d2:00}";
                    }
                }
            }
            catch { }
            return "1300/01/01";
        }

        private HashSet<string> GetExistingNationalIDs()
        {
            var set = new HashSet<string>(StringComparer.Ordinal);
            var dt = _db.ExecuteQuery("SELECT NationalID FROM Personnel WHERE NationalID IS NOT NULL");
            if (dt == null) return set;
            foreach (DataRow row in dt.Rows)
            {
                string nid = row["NationalID"]?.ToString()?.Trim();
                if (!string.IsNullOrWhiteSpace(nid)) set.Add(nid);
            }
            return set;
        }
    }
}
