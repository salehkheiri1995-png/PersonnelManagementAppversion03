using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace PersonnelManagementApp
{
    /// <summary>
    /// مدیریت یکپارچه پرسنل: جستجو، افزودن، ویرایش، حذف و خروجی اکسل.
    /// </summary>
    public sealed class FormPersonnelManager : Form
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

        private readonly DbHelper db = new();

        private DataTable? _allData;
        private readonly TextBox txtSearch = new();
        private readonly ComboBox cmbProvince = new();
        private DataGridView dgvPersonnel = null!;
        private Label lblStatus = null!;
        private Label lblCount = null!;
        private bool _suppressFilter;

        public FormPersonnelManager()
        {
            BuildUi();
            FontSettings.ApplyFontToForm(this);
            LoadData();
        }

        private void BuildUi()
        {
            Text = "مدیریت پرسنل";
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
            root.Controls.Add(BuildGridPanel(), 0, 2);
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
                Text = "مدیریت پرسنل",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font(FontSettings.FontFamilyName, 17F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var subtitle = new Label
            {
                Text = "جستجو، افزودن، ویرایش، حذف و خروجی پرسنل در یک محیط یکپارچه",
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

            txtSearch.Dock = DockStyle.Fill;
            txtSearch.Font = new Font(FontSettings.FontFamilyName, 9.5F);
            txtSearch.PlaceholderText = "جستجو در نام، نام خانوادگی، شماره پرسنلی، کد ملی، موبایل...";
            txtSearch.TextChanged += (_, _) => ApplyFilter();

            cmbProvince.Dock = DockStyle.Fill;
            cmbProvince.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbProvince.Font = new Font(FontSettings.FontFamilyName, 9.5F);
            cmbProvince.SelectedIndexChanged += ProvinceFilter_Changed;
            cmbProvince.Items.Add("همه استان‌ها");
            cmbProvince.SelectedIndex = 0;

            var btnAdd = CreateToolButton("➕ ثبت پرسنل جدید", SuccessColor);
            btnAdd.Click += (_, _) => OpenEditor(null);

            var btnEdit = CreateToolButton("✏️ ویرایش", PrimaryColor);
            btnEdit.Click += (_, _) => EditSelected();

            var btnDelete = CreateToolButton("🗑️ حذف", DangerColor);
            btnDelete.Click += (_, _) => DeleteSelected();

            var btnExcel = CreateToolButton("📥 خروجی اکسل", Color.FromArgb(20, 184, 166));
            btnExcel.Click += (_, _) => ExportExcel();

            var btnRefresh = CreateToolButton("🔄 به‌روزرسانی", Color.FromArgb(139, 92, 246));
            btnRefresh.Click += (_, _) => LoadData();

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = CardColor,
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 340F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = CardColor,
                Margin = new Padding(0)
            };
            actions.Controls.Add(btnRefresh);
            actions.Controls.Add(btnExcel);
            actions.Controls.Add(btnDelete);
            actions.Controls.Add(btnEdit);
            actions.Controls.Add(btnAdd);

            layout.Controls.Add(txtSearch, 0, 0);
            layout.Controls.Add(cmbProvince, 1, 0);
            layout.Controls.Add(actions, 2, 0);
            toolbar.Controls.Add(layout);
            return toolbar;
        }

        private static Button CreateToolButton(string text, Color backColor)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = false,
                Width = 165,
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

        private Control BuildGridPanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = CardColor,
                Padding = new Padding(12),
                Margin = new Padding(0)
            };

            dgvPersonnel = new DataGridView
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
                MultiSelect = true,
                AutoGenerateColumns = false,
                RightToLeft = RightToLeft.Yes,
                GridColor = BorderColor,
                EnableHeadersVisualStyles = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvPersonnel.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 64, 175);
            dgvPersonnel.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPersonnel.ColumnHeadersDefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F, FontStyle.Bold);
            dgvPersonnel.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPersonnel.ColumnHeadersHeight = 38;
            dgvPersonnel.DefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F);
            dgvPersonnel.DefaultCellStyle.ForeColor = TextColor;
            dgvPersonnel.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            dgvPersonnel.DefaultCellStyle.SelectionForeColor = TextColor;
            dgvPersonnel.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);
            dgvPersonnel.RowTemplate.Height = 30;

            AddColumn("شناسه پرسنل", "شناسه", 70, typeof(int));
            AddColumn("نام", "نام", 100, typeof(string));
            AddColumn("نام خانوادگی", "نام خانوادگی", 120, typeof(string));
            AddColumn("شماره پرسنلی", "شماره پرسنلی", 100, typeof(string));
            AddColumn("کد ملی", "کد ملی", 110, typeof(string));
            AddColumn("شماره موبایل", "موبایل", 110, typeof(string));
            AddColumn("استان", "استان", 90, typeof(string));
            AddColumn("شهر", "شهر", 100, typeof(string));
            AddColumn("اداره", "اداره", 120, typeof(string));
            AddColumn("ناحیه", "ناحیه", 100, typeof(string));
            AddColumn("نام پست", "نام پست", 120, typeof(string));
            AddColumn("نوع قرارداد", "نوع قرارداد", 110, typeof(string));
            AddColumn("وضعیت حضور", "وضعیت حضور", 100, typeof(string));

            dgvPersonnel.CellDoubleClick += (_, e) =>
            {
                if (e.RowIndex >= 0)
                    EditSelected();
            };

            panel.Controls.Add(dgvPersonnel);
            return panel;
        }

        private void AddColumn(string name, string header, int width, Type type)
        {
            dgvPersonnel.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = header,
                DataPropertyName = name,
                ValueType = type,
                FillWeight = width,
                MinimumWidth = Math.Min(60, width),
                SortMode = DataGridViewColumnSortMode.Automatic
            });
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

            lblStatus = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font(FontSettings.FontFamilyName, 8.5F),
                ForeColor = Color.FromArgb(71, 85, 105),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(18, 0, 18, 0),
                Text = "در حال بارگذاری..."
            };

            lblCount = new Label
            {
                Dock = DockStyle.Right,
                Width = 220,
                Font = new Font(FontSettings.FontFamilyName, 9F, FontStyle.Bold),
                ForeColor = PrimaryColor,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(0, 0, 12, 0),
                Text = "۰ پرسنل"
            };

            statusPanel.Controls.Add(lblStatus);
            statusPanel.Controls.Add(lblCount);
            return statusPanel;
        }

        // ─────────────────────────────────────────────
        // داده‌ها
        // ─────────────────────────────────────────────
        private void LoadData()
        {
            UseWaitCursor = true;
            try
            {
                lblStatus.Text = "در حال بارگذاری پرسنل...";
                _allData = BuildPersonnelTable();
                if (_allData == null || _allData.Rows.Count == 0)
                {
                    lblStatus.Text = "⚠️ داده‌ای در جدول پرسنل یافت نشد.";
                    return;
                }

                FillProvinceFilter();
                ApplyFilter();
                lblStatus.Text = $"✅ {_allData.Rows.Count} پرسنل بارگذاری شد.";
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPersonnelManager.LoadData", ex);
                lblStatus.Text = "❌ خطا در بارگذاری داده‌ها.";
                MessageBox.Show(this, $"خطا در بارگذاری داده‌ها:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        private DataTable? BuildPersonnelTable()
        {
            DataTable? personnel = db.ExecuteQuery(
                "SELECT PersonnelID, ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, VoltageID, " +
                "WorkShiftID, GenderID, FirstName, LastName, FatherName, PersonnelNumber, NationalID, " +
                "MobileNumber, BirthDate, HireDate, StartDateOperation, ContractTypeID, JobLevelID, " +
                "CompanyID, DegreeID, DegreeFieldID, MainJobTitle, CurrentActivity, Inconsistency, " +
                "Description, StatusID FROM Personnel", null, showError: false);
            if (personnel == null)
                return null;

            DataTable? provinces = db.GetProvinces();
            DataTable? cities = db.GetCitiesByProvince(0);
            DataTable? affairs = db.GetAffairsByProvince(0);
            DataTable? depts = db.GetDeptsByAffair(0);
            DataTable? districts = db.GetDistrictsByDept(0);
            DataTable? postNames = db.GetPostNamesByDistrict(0);
            DataTable? contractTypes = db.GetContractTypes();
            DataTable? statusPresence = db.GetStatusPresence();

            var result = new DataTable("PersonnelList");
            result.Columns.Add("شناسه پرسنل", typeof(int));
            result.Columns.Add("نام", typeof(string));
            result.Columns.Add("نام خانوادگی", typeof(string));
            result.Columns.Add("شماره پرسنلی", typeof(string));
            result.Columns.Add("کد ملی", typeof(string));
            result.Columns.Add("شماره موبایل", typeof(string));
            result.Columns.Add("استان", typeof(string));
            result.Columns.Add("شهر", typeof(string));
            result.Columns.Add("اداره", typeof(string));
            result.Columns.Add("ناحیه", typeof(string));
            result.Columns.Add("نام پست", typeof(string));
            result.Columns.Add("نوع قرارداد", typeof(string));
            result.Columns.Add("وضعیت حضور", typeof(string));

            foreach (DataRow p in personnel.Rows)
            {
                DataRow r = result.NewRow();
                r["شناسه پرسنل"] = p["PersonnelID"];
                r["نام"] = p["FirstName"];
                r["نام خانوادگی"] = p["LastName"];
                r["شماره پرسنلی"] = p["PersonnelNumber"];
                r["کد ملی"] = p["NationalID"];
                r["شماره موبایل"] = p["MobileNumber"];
                r["استان"] = LookupName(provinces, "ProvinceID", "ProvinceName", p["ProvinceID"]);
                r["شهر"] = LookupName(cities, "CityID", "CityName", p["CityID"]);
                r["اداره"] = LookupName(depts, "DeptID", "DeptName", p["DeptID"]);
                r["ناحیه"] = LookupName(districts, "DistrictID", "DistrictName", p["DistrictID"]);
                r["نام پست"] = LookupName(postNames, "PostNameID", "PostName", p["PostNameID"]);
                r["نوع قرارداد"] = LookupName(contractTypes, "ContractTypeID", "ContractTypeName", p["ContractTypeID"]);
                r["وضعیت حضور"] = LookupName(statusPresence, "StatusID", "StatusName", p["StatusID"]);
                result.Rows.Add(r);
            }

            return result;
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

        private void ProvinceFilter_Changed(object? sender, EventArgs e)
        {
            if (!_suppressFilter)
                ApplyFilter();
        }

        private void FillProvinceFilter()
        {
            object? previous = cmbProvince.SelectedItem;
            _suppressFilter = true;
            try
            {
                cmbProvince.Items.Clear();
                cmbProvince.Items.Add("همه استان‌ها");
                foreach (string name in _allData!.AsEnumerable()
                    .Select(row => row["استان"]?.ToString() ?? "")
                    .Where(value => !string.IsNullOrWhiteSpace(value) && value != "نامشخص")
                    .Distinct(StringComparer.CurrentCultureIgnoreCase)
                    .OrderBy(value => value, StringComparer.CurrentCulture))
                {
                    cmbProvince.Items.Add(name);
                }
                cmbProvince.SelectedIndex = previous is string p && cmbProvince.Items.Contains(p)
                    ? cmbProvince.Items.IndexOf(p)
                    : 0;
            }
            finally
            {
                _suppressFilter = false;
            }
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            if (_allData == null)
                return;

            string search = txtSearch.Text.Trim();
            string province = cmbProvince.SelectedItem?.ToString() ?? "همه استان‌ها";

            IEnumerable<DataRow> rows = _allData.AsEnumerable();
            if (!string.IsNullOrWhiteSpace(province) && province != "همه استان‌ها")
                rows = rows.Where(row => string.Equals(row["استان"]?.ToString(), province, StringComparison.CurrentCultureIgnoreCase));

            if (!string.IsNullOrWhiteSpace(search))
            {
                rows = rows.Where(row =>
                    (row["نام"]?.ToString() ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase) ||
                    (row["نام خانوادگی"]?.ToString() ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase) ||
                    (row["شماره پرسنلی"]?.ToString() ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase) ||
                    (row["کد ملی"]?.ToString() ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase) ||
                    (row["شماره موبایل"]?.ToString() ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase));
            }

            DataTable filtered = _allData.Clone();
            foreach (DataRow row in rows)
                filtered.ImportRow(row);

            dgvPersonnel.DataSource = filtered;
            lblCount.Text = $"{filtered.Rows.Count:N0} پرسنل";
            lblStatus.Text = $"نمایش {filtered.Rows.Count} از {_allData.Rows.Count} پرسنل  |  " +
                             $"آخرین بروزرسانی: {DateTime.Now:HH:mm:ss}";
        }

        // ─────────────────────────────────────────────
        // عملیات
        // ─────────────────────────────────────────────
        private void OpenEditor(int? personnelId)
        {
            using var editor = new FormPersonnelEditor(personnelId);
            if (editor.ShowDialog(this) == DialogResult.OK)
                LoadData();
        }

        private void EditSelected()
        {
            if (dgvPersonnel.CurrentRow == null || dgvPersonnel.CurrentRow.DataBoundItem is not DataRowView rowView)
            {
                MessageBox.Show(this, "ابتدا یک پرسنل را از جدول انتخاب کنید.", "اطلاع",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int personnelId = Convert.ToInt32(rowView["شناسه پرسنل"]);
            OpenEditor(personnelId);
        }

        private void DeleteSelected()
        {
            var selected = dgvPersonnel.SelectedRows.Cast<DataGridViewRow>()
                .Where(row => row.DataBoundItem is DataRowView)
                .ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, "ابتدا یک یا چند پرسنل را از جدول انتخاب کنید.", "اطلاع",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string question = selected.Count == 1
                ? "آیا مطمئن هستید که می‌خواهید این پرسنل را حذف کنید؟"
                : $"آیا مطمئن هستید که می‌خواهید {selected.Count} پرسنل را حذف کنید؟";

            if (MessageBox.Show(this, question, "تأیید حذف",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            UseWaitCursor = true;
            int deleted = 0;
            try
            {
                foreach (var row in selected)
                {
                    int id = Convert.ToInt32(((DataRowView)row.DataBoundItem!)["شناسه پرسنل"]);
                    int affected = db.ExecuteNonQuery("DELETE FROM Personnel WHERE PersonnelID = ?",
                        new[] { new System.Data.OleDb.OleDbParameter("?", id) }, showError: false);
                    if (affected > 0)
                        deleted++;
                }

                if (deleted > 0)
                {
                    MessageBox.Show(this, $"✅ {deleted} پرسنل با موفقیت حذف شد.", "موفقیت",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                }
                else
                {
                    MessageBox.Show(this, "هیچ پرسنلی حذف نشد.", "هشدار",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPersonnelManager.DeleteSelected", ex);
                MessageBox.Show(this, $"خطا در حذف پرسنل:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        private void ExportExcel()
        {
            if (dgvPersonnel.DataSource is not DataTable dt || dt.Rows.Count == 0)
            {
                MessageBox.Show(this, "داده‌ای برای خروجی وجود ندارد.", "اطلاع",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var dialog = new SaveFileDialog
            {
                Filter = "Excel Workbook|*.xlsx",
                Title = "ذخیره خروجی اکسل پرسنل",
                FileName = $"پرسنل_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                AddExtension = true
            };
            if (dialog.ShowDialog(this) != DialogResult.OK)
                return;

            try
            {
                UseWaitCursor = true;
                WriteWorkbook(dt, dialog.FileName);
                UseWaitCursor = false;
                MessageBox.Show(this, "✅ فایل اکسل با موفقیت ذخیره شد.", "خروجی اکسل",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                UseWaitCursor = false;
                ErrorLogger.Log("FormPersonnelManager.ExportExcel", ex);
                MessageBox.Show(this, $"خطا در ایجاد فایل اکسل:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void WriteWorkbook(DataTable data, string path)
        {
            using var workbook = new XLWorkbook();
            IXLWorksheet sheet = workbook.Worksheets.Add("پرسنل");
            sheet.RightToLeft = true;

            for (int c = 0; c < data.Columns.Count; c++)
                sheet.Cell(1, c + 1).Value = data.Columns[c].ColumnName;

            IXLRange headerRange = sheet.Range(1, 1, 1, data.Columns.Count);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(30, 64, 175);
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Row(1).Height = 28;

            for (int r = 0; r < data.Rows.Count; r++)
            {
                for (int c = 0; c < data.Columns.Count; c++)
                {
                    object? value = data.Rows[r][c];
                    IXLCell cell = sheet.Cell(r + 2, c + 1);
                    if (value == null || value == DBNull.Value)
                        cell.Value = "";
                    else if (value is double or float or decimal or int or long)
                        cell.Value = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
                    else
                        cell.Value = value.ToString();
                }
            }

            sheet.Range(1, 1, Math.Max(1, data.Rows.Count + 1), data.Columns.Count).SetAutoFilter();
            sheet.SheetView.FreezeRows(1);
            sheet.Columns().AdjustToContents(10, 40);
            workbook.SaveAs(path);
        }
    }
}
