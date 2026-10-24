using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    /// <summary>
    /// فرم مدرن افزودن / ویرایش پرسنل.
    /// </summary>
    public sealed class FormPersonnelEditor : Form
    {
        private readonly int? _personnelId;
        private readonly DbHelper db = new();

        private static readonly Color PrimaryColor = Color.FromArgb(0, 122, 204);
        private static readonly Color CanvasColor = Color.FromArgb(241, 245, 249);
        private static readonly Color TextColor = Color.FromArgb(30, 41, 59);
        private static readonly Color SuccessColor = Color.FromArgb(16, 185, 129);
        private static readonly Color DangerColor = Color.FromArgb(239, 68, 68);

        private ComboBox cmbProvince = null!;
        private ComboBox cmbCity = null!;
        private ComboBox cmbAffair = null!;
        private ComboBox cmbDept = null!;
        private ComboBox cmbDistrict = null!;
        private ComboBox cmbPostName = null!;
        private ComboBox cmbVoltage = null!;
        private ComboBox cmbWorkShift = null!;
        private ComboBox cmbGender = null!;
        private ComboBox cmbContractType = null!;
        private ComboBox cmbJobLevel = null!;
        private ComboBox cmbCompany = null!;
        private ComboBox cmbDegree = null!;
        private ComboBox cmbDegreeField = null!;
        private ComboBox cmbMainJobTitle = null!;
        private ComboBox cmbCurrentActivity = null!;
        private ComboBox cmbStatus = null!;

        private TextBox txtFirstName = null!;
        private TextBox txtLastName = null!;
        private TextBox txtFatherName = null!;
        private TextBox txtPersonnelNumber = null!;
        private TextBox txtNationalID = null!;
        private TextBox txtMobileNumber = null!;
        private TextBox txtDescription = null!;

        private DateTimePicker dtpBirthDate = null!;
        private DateTimePicker dtpHireDate = null!;
        private DateTimePicker dtpStartDateOperation = null!;

        private CheckBox chkInconsistency = null!;
        private PictureBox pbPhoto = null!;
        private string selectedPhotoPath = string.Empty;

        private bool _loading;

        public FormPersonnelEditor(int? personnelId = null)
        {
            _personnelId = personnelId;
            BuildUi();
            FontSettings.ApplyFontToForm(this);
            LoadLookups();
            if (_personnelId.HasValue)
                LoadExisting();
        }

        private void BuildUi()
        {
            Text = _personnelId.HasValue ? "ویرایش پرسنل" : "ثبت پرسنل جدید";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(1000, 780);
            BackColor = CanvasColor;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            Font = new Font(FontSettings.FontFamilyName, 9F);
            AutoScroll = true;

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
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 62F));
            Controls.Add(root);

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildBody(), 0, 1);
            root.Controls.Add(BuildFooter(), 0, 2);
        }

        private Control BuildHeader()
        {
            var header = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(15, 23, 42),
                Padding = new Padding(24, 12, 24, 10),
                Margin = new Padding(0)
            };

            var title = new Label
            {
                Text = _personnelId.HasValue ? "✏️ ویرایش پرسنل" : "➕ ثبت پرسنل جدید",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font(FontSettings.FontFamilyName, 16F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var subtitle = new Label
            {
                Text = "اطلاعات شخصی، موقعیت سازمانی، سوابق و عکس پرسنل",
                Dock = DockStyle.Fill,
                Font = new Font(FontSettings.FontFamilyName, 9.5F),
                ForeColor = Color.FromArgb(203, 213, 225),
                TextAlign = ContentAlignment.MiddleLeft
            };

            var accent = new Panel
            {
                Dock = DockStyle.Right,
                Width = 5,
                BackColor = Color.FromArgb(56, 189, 248),
                Margin = new Padding(0, 6, 0, 6)
            };

            header.Controls.Add(subtitle);
            header.Controls.Add(title);
            header.Controls.Add(accent);
            return header;
        }

        private Control BuildBody()
        {
            var scroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = CanvasColor,
                Margin = new Padding(0)
            };

            var grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 4,
                RowCount = 1,
                BackColor = Color.White,
                Padding = new Padding(20, 14, 20, 14),
                Margin = new Padding(0)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 270F));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 270F));

            int row = 0;
            int slot = 0;

            // عکس
            var photoPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 150,
                BackColor = Color.FromArgb(248, 250, 252),
                Margin = new Padding(0, 0, 0, 8)
            };
            pbPhoto = new PictureBox
            {
                Size = new Size(110, 110),
                Location = new Point(photoPanel.Width / 2 - 55, 12),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White
            };
            pbPhoto.Image = ImageHelper.CreateDefaultImage(110, 110);

            var btnSelect = CreateSmallButton("انتخاب عکس", PrimaryColor);
            btnSelect.Location = new Point(photoPanel.Width / 2 - 120, 128);
            btnSelect.Click += (_, _) => SelectPhoto();

            var btnRemove = CreateSmallButton("حذف عکس", DangerColor);
            btnRemove.Location = new Point(photoPanel.Width / 2 + 14, 128);
            btnRemove.Click += (_, _) =>
            {
                selectedPhotoPath = string.Empty;
                pbPhoto.Image = ImageHelper.CreateDefaultImage(110, 110);
            };

            photoPanel.Controls.Add(btnRemove);
            photoPanel.Controls.Add(btnSelect);
            photoPanel.Controls.Add(pbPhoto);
            grid.Controls.Add(photoPanel, 0, row);
            grid.SetColumnSpan(photoPanel, 4);
            row++;

            AddGroupHeader(grid, ref row, ref slot, "👤 اطلاعات شخصی");
            row = AddPair(grid, row, ref slot, "نام", CreateText(out txtFirstName, 60));
            row = AddPair(grid, row, ref slot, "نام خانوادگی", CreateText(out txtLastName, 60));
            row = AddPair(grid, row, ref slot, "نام پدر", CreateText(out txtFatherName, 60));
            row = AddPair(grid, row, ref slot, "شماره پرسنلی", CreateText(out txtPersonnelNumber, 30));
            row = AddPair(grid, row, ref slot, "کد ملی", CreateText(out txtNationalID, 11));
            row = AddPair(grid, row, ref slot, "شماره موبایل", CreateText(out txtMobileNumber, 11));
            row = AddPair(grid, row, ref slot, "جنسیت", CreateCombo(out cmbGender));
            row = AddPair(grid, row, ref slot, "شیفت کاری", CreateCombo(out cmbWorkShift));

            AddGroupHeader(grid, ref row, ref slot, "📍 موقعیت سازمانی");
            row = AddPair(grid, row, ref slot, "استان", CreateCombo(out cmbProvince));
            row = AddPair(grid, row, ref slot, "شهر", CreateCombo(out cmbCity));
            row = AddPair(grid, row, ref slot, "امور انتقال", CreateCombo(out cmbAffair));
            row = AddPair(grid, row, ref slot, "اداره بهره‌برداری", CreateCombo(out cmbDept));
            row = AddPair(grid, row, ref slot, "ناحیه", CreateCombo(out cmbDistrict));
            row = AddPair(grid, row, ref slot, "نام پست", CreateCombo(out cmbPostName));
            row = AddPair(grid, row, ref slot, "سطح ولتاژ", CreateCombo(out cmbVoltage));

            AddGroupHeader(grid, ref row, ref slot, "🏢 مشخصات شغلی و استخدامی");
            row = AddPair(grid, row, ref slot, "نوع قرارداد", CreateCombo(out cmbContractType));
            row = AddPair(grid, row, ref slot, "سطح شغل", CreateCombo(out cmbJobLevel));
            row = AddPair(grid, row, ref slot, "شرکت", CreateCombo(out cmbCompany));
            row = AddPair(grid, row, ref slot, "مدرک تحصیلی", CreateCombo(out cmbDegree));
            row = AddPair(grid, row, ref slot, "رشته تحصیلی", CreateCombo(out cmbDegreeField));
            row = AddPair(grid, row, ref slot, "عنوان شغلی اصلی", CreateCombo(out cmbMainJobTitle));
            row = AddPair(grid, row, ref slot, "فعالیت فعلی", CreateCombo(out cmbCurrentActivity));
            row = AddPair(grid, row, ref slot, "وضعیت حضور", CreateCombo(out cmbStatus));

            chkInconsistency = new CheckBox
            {
                Dock = DockStyle.Fill,
                Text = "ندارد",
                Checked = false,
                Enabled = false,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                Margin = new Padding(0, 4, 12, 4)
            };
            row = AddPair(grid, row, ref slot, "مغایرت", chkInconsistency);

            AddGroupHeader(grid, ref row, ref slot, "📅 تاریخ‌ها");
            row = AddPair(grid, row, ref slot, "تاریخ تولد", CreateDate(out dtpBirthDate, -25));
            row = AddPair(grid, row, ref slot, "تاریخ استخدام", CreateDate(out dtpHireDate, 0));
            row = AddPair(grid, row, ref slot, "تاریخ شروع بکار", CreateDate(out dtpStartDateOperation, 0));

            AddGroupHeader(grid, ref row, ref slot, "📝 توضیحات");
            var txtDescriptionBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                Height = 60,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                Margin = new Padding(0, 4, 12, 4)
            };
            txtDescription = txtDescriptionBox;
            grid.Controls.Add(CreateCaption("توضیحات"), 0, row);
            grid.Controls.Add(txtDescriptionBox, 1, row);
            grid.SetColumnSpan(txtDescriptionBox, 3);
            row++;
            slot = 0;

            grid.RowCount = row;
            scroll.Controls.Add(grid);
            return scroll;
        }

        private static Label CreateCaption(string text)
        {
            return new Label
            {
                Text = text + ":",
                Dock = DockStyle.Fill,
                Font = new Font(FontSettings.FontFamilyName, 9F, FontStyle.Bold),
                ForeColor = TextColor,
                TextAlign = ContentAlignment.MiddleRight,
                Padding = new Padding(6, 0, 0, 0)
            };
        }

        private static ComboBox CreateCombo(out ComboBox combo)
        {
            combo = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                Margin = new Padding(0, 4, 12, 4)
            };
            return combo;
        }

        private static TextBox CreateText(out TextBox box, int maxLength)
        {
            box = new TextBox
            {
                Dock = DockStyle.Fill,
                MaxLength = maxLength,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                Margin = new Padding(0, 4, 12, 4)
            };
            return box;
        }

        private static DateTimePicker CreateDate(out DateTimePicker picker, int yearsOffset)
        {
            picker = new DateTimePicker
            {
                Dock = DockStyle.Fill,
                Format = DateTimePickerFormat.Short,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                Margin = new Padding(0, 4, 12, 4),
                Value = DateTime.Now.AddYears(yearsOffset)
            };
            return picker;
        }

        private static Button CreateSmallButton(string text, Color backColor)
        {
            var button = new Button
            {
                Text = text,
                Size = new Size(106, 30),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(FontSettings.FontFamilyName, 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private static void AddGroupHeader(TableLayoutPanel grid, ref int row, ref int slot, string text)
        {
            if (slot == 1)
            {
                row++;
                slot = 0;
            }

            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 28,
                BackColor = Color.FromArgb(248, 250, 252),
                Margin = new Padding(0, 8, 0, 4)
            };
            header.Controls.Add(new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = new Font(FontSettings.FontFamilyName, 10.5F, FontStyle.Bold),
                ForeColor = PrimaryColor,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            });
            header.Controls.Add(new Panel
            {
                Dock = DockStyle.Right,
                Width = 4,
                BackColor = PrimaryColor,
                Margin = new Padding(0)
            });
            grid.Controls.Add(header, 0, row);
            grid.SetColumnSpan(header, 4);
            row++;
        }

        private int AddPair(TableLayoutPanel grid, int row, ref int slot, string caption, Control control)
        {
            int captionColumn = slot == 0 ? 0 : 2;
            int controlColumn = slot == 0 ? 1 : 3;
            grid.Controls.Add(CreateCaption(caption), captionColumn, row);
            grid.Controls.Add(control, controlColumn, row);
            slot++;
            if (slot == 2)
            {
                slot = 0;
                row++;
            }
            return row;
        }

        private Control BuildFooter()
        {
            var footer = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(16, 10, 16, 10),
                Margin = new Padding(0)
            };

            var btnSave = CreateActionButton("💾 ذخیره", SuccessColor);
            btnSave.Dock = DockStyle.Right;
            btnSave.Width = 160;
            btnSave.Click += (_, _) => Save();

            var btnCancel = CreateActionButton("❌ لغو", DangerColor);
            btnCancel.Dock = DockStyle.Right;
            btnCancel.Width = 120;
            btnCancel.Margin = new Padding(0, 0, 12, 0);
            btnCancel.Click += (_, _) => Close();

            footer.Controls.Add(btnSave);
            footer.Controls.Add(btnCancel);
            return footer;
        }

        private static Button CreateActionButton(string text, Color backColor)
        {
            var button = new Button
            {
                Text = text,
                Height = 40,
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(FontSettings.FontFamilyName, 10F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        // ─────────────────────────────────────────────
        // داده‌ها
        // ─────────────────────────────────────────────
        private void LoadLookups()
        {
            _loading = true;
            try
            {
                BindCombo(cmbProvince, db.GetProvinces(), "ProvinceName", "ProvinceID");
                BindCombo(cmbGender, db.GetGenders(), "GenderName", "GenderID");
                BindCombo(cmbVoltage, db.GetVoltageLevels(), "VoltageName", "VoltageID");
                BindCombo(cmbWorkShift, db.GetWorkShifts(), "WorkShiftName", "WorkShiftID");
                BindCombo(cmbContractType, db.GetContractTypes(), "ContractTypeName", "ContractTypeID");
                BindCombo(cmbJobLevel, db.GetJobLevels(), "JobLevelName", "JobLevelID");
                BindCombo(cmbCompany, db.GetCompanies(), "CompanyName", "CompanyID");
                BindCombo(cmbDegree, db.GetDegrees(), "DegreeName", "DegreeID");
                BindCombo(cmbDegreeField, db.GetDegreeFields(), "DegreeFieldName", "DegreeFieldID");
                BindCombo(cmbMainJobTitle, db.GetChartAffairs(), "ChartName", "ChartID");
                BindCombo(cmbCurrentActivity, db.GetChartAffairs1(), "ChartName", "ChartID");
                BindCombo(cmbStatus, db.GetStatusPresence(), "StatusName", "StatusID");
            }
            finally
            {
                _loading = false;
            }

            cmbProvince.SelectedIndexChanged += CmbProvince_Changed;
            cmbAffair.SelectedIndexChanged += CmbAffair_Changed;
            cmbDept.SelectedIndexChanged += CmbDept_Changed;
            cmbDistrict.SelectedIndexChanged += CmbDistrict_Changed;
            cmbMainJobTitle.SelectedIndexChanged += (_, _) => UpdateInconsistency();
            cmbCurrentActivity.SelectedIndexChanged += (_, _) => UpdateInconsistency();

            // اگر استان به‌صورت پیش‌فرض انتخاب شده، شهرها و امور متناظر را بارگذاری کن
            if (SelectedId(cmbProvince) is int pid && pid > 0)
                CmbProvince_Changed(null, EventArgs.Empty);
        }

        private static void BindCombo(ComboBox combo, DataTable? table, string display, string value)
        {
            if (table != null && table.Rows.Count > 0)
            {
                combo.DataSource = table;
                combo.DisplayMember = display;
                combo.ValueMember = value;
            }
        }

        private void CmbProvince_Changed(object? sender, EventArgs e)
        {
            if (_loading)
                return;

            if (SelectedId(cmbProvince) is int provinceId && provinceId > 0)
            {
                BindCombo(cmbCity, db.GetCitiesByProvince(provinceId), "CityName", "CityID");
                BindCombo(cmbAffair, db.GetAffairsByProvince(provinceId), "AffairName", "AffairID");
            }
            else
            {
                cmbCity.DataSource = null;
                cmbAffair.DataSource = null;
            }
            cmbDept.DataSource = null;
            cmbDistrict.DataSource = null;
            cmbPostName.DataSource = null;
        }

        private void CmbAffair_Changed(object? sender, EventArgs e)
        {
            if (_loading)
                return;

            if (SelectedId(cmbAffair) is int affairId && affairId > 0)
                BindCombo(cmbDept, db.GetDeptsByAffair(affairId), "DeptName", "DeptID");
            else
                cmbDept.DataSource = null;

            cmbDistrict.DataSource = null;
            cmbPostName.DataSource = null;
        }

        private void CmbDept_Changed(object? sender, EventArgs e)
        {
            if (_loading)
                return;

            if (SelectedId(cmbDept) is int deptId && deptId > 0)
                BindCombo(cmbDistrict, db.GetDistrictsByDept(deptId), "DistrictName", "DistrictID");
            else
                cmbDistrict.DataSource = null;

            cmbPostName.DataSource = null;
        }

        private void CmbDistrict_Changed(object? sender, EventArgs e)
        {
            if (_loading)
                return;

            if (SelectedId(cmbDistrict) is int districtId && districtId > 0)
                BindCombo(cmbPostName, db.GetPostNamesByDistrict(districtId), "PostName", "PostNameID");
            else
                cmbPostName.DataSource = null;
        }

        private static int? SelectedId(ComboBox combo)
        {
            if (combo.SelectedIndex == -1 || combo.SelectedValue == null)
                return null;

            return int.TryParse(combo.SelectedValue.ToString(), out int value) ? value : null;
        }

        private void UpdateInconsistency()
        {
            if (cmbMainJobTitle.SelectedIndex != -1 && cmbCurrentActivity.SelectedIndex != -1)
            {
                bool mismatch = !string.Equals(cmbMainJobTitle.Text?.Trim(), cmbCurrentActivity.Text?.Trim(), StringComparison.OrdinalIgnoreCase);
                chkInconsistency.Checked = mismatch;
                chkInconsistency.Text = mismatch ? "دارد" : "ندارد";
            }
        }

        private void SelectPhoto()
        {
            string? path = ImageHelper.OpenImageDialog();
            if (string.IsNullOrWhiteSpace(path))
                return;

            try
            {
                using var image = Image.FromFile(path);
                pbPhoto.Image = (Image)image.Clone();
                selectedPhotoPath = path;
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPersonnelEditor.SelectPhoto", ex);
                MessageBox.Show(this, $"خطا در بارگذاری عکس:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadExisting()
        {
            DataTable? rows = db.ExecuteQuery("SELECT * FROM Personnel WHERE PersonnelID = ?",
                new[] { new OleDbParameter("?", _personnelId!.Value) }, showError: false);
            if (rows == null || rows.Rows.Count == 0)
                return;

            DataRow r = rows.Rows[0];
            _loading = true;
            try
            {
                SelectComboValue(cmbProvince, r["ProvinceID"]);

                if (SelectedId(cmbProvince) is int provinceId && provinceId > 0)
                {
                    BindCombo(cmbCity, db.GetCitiesByProvince(provinceId), "CityName", "CityID");
                    BindCombo(cmbAffair, db.GetAffairsByProvince(provinceId), "AffairName", "AffairID");
                }
                SelectComboValue(cmbCity, r["CityID"]);
                SelectComboValue(cmbAffair, r["AffairID"]);

                if (SelectedId(cmbAffair) is int affairId && affairId > 0)
                    BindCombo(cmbDept, db.GetDeptsByAffair(affairId), "DeptName", "DeptID");
                SelectComboValue(cmbDept, r["DeptID"]);

                if (SelectedId(cmbDept) is int deptId && deptId > 0)
                    BindCombo(cmbDistrict, db.GetDistrictsByDept(deptId), "DistrictName", "DistrictID");
                SelectComboValue(cmbDistrict, r["DistrictID"]);

                if (SelectedId(cmbDistrict) is int districtId && districtId > 0)
                    BindCombo(cmbPostName, db.GetPostNamesByDistrict(districtId), "PostName", "PostNameID");
                SelectComboValue(cmbPostName, r["PostNameID"]);

                SelectComboValue(cmbVoltage, r["VoltageID"]);
                SelectComboValue(cmbWorkShift, r["WorkShiftID"]);
                SelectComboValue(cmbGender, r["GenderID"]);
                SelectComboValue(cmbContractType, r["ContractTypeID"]);
                SelectComboValue(cmbJobLevel, r["JobLevelID"]);
                SelectComboValue(cmbCompany, r["CompanyID"]);
                SelectComboValue(cmbDegree, r["DegreeID"]);
                SelectComboValue(cmbDegreeField, r["DegreeFieldID"]);
                SelectComboValue(cmbMainJobTitle, r["MainJobTitle"]);
                SelectComboValue(cmbCurrentActivity, r["CurrentActivity"]);
                SelectComboValue(cmbStatus, r["StatusID"]);

                txtFirstName.Text = r["FirstName"]?.ToString() ?? "";
                txtLastName.Text = r["LastName"]?.ToString() ?? "";
                txtFatherName.Text = r["FatherName"]?.ToString() ?? "";
                txtPersonnelNumber.Text = r["PersonnelNumber"]?.ToString() ?? "";
                txtNationalID.Text = r["NationalID"]?.ToString() ?? "";
                txtMobileNumber.Text = r["MobileNumber"]?.ToString() ?? "";
                txtDescription.Text = r["Description"]?.ToString() ?? "";

                SetPickerDate(dtpBirthDate, ParseDate(r["BirthDate"]), DateTime.Now.AddYears(-25));
                SetPickerDate(dtpHireDate, ParseDate(r["HireDate"]), DateTime.Now);
                SetPickerDate(dtpStartDateOperation, ParseDate(r["StartDateOperation"]), DateTime.Now);

                chkInconsistency.Checked = Convert.ToInt32(r["Inconsistency"]) == 1;
                chkInconsistency.Text = chkInconsistency.Checked ? "دارد" : "ندارد";

                try
                {
                    if (ImageHelper.ImageExists(txtNationalID.Text.Trim()))
                        ImageHelper.DrawImageInPictureBox(pbPhoto, ImageHelper.LoadImage(txtNationalID.Text.Trim()));
                }
                catch (Exception ex)
                {
                    ErrorLogger.Log("FormPersonnelEditor.LoadExistingPhoto", ex);
                }
            }
            finally
            {
                _loading = false;
            }
        }

        private static void SelectComboValue(ComboBox combo, object? value)
        {
            if (value == null || value == DBNull.Value)
                return;

            if (!int.TryParse(value.ToString(), out int id) || id <= 0)
                return;

            if (combo.DataSource is DataTable table)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    if (Convert.ToInt32(table.Rows[i][combo.ValueMember]) == id)
                    {
                        combo.SelectedIndex = i;
                        return;
                    }
                }
            }
        }

        private static DateTime? ParseDate(object? value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            string s = (value.ToString() ?? "").Trim();
            if (s.Length == 0)
                return null;

            // فرمت‌های میلادی
            foreach (string format in new[] { "yyyy-MM-dd", "yyyy/MM/dd", "MM/dd/yyyy", "M/d/yyyy" })
            {
                if (DateTime.TryParseExact(s, format, System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out DateTime exact))
                    return exact;
            }

            // تاریخ شمسی (سال ۱۳۰۰ به بالا)
            var parts = s.Split(new[] { '/', '-', '.' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 3 &&
                int.TryParse(parts[0], out int py) && py >= 1300 && py <= 1500 &&
                int.TryParse(parts[1], out int pm) && pm >= 1 && pm <= 12 &&
                int.TryParse(parts[2], out int pd) && pd >= 1 && pd <= 31)
            {
                try
                {
                    return new System.Globalization.PersianCalendar().ToDateTime(py, pm, pd, 0, 0, 0, 0);
                }
                catch
                {
                    return null;
                }
            }

            return DateTime.TryParse(s, out DateTime parsed) ? parsed : null;
        }

        private static void SetPickerDate(DateTimePicker picker, DateTime? value, DateTime fallback)
        {
            DateTime target = value ?? fallback;
            if (target < picker.MinDate || target > picker.MaxDate)
                target = fallback;
            picker.Value = target;
        }

        // ─────────────────────────────────────────────
        // اعتبارسنجی و ذخیره
        // ─────────────────────────────────────────────
        private bool IsComboValid(ComboBox combo)
            => combo.SelectedIndex != -1 && SelectedId(combo) is int id && id > 0;

        private string? ValidateForm()
        {
            if (!IsComboValid(cmbProvince)) return "لطفاً استان را انتخاب کنید.";
            if (!IsComboValid(cmbCity)) return "لطفاً شهر را انتخاب کنید.";
            if (!IsComboValid(cmbAffair)) return "لطفاً امور انتقال را انتخاب کنید.";
            if (!IsComboValid(cmbDept)) return "لطفاً اداره بهره‌برداری را انتخاب کنید.";
            if (!IsComboValid(cmbDistrict)) return "لطفاً ناحیه بهره‌برداری را انتخاب کنید.";
            if (!IsComboValid(cmbPostName)) return "لطفاً نام پست را انتخاب کنید.";
            if (!IsComboValid(cmbVoltage)) return "لطفاً سطح ولتاژ را انتخاب کنید.";
            if (!IsComboValid(cmbWorkShift)) return "لطفاً شیفت کاری را انتخاب کنید.";
            if (!IsComboValid(cmbGender)) return "لطفاً جنسیت را انتخاب کنید.";
            if (!IsComboValid(cmbContractType)) return "لطفاً نوع قرارداد را انتخاب کنید.";
            if (!IsComboValid(cmbJobLevel)) return "لطفاً سطح شغل را انتخاب کنید.";
            if (!IsComboValid(cmbCompany)) return "لطفاً شرکت را انتخاب کنید.";
            if (!IsComboValid(cmbDegree)) return "لطفاً مدرک تحصیلی را انتخاب کنید.";
            if (!IsComboValid(cmbDegreeField)) return "لطفاً رشته تحصیلی را انتخاب کنید.";
            if (!IsComboValid(cmbMainJobTitle)) return "لطفاً عنوان شغلی اصلی را انتخاب کنید.";
            if (!IsComboValid(cmbCurrentActivity)) return "لطفاً فعالیت فعلی را انتخاب کنید.";
            if (!IsComboValid(cmbStatus)) return "لطفاً وضعیت حضور را انتخاب کنید.";

            if (string.IsNullOrWhiteSpace(txtFirstName.Text)) return "لطفاً نام را وارد کنید.";
            if (string.IsNullOrWhiteSpace(txtLastName.Text)) return "لطفاً نام خانوادگی را وارد کنید.";
            if (string.IsNullOrWhiteSpace(txtFatherName.Text)) return "لطفاً نام پدر را وارد کنید.";
            if (string.IsNullOrWhiteSpace(txtPersonnelNumber.Text)) return "لطفاً شماره پرسنلی را وارد کنید.";
            if (string.IsNullOrWhiteSpace(txtNationalID.Text)) return "لطفاً کد ملی را وارد کنید.";
            if (string.IsNullOrWhiteSpace(txtMobileNumber.Text)) return "لطفاً شماره موبایل را وارد کنید.";

            if (!Regex.IsMatch(txtNationalID.Text.Trim(), @"^\d{10,11}$"))
                return "کد ملی باید ۱۰ یا ۱۱ رقم باشد.";
            if (!Regex.IsMatch(txtMobileNumber.Text.Trim(), @"^09\d{9}$"))
                return "شماره موبایل باید ۱۱ رقم و شروع با 09 باشد.";
            if (!Regex.IsMatch(txtPersonnelNumber.Text.Trim(), @"^\d+$"))
                return "شماره پرسنلی باید فقط شامل اعداد باشد.";

            if (dtpBirthDate.Value >= DateTime.Now)
                return "تاریخ تولد نمی‌تواند در آینده باشد.";
            if (dtpBirthDate.Value.AddYears(18) > DateTime.Now)
                return "پرسنل باید حداقل ۱۸ سال سن داشته باشد.";
            if (dtpHireDate.Value > DateTime.Now)
                return "تاریخ استخدام نمی‌تواند در آینده باشد.";
            if (dtpStartDateOperation.Value > DateTime.Now)
                return "تاریخ شروع بکار نمی‌تواند در آینده باشد.";
            if (dtpHireDate.Value > dtpStartDateOperation.Value)
                return "تاریخ استخدام نمی‌تواند بعد از تاریخ شروع بکار باشد.";

            return null;
        }

        private object ComboValue(ComboBox combo)
        {
            int? id = SelectedId(combo);
            return id.HasValue ? id.Value : (object)DBNull.Value;
        }

        private void Save()
        {
            string? error = ValidateForm();
            if (error != null)
            {
                MessageBox.Show(this, error, "خطای اعتبارسنجی", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                bool isNew = !_personnelId.HasValue;
                string query = isNew
                    ? @"INSERT INTO Personnel
                        (ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, VoltageID, WorkShiftID, GenderID,
                         FirstName, LastName, FatherName, PersonnelNumber, NationalID, MobileNumber,
                         BirthDate, HireDate, StartDateOperation, ContractTypeID, JobLevelID, CompanyID,
                         DegreeID, DegreeFieldID, MainJobTitle, CurrentActivity, Inconsistency, Description, StatusID)
                       VALUES
                        (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                    : @"UPDATE Personnel SET
                        ProvinceID = ?, CityID = ?, AffairID = ?, DeptID = ?, DistrictID = ?, PostNameID = ?, VoltageID = ?, WorkShiftID = ?, GenderID = ?,
                        FirstName = ?, LastName = ?, FatherName = ?, PersonnelNumber = ?, NationalID = ?, MobileNumber = ?,
                        BirthDate = ?, HireDate = ?, StartDateOperation = ?, ContractTypeID = ?, JobLevelID = ?, CompanyID = ?,
                        DegreeID = ?, DegreeFieldID = ?, MainJobTitle = ?, CurrentActivity = ?, Inconsistency = ?, Description = ?, StatusID = ?
                       WHERE PersonnelID = ?";

                var parameters = new System.Collections.Generic.List<OleDbParameter>
                {
                    new("?", ComboValue(cmbProvince)),
                    new("?", ComboValue(cmbCity)),
                    new("?", ComboValue(cmbAffair)),
                    new("?", ComboValue(cmbDept)),
                    new("?", ComboValue(cmbDistrict)),
                    new("?", ComboValue(cmbPostName)),
                    new("?", ComboValue(cmbVoltage)),
                    new("?", ComboValue(cmbWorkShift)),
                    new("?", ComboValue(cmbGender)),
                    new("?", txtFirstName.Text.Trim()),
                    new("?", txtLastName.Text.Trim()),
                    new("?", txtFatherName.Text.Trim()),
                    new("?", txtPersonnelNumber.Text.Trim()),
                    new("?", txtNationalID.Text.Trim()),
                    new("?", txtMobileNumber.Text.Trim()),
                    new("?", dtpBirthDate.Value.ToString("yyyy-MM-dd")),
                    new("?", dtpHireDate.Value.ToString("yyyy-MM-dd")),
                    new("?", dtpStartDateOperation.Value.ToString("yyyy-MM-dd")),
                    new("?", ComboValue(cmbContractType)),
                    new("?", ComboValue(cmbJobLevel)),
                    new("?", ComboValue(cmbCompany)),
                    new("?", ComboValue(cmbDegree)),
                    new("?", ComboValue(cmbDegreeField)),
                    new("?", ComboValue(cmbMainJobTitle)),
                    new("?", ComboValue(cmbCurrentActivity)),
                    new("?", chkInconsistency.Checked ? 1 : 0),
                    new("?", string.IsNullOrWhiteSpace(txtDescription.Text) ? DBNull.Value : txtDescription.Text.Trim()),
                    new("?", ComboValue(cmbStatus))
                };

                if (!isNew)
                    parameters.Add(new OleDbParameter("?", _personnelId!.Value));

                int affected = db.ExecuteNonQuery(query, parameters.ToArray(), showError: false);
                if (affected <= 0)
                {
                    MessageBox.Show(this, "ذخیره‌سازی انجام نشد. لطفاً اتصال پایگاه داده را بررسی کنید.",
                        "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!string.IsNullOrEmpty(selectedPhotoPath))
                {
                    bool photoSaved = ImageHelper.SaveImage(selectedPhotoPath, txtNationalID.Text.Trim());
                    if (!photoSaved)
                        MessageBox.Show(this, "پرسنل ذخیره شد اما عکس ذخیره نشد.", "اخطار",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPersonnelEditor.Save", ex);
                MessageBox.Show(this, $"خطا در ذخیره پرسنل:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
