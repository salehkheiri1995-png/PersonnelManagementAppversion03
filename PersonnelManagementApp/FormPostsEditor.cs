using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    /// <summary>
    /// فرم افزودن / ویرایش یک پست برق (Post).
    /// </summary>
    public sealed class FormPostsEditor : Form
    {
        private readonly int? _postId;
        private readonly PostsCrudService _crud;

        private static readonly Color PrimaryColor = Color.FromArgb(0, 122, 204);
        private static readonly Color CanvasColor = Color.FromArgb(241, 245, 249);
        private static readonly Color TextColor = Color.FromArgb(30, 41, 59);
        private static readonly Color MutedColor = Color.FromArgb(100, 116, 139);
        private static readonly Color BorderColor = Color.FromArgb(226, 232, 240);

        // کامبو‌های وابسته
        private ComboBox cmbProvince = null!;
        private ComboBox cmbCity = null!;
        private ComboBox cmbAffair = null!;
        private ComboBox cmbDept = null!;
        private ComboBox cmbDistrict = null!;
        private ComboBox cmbPostName = null!;

        private readonly Dictionary<string, NumericUpDown> _numericBoxes = new();
        private readonly Dictionary<string, TextBox> _textBoxes = new();
        private readonly Dictionary<string, ComboBox> _lookupCombos = new();

        // داده‌های وابسته در حافظه (برای فیلتر سریع بدون کوئری تکراری)
        private List<LookupItem> _cities = new();
        private List<LookupItem> _affairs = new();
        private List<LookupItem> _depts = new();
        private List<LookupItem> _districts = new();
        private List<LookupItem> _postNames = new();

        private bool _loading;

        private readonly struct LookupItem
        {
            public LookupItem(int id, string name, int parentId)
            {
                Id = id;
                Name = name;
                ParentId = parentId;
            }

            public int Id { get; }
            public string Name { get; }
            public int ParentId { get; }
        }

        public FormPostsEditor(int? postId = null)
        {
            _postId = postId;
            _crud = new PostsCrudService();

            BuildUi();
            FontSettings.ApplyFontToForm(this);
            LoadLookups();
            if (_postId.HasValue)
                LoadExisting();
        }

        private void BuildUi()
        {
            Text = _postId.HasValue ? "ویرایش پست" : "ثبت پست جدید";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(980, 760);
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
                Text = _postId.HasValue ? "✏️ ویرایش پست برق" : "➕ ثبت پست برق جدید",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font(FontSettings.FontFamilyName, 16F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var subtitle = new Label
            {
                Text = "اطلاعات موقعیت، مشخصات فنی، ظرفیت و آدرس پست را وارد کنید",
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
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 245F));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 245F));

            int row = 0;
            int pairSlot = 0;

            AddGroupHeader(grid, ref row, ref pairSlot, "📍 موقعیت و سازمان");
            row = AddComboPair(grid, row, ref pairSlot, "استان", out cmbProvince, "Province");
            row = AddComboPair(grid, row, ref pairSlot, "شهر", out cmbCity, "City");
            row = AddComboPair(grid, row, ref pairSlot, "امور انتقال", out cmbAffair, "Affair");
            row = AddComboPair(grid, row, ref pairSlot, "اداره بهره‌برداری", out cmbDept, "Dept");
            row = AddComboPair(grid, row, ref pairSlot, "ناحیه", out cmbDistrict, "District");
            row = AddComboPair(grid, row, ref pairSlot, "نام پست", out cmbPostName, "PostName");

            AddGroupHeader(grid, ref row, ref pairSlot, "⚙️ مشخصات فنی");
            row = AddComboPair(grid, row, ref pairSlot, "سطح ولتاژ", "Voltage");
            row = AddComboPair(grid, row, ref pairSlot, "استاندارد پست", "Standard");
            row = AddComboPair(grid, row, ref pairSlot, "نوع پست", "Type");
            row = AddComboPair(grid, row, ref pairSlot, "اتصال توزیع", "Conn");
            row = AddComboPair(grid, row, ref pairSlot, "نوع عایق", "Ins");
            row = AddComboPair(grid, row, ref pairSlot, "نوع پست ۲", "Pt2");
            row = AddComboPair(grid, row, ref pairSlot, "ثابت / سیار", "Fm");
            row = AddComboPair(grid, row, ref pairSlot, "وضعیت مدار", "Circuit");
            row = AddComboPair(grid, row, ref pairSlot, "دیزل ژنراتور", "Diesel");
            row = AddComboPair(grid, row, ref pairSlot, "نوع توزیع", "Feed");
            row = AddComboPair(grid, row, ref pairSlot, "وضعیت آب", "Water");
            row = AddComboPair(grid, row, ref pairSlot, "مهمان‌سرا", "Guest");

            AddGroupHeader(grid, ref row, ref pairSlot, "📊 ظرفیت، ابعاد و بهره‌برداری");
            row = AddNumericPair(grid, row, ref pairSlot, "سال بهره‌برداری", "OperationYear", 0, 2100, 0);
            row = AddNumericPair(grid, row, ref pairSlot, "ظرفیت HV", "CapacityHv", 0, 1000000, 1);
            row = AddNumericPair(grid, row, ref pairSlot, "ظرفیت MV", "CapacityMv", 0, 1000000, 1);
            row = AddNumericPair(grid, row, ref pairSlot, "ظرفیت توزیع‌شده", "DistributedCapacity", 0, 1000000, 1);
            row = AddNumericPair(grid, row, ref pairSlot, "مساحت کل", "TotalArea", 0, 10000000, 1);
            row = AddNumericPair(grid, row, ref pairSlot, "مساحت ساختمان", "BuildingsArea", 0, 10000000, 1);
            row = AddNumericPair(grid, row, ref pairSlot, "مساحت محوطه", "SiteArea", 0, 10000000, 1);
            row = AddNumericPair(grid, row, ref pairSlot, "تعداد پارکینگ", "ParkingBays", 0, 100000, 0);
            row = AddNumericPair(grid, row, ref pairSlot, "فاصله تا مرکز", "DistanceToCapital", 0, 10000000, 1);

            AddGroupHeader(grid, ref row, ref pairSlot, "📝 سایر اطلاعات");
            row = AddTextPair(grid, row, ref pairSlot, "کد پستی", "PostalCode", 20);
            row = AddTextPair(grid, row, ref pairSlot, "طول جغرافیایی", "Longitude", 60);
            row = AddTextPair(grid, row, ref pairSlot, "عرض جغرافیایی", "Latitude", 60);
            row = AddTextPair(grid, row, ref pairSlot, "آدرس", "PostalAddress", 255);
            row = AddFullText(grid, row, ref pairSlot, "توضیحات", "Description", 255);

            grid.RowCount = row;
            scroll.Controls.Add(grid);
            return scroll;
        }

        private static void AddGroupHeader(TableLayoutPanel grid, ref int row, ref int pairSlot, string text)
        {
            if (pairSlot == 1)
            {
                row++;
                pairSlot = 0;
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

        private static ComboBox CreateCombo()
        {
            return new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                Margin = new Padding(0, 4, 12, 4)
            };
        }

        private int AddComboPair(TableLayoutPanel grid, int row, ref int pairSlot, string caption, out ComboBox combo, string key)
        {
            combo = CreateCombo();
            _lookupCombos[key] = combo;
            return AddField(grid, row, ref pairSlot, caption, combo);
        }

        private int AddComboPair(TableLayoutPanel grid, int row, ref int pairSlot, string caption, string key)
        {
            return AddComboPair(grid, row, ref pairSlot, caption, out _, key);
        }

        private int AddNumericPair(TableLayoutPanel grid, int row, ref int pairSlot, string caption, string key, decimal min, decimal max, int decimals)
        {
            var numeric = new NumericUpDown
            {
                Dock = DockStyle.Fill,
                Minimum = min,
                Maximum = max,
                DecimalPlaces = decimals,
                ThousandsSeparator = true,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                Margin = new Padding(0, 4, 12, 4)
            };
            _numericBoxes[key] = numeric;
            return AddField(grid, row, ref pairSlot, caption, numeric);
        }

        private int AddTextPair(TableLayoutPanel grid, int row, ref int pairSlot, string caption, string key, int maxLength)
        {
            var textBox = new TextBox
            {
                Dock = DockStyle.Fill,
                MaxLength = maxLength,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                Margin = new Padding(0, 4, 12, 4)
            };
            _textBoxes[key] = textBox;
            return AddField(grid, row, ref pairSlot, caption, textBox);
        }

        private int AddFullText(TableLayoutPanel grid, int row, ref int pairSlot, string caption, string key, int maxLength)
        {
            var textBox = new TextBox
            {
                Dock = DockStyle.Fill,
                MaxLength = maxLength,
                Multiline = true,
                Height = 64,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font(FontSettings.FontFamilyName, 9F),
                Margin = new Padding(0, 4, 12, 4)
            };
            _textBoxes[key] = textBox;
            grid.Controls.Add(CreateCaption(caption), 0, row);
            grid.Controls.Add(textBox, 1, row);
            grid.SetColumnSpan(textBox, 3);
            row++;
            pairSlot = 0;
            return row;
        }

        private int AddField(TableLayoutPanel grid, int row, ref int pairSlot, string caption, Control control)
        {
            int captionColumn = pairSlot == 0 ? 0 : 2;
            int controlColumn = pairSlot == 0 ? 1 : 3;
            grid.Controls.Add(CreateCaption(caption), captionColumn, row);
            grid.Controls.Add(control, controlColumn, row);
            pairSlot++;
            if (pairSlot == 2)
            {
                pairSlot = 0;
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

            var btnSave = CreateActionButton("💾 ذخیره", Color.FromArgb(16, 185, 129));
            btnSave.Dock = DockStyle.Right;
            btnSave.Width = 160;
            btnSave.Click += (_, _) => Save();

            var btnCancel = CreateActionButton("❌ لغو", Color.FromArgb(239, 68, 68));
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
        // بارگذاری داده‌ها
        // ─────────────────────────────────────────────
        private void LoadLookups()
        {
            _loading = true;
            try
            {
                FillCombo(cmbProvince, _crud.GetLookupValues("Provinces", "ProvinceID", "ProvinceName"));

                // بارگذاری کامل داده‌های وابسته در حافظه
                _cities = LoadDependent("Cities", "CityID", "CityName", "ProvinceID");
                _affairs = LoadDependent("TransferAffairs", "AffairID", "AffairName", "ProvinceID");
                _depts = LoadDependent("OperationDepartments", "DeptID", "DeptName", "AffairID");
                _districts = LoadDependent("Districts", "DistrictID", "DistrictName", "DeptID");
                _postNames = LoadDependent("PostsNames", "PostNameID", "PostName", "DistrictID");

                FillCombo(cmbCity, ToPairs(_cities));
                FillCombo(cmbAffair, ToPairs(_affairs));
                FillCombo(cmbDept, ToPairs(_depts));
                FillCombo(cmbDistrict, ToPairs(_districts));
                FillCombo(cmbPostName, ToPairs(_postNames));

                FillCombo(GetCombo("Voltage"), _crud.GetLookupValues("VoltageLevels", "VoltageID", "VoltageName"));
                FillCombo(GetCombo("Standard"), _crud.GetLookupValues("PostStandards", "StandardID", "StandardName"));
                FillCombo(GetCombo("Type"), _crud.GetLookupValues("PostTypes", "TypeID", "TypeName"));
                FillCombo(GetCombo("Conn"), _crud.GetLookupValues("DistributedConnections", "ConnID", "ConnName"));
                FillCombo(GetCombo("Ins"), _crud.GetLookupValues("InsulationTypes", "InsID", "InsName"));
                FillCombo(GetCombo("Pt2"), _crud.GetLookupValues("PostTypeTwos", "PT2ID", "PT2Name"));
                FillCombo(GetCombo("Fm"), _crud.GetLookupValues("FixedMobiles", "FMID", "FMName"));
                FillCombo(GetCombo("Circuit"), _crud.GetLookupValues("CircuitStatuses", "CircuitID", "CircuitName"));
                FillCombo(GetCombo("Diesel"), _crud.GetLookupValues("DieselGenerators", "DieselID", "DieselName"));
                FillCombo(GetCombo("Feed"), _crud.GetLookupValues("DistributionFeeds", "FeedID", "FeedName"));
                FillCombo(GetCombo("Water"), _crud.GetLookupValues("WaterStatuses", "WaterID", "WaterName"));
                FillCombo(GetCombo("Guest"), _crud.GetLookupValues("GuestHouses", "GuestID", "GuestName"));

                cmbProvince.SelectedIndexChanged += (_, _) => OnProvinceChanged();
                cmbAffair.SelectedIndexChanged += (_, _) => OnAffairChanged();
                cmbDept.SelectedIndexChanged += (_, _) => OnDeptChanged();
                cmbDistrict.SelectedIndexChanged += (_, _) => OnDistrictChanged();
            }
            finally
            {
                _loading = false;
            }
        }

        private static void FillCombo(ComboBox combo, List<KeyValuePair<int, string>> values)
        {
            combo.Items.Clear();
            foreach (KeyValuePair<int, string> pair in values)
                combo.Items.Add(pair);
            combo.DisplayMember = nameof(KeyValuePair<int, string>.Value);
            combo.ValueMember = nameof(KeyValuePair<int, string>.Key);
            if (combo.Items.Count > 0)
                combo.SelectedIndex = -1;
        }

        private List<LookupItem> LoadDependent(string table, string idColumn, string nameColumn, string parentColumn)
        {
            var result = new List<LookupItem>();
            var db = new DbHelper();
            DataTable? data = db.ExecuteQuery(
                $"SELECT [{idColumn}], [{nameColumn}], [{parentColumn}] FROM [{table}] ORDER BY [{nameColumn}]",
                null,
                showError: false);
            if (data == null)
                return result;

            foreach (DataRow row in data.Rows)
            {
                int id = ToInt(row[idColumn]);
                string name = row[nameColumn]?.ToString()?.Trim() ?? string.Empty;
                int parent = ToInt(row[parentColumn]);
                if (id > 0 && !string.IsNullOrWhiteSpace(name))
                    result.Add(new LookupItem(id, name, parent));
            }

            return result;
        }

        private static List<KeyValuePair<int, string>> ToPairs(List<LookupItem> items)
            => items.Select(item => new KeyValuePair<int, string>(item.Id, item.Name)).ToList();

        private ComboBox GetCombo(string key) => _lookupCombos[key];

        private void OnProvinceChanged()
        {
            if (_loading)
                return;

            int provinceId = SelectedId(cmbProvince);
            FillCombo(cmbCity, ToPairs(_cities.Where(item => provinceId == 0 || item.ParentId == provinceId).ToList()));
            FillCombo(cmbAffair, ToPairs(_affairs.Where(item => provinceId == 0 || item.ParentId == provinceId).ToList()));
            RefreshDeptDistrictPostName();
        }

        private void OnAffairChanged()
        {
            if (_loading)
                return;

            int affairId = SelectedId(cmbAffair);
            FillCombo(cmbDept, ToPairs(_depts.Where(item => affairId == 0 || item.ParentId == affairId).ToList()));
            RefreshDistrictPostName();
        }

        private void OnDeptChanged()
        {
            if (_loading)
                return;

            int deptId = SelectedId(cmbDept);
            FillCombo(cmbDistrict, ToPairs(_districts.Where(item => deptId == 0 || item.ParentId == deptId).ToList()));
            RefreshPostName();
        }

        private void OnDistrictChanged()
        {
            if (_loading)
                return;

            int districtId = SelectedId(cmbDistrict);
            FillCombo(cmbPostName, ToPairs(_postNames.Where(item => districtId == 0 || item.ParentId == districtId).ToList()));
        }

        private void RefreshDeptDistrictPostName()
        {
            int affairId = SelectedId(cmbAffair);
            FillCombo(cmbDept, ToPairs(_depts.Where(item => affairId == 0 || item.ParentId == affairId).ToList()));
            RefreshDistrictPostName();
        }

        private void RefreshDistrictPostName()
        {
            int deptId = SelectedId(cmbDept);
            FillCombo(cmbDistrict, ToPairs(_districts.Where(item => deptId == 0 || item.ParentId == deptId).ToList()));
            RefreshPostName();
        }

        private void RefreshPostName()
        {
            int districtId = SelectedId(cmbDistrict);
            FillCombo(cmbPostName, ToPairs(_postNames.Where(item => districtId == 0 || item.ParentId == districtId).ToList()));
        }

        private void LoadExisting()
        {
            var db = new DbHelper();
            DataTable? posts = db.ExecuteQuery("SELECT * FROM Posts WHERE PostID = ?",
                new[] { new System.Data.OleDb.OleDbParameter("?", _postId!.Value) },
                showError: false);
            if (posts == null || posts.Rows.Count == 0)
                return;

            DataRow row = posts.Rows[0];
            _loading = true;
            try
            {
                SelectId(cmbProvince, ToInt(row["ProvinceID"]));
                OnProvinceChanged();
                SelectId(cmbCity, ToInt(row["CityID"]));
                SelectId(cmbAffair, ToInt(row["AffairID"]));
                OnAffairChanged();
                SelectId(cmbDept, ToInt(row["DeptID"]));
                OnDeptChanged();
                SelectId(cmbDistrict, ToInt(row["DistrictID"]));
                OnDistrictChanged();
                SelectId(cmbPostName, ToInt(row["PostNameID"]));

                SelectId(GetCombo("Voltage"), ToInt(row["VoltageID"]));
                SelectId(GetCombo("Standard"), ToInt(row["StandardID"]));
                SelectId(GetCombo("Type"), ToInt(row["TypeID"]));
                SelectId(GetCombo("Conn"), ToInt(row["ConnID"]));
                SelectId(GetCombo("Ins"), ToInt(row["InsID"]));
                SelectId(GetCombo("Pt2"), ToInt(row["PT2ID"]));
                SelectId(GetCombo("Fm"), ToInt(row["FMID"]));
                SelectId(GetCombo("Circuit"), ToInt(row["CircuitID"]));
                SelectId(GetCombo("Diesel"), ToInt(row["DieselID"]));
                SelectId(GetCombo("Feed"), ToInt(row["FeedID"]));
                SelectId(GetCombo("Water"), ToInt(row["WaterID"]));
                SelectId(GetCombo("Guest"), ToInt(row["GuestID"]));

                _numericBoxes["OperationYear"].Value = Clamp(row["OperationYear"], 0, 2100);
                _numericBoxes["CapacityHv"].Value = Clamp(row["CapacityHV"], 0, 1000000);
                _numericBoxes["CapacityMv"].Value = Clamp(row["CapacityMV"], 0, 1000000);
                _numericBoxes["DistributedCapacity"].Value = Clamp(row["DistributedCapacity"], 0, 1000000);
                _numericBoxes["TotalArea"].Value = Clamp(row["TotalArea"], 0, 10000000);
                _numericBoxes["BuildingsArea"].Value = Clamp(row["BuildingsArea"], 0, 10000000);
                _numericBoxes["SiteArea"].Value = Clamp(row["SiteArea"], 0, 10000000);
                _numericBoxes["ParkingBays"].Value = Clamp(row["ParkingBays"], 0, 100000);
                _numericBoxes["DistanceToCapital"].Value = Clamp(row["DistanceToCapital"], 0, 10000000);

                _textBoxes["PostalCode"].Text = row["PostalCode"]?.ToString() ?? "";
                _textBoxes["Longitude"].Text = row["Longitude"]?.ToString() ?? "";
                _textBoxes["Latitude"].Text = row["Latitude"]?.ToString() ?? "";
                _textBoxes["PostalAddress"].Text = row["PostalAddress"]?.ToString() ?? "";
                _textBoxes["Description"].Text = row["Description"]?.ToString() ?? "";
            }
            finally
            {
                _loading = false;
            }
        }

        private static int ToInt(object? value)
        {
            if (value == null || value == DBNull.Value)
                return 0;

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private static decimal? ToDecimal(object? value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            return Convert.ToDecimal(value, CultureInfo.InvariantCulture);
        }

        private static decimal Clamp(object? value, decimal min, decimal max)
        {
            decimal? d = ToDecimal(value);
            if (!d.HasValue)
                return min;

            return Math.Max(min, Math.Min(max, d.Value));
        }

        private static void SelectId(ComboBox combo, int id)
        {
            if (id <= 0)
                return;

            for (int i = 0; i < combo.Items.Count; i++)
            {
                if (combo.Items[i] is KeyValuePair<int, string> pair && pair.Key == id)
                {
                    combo.SelectedIndex = i;
                    return;
                }
            }
        }

        private static int SelectedId(ComboBox combo)
            => combo.SelectedItem is KeyValuePair<int, string> pair ? pair.Key : 0;

        // ─────────────────────────────────────────────
        // ذخیره
        // ─────────────────────────────────────────────
        private void Save()
        {
            try
            {
                var data = new PostEditData
                {
                    ProvinceId = SelectedId(cmbProvince),
                    CityId = SelectedId(cmbCity),
                    AffairId = SelectedId(cmbAffair),
                    DeptId = SelectedId(cmbDept),
                    DistrictId = SelectedId(cmbDistrict),
                    PostNameId = SelectedId(cmbPostName),
                    VoltageId = SelectedId(GetCombo("Voltage")),
                    StandardId = SelectedId(GetCombo("Standard")),
                    TypeId = SelectedId(GetCombo("Type")),
                    ConnId = SelectedId(GetCombo("Conn")),
                    InsId = SelectedId(GetCombo("Ins")),
                    Pt2Id = SelectedId(GetCombo("Pt2")),
                    FmId = SelectedId(GetCombo("Fm")),
                    CircuitId = SelectedId(GetCombo("Circuit")),
                    DieselId = SelectedId(GetCombo("Diesel")),
                    FeedId = SelectedId(GetCombo("Feed")),
                    WaterId = SelectedId(GetCombo("Water")),
                    GuestId = SelectedId(GetCombo("Guest")),
                    OperationYear = (int)_numericBoxes["OperationYear"].Value,
                    CapacityHv = (double)_numericBoxes["CapacityHv"].Value,
                    CapacityMv = (double)_numericBoxes["CapacityMv"].Value,
                    DistributedCapacity = (double)_numericBoxes["DistributedCapacity"].Value,
                    TotalArea = (double)_numericBoxes["TotalArea"].Value,
                    BuildingsArea = (double)_numericBoxes["BuildingsArea"].Value,
                    SiteArea = (double)_numericBoxes["SiteArea"].Value,
                    ParkingBays = (int)_numericBoxes["ParkingBays"].Value,
                    DistanceToCapital = (double)_numericBoxes["DistanceToCapital"].Value,
                    PostalCode = _textBoxes["PostalCode"].Text.Trim(),
                    Longitude = _textBoxes["Longitude"].Text.Trim(),
                    Latitude = _textBoxes["Latitude"].Text.Trim(),
                    PostalAddress = _textBoxes["PostalAddress"].Text.Trim(),
                    Description = _textBoxes["Description"].Text.Trim()
                };

                bool ok;
                if (_postId.HasValue)
                {
                    ok = _crud.UpdatePost(_postId.Value, data);
                }
                else
                {
                    int newId = _crud.InsertPost(data);
                    ok = newId > 0;
                }

                if (!ok)
                {
                    MessageBox.Show(this,
                        "ذخیره‌سازی انجام نشد. لطفاً اتصال پایگاه داده را بررسی کنید.",
                        "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPostsEditor.Save", ex);
                MessageBox.Show(this, $"خطا در ذخیره پست:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
