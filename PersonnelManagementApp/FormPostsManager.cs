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
    /// مدیریت پست‌ها: جستجو، افزودن، ویرایش، حذف و خروجی اکسل.
    /// </summary>
    public sealed class FormPostsManager : Form
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

        private readonly PostsDataService _postsService = new();
        private readonly PostsCrudService _crud = new();

        private DataTable? _allData;
        private readonly TextBox txtSearch = new();
        private readonly ComboBox cmbProvince = new();
        private DataGridView dgvPosts = null!;
        private Label lblStatus = null!;
        private Label lblCount = null!;
        private bool _suppressFilter;

        public FormPostsManager()
        {
            BuildUi();
            FontSettings.ApplyFontToForm(this);
            LoadData();
        }

        private void BuildUi()
        {
            Text = "مدیریت پست‌ها";
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
                Text = "مدیریت پست‌های برق",
                Dock = DockStyle.Top,
                Height = 34,
                Font = new Font(FontSettings.FontFamilyName, 17F, FontStyle.Bold),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };

            var subtitle = new Label
            {
                Text = "جستجو، افزودن، ویرایش و حذف پست‌های برق",
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

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                BackColor = CardColor,
                Margin = new Padding(0)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 280F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 830F));

            txtSearch.Dock = DockStyle.Fill;
            txtSearch.Font = new Font(FontSettings.FontFamilyName, 9.5F);
            txtSearch.PlaceholderText = "جستجو در نام پست، استان، شهر، اداره یا آدرس...";
            txtSearch.TextChanged += (_, _) => ApplyFilter();

            cmbProvince.Dock = DockStyle.Fill;
            cmbProvince.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbProvince.Font = new Font(FontSettings.FontFamilyName, 9.5F);
            cmbProvince.SelectedIndexChanged += ProvinceFilter_Changed;
            cmbProvince.Items.Add("همه استان‌ها");
            cmbProvince.SelectedIndex = 0;

            var btnAdd = CreateToolButton("➕ ثبت پست جدید", SuccessColor);
            btnAdd.Click += (_, _) => OpenEditor(null);

            var btnEdit = CreateToolButton("✏️ ویرایش", PrimaryColor);
            btnEdit.Click += (_, _) => EditSelected();

            var btnDelete = CreateToolButton("🗑️ حذف", DangerColor);
            btnDelete.Click += (_, _) => DeleteSelected();

            var btnExcel = CreateToolButton("📥 خروجی اکسل", Color.FromArgb(20, 184, 166));
            btnExcel.Click += (_, _) => ExportExcel();

            var btnRefresh = CreateToolButton("🔄 به‌روزرسانی", Color.FromArgb(71, 85, 105));
            btnRefresh.Click += (_, _) => LoadData();

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
            layout.Controls.Add(new Panel { Dock = DockStyle.Fill, BackColor = CardColor }, 2, 0);
            layout.Controls.Add(actions, 3, 0);
            toolbar.Controls.Add(layout);
            return toolbar;
        }

        private static Button CreateToolButton(string text, Color backColor)
        {
            var button = new Button
            {
                Text = text,
                AutoSize = false,
                Width = 150,
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

            dgvPosts = new DataGridView
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
            dgvPosts.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 64, 175);
            dgvPosts.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPosts.ColumnHeadersDefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F, FontStyle.Bold);
            dgvPosts.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPosts.ColumnHeadersHeight = 38;
            dgvPosts.DefaultCellStyle.Font = new Font(FontSettings.FontFamilyName, 9F);
            dgvPosts.DefaultCellStyle.ForeColor = TextColor;
            dgvPosts.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            dgvPosts.DefaultCellStyle.SelectionForeColor = TextColor;
            dgvPosts.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);
            dgvPosts.RowTemplate.Height = 30;

            AddColumn("PostID", "شناسه", 70);
            AddColumn("PostName", "نام پست", 140);
            AddColumn("ProvinceName", "استان", 90);
            AddColumn("CityName", "شهر", 100);
            AddColumn("AffairName", "امور", 100);
            AddColumn("DeptName", "اداره", 120);
            AddColumn("DistrictName", "ناحیه", 100);
            AddColumn("VoltageName", "ولتاژ", 80);
            AddColumn("TypeName", "نوع", 80);
            AddColumn("OperationYear", "سال", 60);
            AddColumn("CapacityHV", "ظرفیت HV", 90);
            AddColumn("CapacityMV", "ظرفیت MV", 90);
            AddColumn("TotalArea", "مساحت", 80);
            AddColumn("PostalAddress", "آدرس", 220);

            dgvPosts.CellDoubleClick += (_, e) =>
            {
                if (e.RowIndex >= 0)
                    EditSelected();
            };

            panel.Controls.Add(dgvPosts);
            return panel;
        }

        private void AddColumn(string name, string header, int width)
        {
            dgvPosts.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = header,
                DataPropertyName = name,
                FillWeight = width,
                MinimumWidth = Math.Min(50, width),
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
                Text = "۰ پست"
            };

            statusPanel.Controls.Add(lblStatus);
            statusPanel.Controls.Add(lblCount);
            return statusPanel;
        }

        private void LoadData()
        {
            UseWaitCursor = true;
            try
            {
                lblStatus.Text = "در حال بارگذاری پست‌ها...";
                _allData = _postsService.LoadAllPosts(showError: false);
                if (_allData == null || _allData.Rows.Count == 0)
                {
                    lblStatus.Text = "⚠️ داده‌ای در جدول پست‌ها یافت نشد.";
                    return;
                }

                FillProvinceFilter();
                ApplyFilter();
                lblStatus.Text = $"✅ {_allData.Rows.Count} پست بارگذاری شد.";
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPostsManager.LoadData", ex);
                lblStatus.Text = "❌ خطا در بارگذاری داده‌ها.";
                MessageBox.Show(this, $"خطا در بارگذاری داده‌ها:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
            }
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
                    .Select(row => row["ProvinceName"]?.ToString() ?? "")
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
                rows = rows.Where(row => string.Equals(row["ProvinceName"]?.ToString(), province, StringComparison.CurrentCultureIgnoreCase));

            if (!string.IsNullOrWhiteSpace(search))
            {
                rows = rows.Where(row =>
                    Contains(row, "PostName", search) ||
                    Contains(row, "ProvinceName", search) ||
                    Contains(row, "CityName", search) ||
                    Contains(row, "DeptName", search) ||
                    Contains(row, "PostalAddress", search));
            }

            DataTable filtered = _allData.Clone();
            foreach (DataRow row in rows)
                filtered.ImportRow(row);

            dgvPosts.DataSource = filtered;
            lblCount.Text = $"{filtered.Rows.Count:N0} پست";
            lblStatus.Text = $"نمایش {filtered.Rows.Count} از {_allData.Rows.Count} پست  |  " +
                             $"آخرین بروزرسانی: {DateTime.Now:HH:mm:ss}";
        }

        private static bool Contains(DataRow row, string column, string search)
            => row.Table.Columns.Contains(column) &&
               (row[column]?.ToString() ?? "").Contains(search, StringComparison.CurrentCultureIgnoreCase);

        private void OpenEditor(int? postId)
        {
            using var editor = new FormPostsEditor(postId);
            if (editor.ShowDialog(this) == DialogResult.OK)
                LoadData();
        }

        private void EditSelected()
        {
            if (dgvPosts.CurrentRow == null || dgvPosts.CurrentRow.DataBoundItem is not DataRowView rowView)
            {
                MessageBox.Show(this, "ابتدا یک پست را از جدول انتخاب کنید.", "اطلاع",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int postId = Convert.ToInt32(rowView["PostID"]);
            OpenEditor(postId);
        }

        private void DeleteSelected()
        {
            var selected = dgvPosts.SelectedRows.Cast<DataGridViewRow>()
                .Where(row => row.DataBoundItem is DataRowView)
                .ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, "ابتدا یک یا چند پست را از جدول انتخاب کنید.", "اطلاع",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string question = selected.Count == 1
                ? "آیا مطمئن هستید که می‌خواهید این پست را حذف کنید؟"
                : $"آیا مطمئن هستید که می‌خواهید {selected.Count} پست را حذف کنید؟";

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
                    int postId = Convert.ToInt32(((DataRowView)row.DataBoundItem!)["PostID"]);
                    if (_crud.DeletePost(postId))
                        deleted++;
                }

                if (deleted > 0)
                {
                    MessageBox.Show(this, $"✅ {deleted} پست با موفقیت حذف شد.", "موفقیت",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                }
                else
                {
                    MessageBox.Show(this, "هیچ پستی حذف نشد.", "هشدار",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("FormPostsManager.DeleteSelected", ex);
                MessageBox.Show(this, $"خطا در حذف پست:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        private void ExportExcel()
        {
            if (dgvPosts.DataSource is not DataTable dt || dt.Rows.Count == 0)
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
                WriteWorkbook(dt, dialog.FileName);
                UseWaitCursor = false;
                MessageBox.Show(this, "✅ فایل اکسل با موفقیت ذخیره شد.", "خروجی اکسل",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                UseWaitCursor = false;
                ErrorLogger.Log("FormPostsManager.ExportExcel", ex);
                MessageBox.Show(this, $"خطا در ایجاد فایل اکسل:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void WriteWorkbook(DataTable data, string path)
        {
            string[] columns =
            {
                "PostName", "ProvinceName", "CityName", "AffairName", "DeptName", "DistrictName",
                "VoltageName", "StandardName", "TypeName", "ConnName", "InsName", "PT2Name",
                "FMName", "CircuitName", "DieselName", "FeedName", "WaterName", "GuestName",
                "OperationYear", "CapacityHV", "CapacityMV", "DistributedCapacity", "TotalArea",
                "BuildingsArea", "SiteArea", "ParkingBays", "DistanceToCapital",
                "PostalCode", "PostalAddress", "Description"
            };
            string[] headers =
            {
                "نام پست", "استان", "شهر", "امور", "اداره", "ناحیه",
                "سطح ولتاژ", "استاندارد", "نوع پست", "اتصال توزیع", "نوع عایق", "نوع پست ۲",
                "ثابت / سیار", "وضعیت مدار", "دیزل ژنراتور", "نوع توزیع", "وضعیت آب", "مهمان‌سرا",
                "سال بهره‌برداری", "ظرفیت HV", "ظرفیت MV", "ظرفیت توزیع‌شده", "مساحت کل",
                "مساحت ساختمان", "مساحت محوطه", "تعداد پارکینگ", "فاصله تا مرکز",
                "کد پستی", "آدرس", "توضیحات"
            };

            using var workbook = new XLWorkbook();
            IXLWorksheet sheet = workbook.Worksheets.Add("پست‌ها");
            sheet.RightToLeft = true;

            for (int c = 0; c < columns.Length; c++)
                sheet.Cell(1, c + 1).Value = headers[c];

            IXLRange headerRange = sheet.Range(1, 1, 1, columns.Length);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(30, 64, 175);
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            sheet.Row(1).Height = 28;

            for (int r = 0; r < data.Rows.Count; r++)
            {
                for (int c = 0; c < columns.Length; c++)
                {
                    string source = columns[c];
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
    }
}
