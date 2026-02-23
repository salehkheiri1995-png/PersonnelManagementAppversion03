using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    /// <summary>
    /// فرم عمومی برای مدیریت جداول مرجع (Lookup Tables)
    /// این فرم برای مدیریت جداولی مانند پست‌ها، امور، شهر، اداره، ناحیه، سطح شغلی و... استفاده می‌شود
    /// </summary>
    public partial class FormLookupTableManager : Form
    {
        private readonly string tableName;
        private readonly string idColumnName;
        private readonly string nameColumnName;
        private readonly string tableDisplayName;
        private readonly DbHelper dbHelper;

        private DataGridView dgvData;
        private TextBox txtSearch;
        private TextBox txtNewValue;
        private TextBox txtEditValue;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnDelete;
        private Button btnRefresh;
        private Button btnClose;
        private Label lblRecordCount;

        // رنگ‌های مدرن
        private readonly Color PrimaryColor = Color.FromArgb(33, 150, 243);
        private readonly Color AccentColor = Color.FromArgb(76, 175, 80);
        private readonly Color DangerColor = Color.FromArgb(244, 67, 54);
        private readonly Color WarningColor = Color.FromArgb(255, 152, 0);
        private readonly Color BackgroundColor = Color.FromArgb(250, 250, 250);
        private readonly Color CardBackground = Color.White;
        private readonly Color TextPrimary = Color.FromArgb(33, 33, 33);
        private readonly Color TextSecondary = Color.FromArgb(117, 117, 117);

        /// <summary>
        /// سازنده فرم مدیریت جداول مرجع
        /// </summary>
        /// <param name="tableName">نام جدول در دیتابیس</param>
        /// <param name="idColumnName">نام ستون شناسه (ID)</param>
        /// <param name="nameColumnName">نام ستون مقدار (Name)</param>
        /// <param name="displayName">نام نمایشی جدول</param>
        public FormLookupTableManager(string tableName, string idColumnName, string nameColumnName, string displayName)
        {
            this.tableName = tableName;
            this.idColumnName = idColumnName;
            this.nameColumnName = nameColumnName;
            this.tableDisplayName = displayName;
            this.dbHelper = new DbHelper();

            InitializeComponent();
            FontSettings.ApplyFontToForm(this);
            LoadData();
        }

        private Font GetSafeFont(string familyName, float size, FontStyle style = FontStyle.Regular)
        {
            try { return new Font(familyName, size, style); }
            catch { return new Font("Tahoma", size, style); }
        }

        private void InitializeComponent()
        {
            this.Text = $"🗂️ مدیریت {tableDisplayName}";
            this.Size = new Size(900, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.RightToLeft = RightToLeft.Yes;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.BackColor = BackgroundColor;
            this.Padding = new Padding(15);

            // هدر
            Panel headerPanel = CreateHeaderPanel();
            this.Controls.Add(headerPanel);

            // پنل جستجو
            Panel searchPanel = CreateSearchPanel();
            this.Controls.Add(searchPanel);

            // DataGridView
            dgvData = CreateDataGridView();
            this.Controls.Add(dgvData);

            // پنل عملیات
            Panel operationsPanel = CreateOperationsPanel();
            this.Controls.Add(operationsPanel);

            // پنل دکمه‌ها
            Panel buttonPanel = CreateButtonPanel();
            this.Controls.Add(buttonPanel);

            // Label تعداد رکوردها
            lblRecordCount = new Label
            {
                Location = new Point(20, 575),
                Size = new Size(300, 25),
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblRecordCount);
        }

        private Panel CreateHeaderPanel()
        {
            Panel panel = new Panel
            {
                Location = new Point(15, 15),
                Size = new Size(850, 60),
                BackColor = PrimaryColor
            };
            ApplyRoundedCorners(panel, 10);

            panel.Controls.Add(new Label
            {
                Text = $"🗂️ مدیریت {tableDisplayName}",
                Font = GetSafeFont(FontSettings.TitleFont?.FontFamily.Name ?? "Tahoma", 16, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(20, 10),
                Size = new Size(600, 30),
                TextAlign = ContentAlignment.MiddleRight
            });

            panel.Controls.Add(new Label
            {
                Text = $"افزودن، ویرایش و حذف رکوردهای {tableDisplayName}",
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = Color.FromArgb(230, 240, 255),
                Location = new Point(20, 38),
                Size = new Size(400, 18),
                TextAlign = ContentAlignment.TopRight
            });

            return panel;
        }

        private Panel CreateSearchPanel()
        {
            Panel panel = new Panel
            {
                Location = new Point(15, 85),
                Size = new Size(850, 60),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(panel, 8);
            ApplyCardShadow(panel);

            panel.Controls.Add(new Label
            {
                Text = "🔍 جستجو:",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(750, 18),
                Size = new Size(80, 25),
                TextAlign = ContentAlignment.MiddleRight
            });

            txtSearch = new TextBox
            {
                Location = new Point(420, 18),
                Size = new Size(320, 28),
                Font = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 10),
                BorderStyle = BorderStyle.FixedSingle
            };
            txtSearch.TextChanged += TxtSearch_TextChanged;
            panel.Controls.Add(txtSearch);

            btnRefresh = CreateModernButton("🔄 بروزرسانی", PrimaryColor, 120, 32);
            btnRefresh.Location = new Point(20, 15);
            btnRefresh.Click += (s, e) => LoadData();
            panel.Controls.Add(btnRefresh);

            return panel;
        }

        private DataGridView CreateDataGridView()
        {
            DataGridView dgv = new DataGridView
            {
                Location = new Point(15, 155),
                Size = new Size(850, 280),
                BackgroundColor = CardBackground,
                BorderStyle = BorderStyle.None,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                EnableHeadersVisualStyles = false,
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9)
            };

            dgv.ColumnHeadersDefaultCellStyle.BackColor = PrimaryColor;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.ColumnHeadersHeight = 40;

            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 249, 250);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(200, 230, 255);
            dgv.DefaultCellStyle.SelectionForeColor = TextPrimary;

            dgv.SelectionChanged += DgvData_SelectionChanged;

            return dgv;
        }

        private Panel CreateOperationsPanel()
        {
            Panel panel = new Panel
            {
                Location = new Point(15, 445),
                Size = new Size(850, 110),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(panel, 8);
            ApplyCardShadow(panel);

            // بخش افزودن
            panel.Controls.Add(new Label
            {
                Text = "➕ افزودن رکورد جدید:",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(650, 15),
                Size = new Size(180, 25),
                TextAlign = ContentAlignment.MiddleRight
            });

            txtNewValue = new TextBox
            {
                Location = new Point(350, 15),
                Size = new Size(290, 28),
                Font = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 10),
                BorderStyle = BorderStyle.FixedSingle
            };
            panel.Controls.Add(txtNewValue);

            btnAdd = CreateModernButton("➕ افزودن", AccentColor, 120, 32);
            btnAdd.Location = new Point(220, 13);
            btnAdd.Click += BtnAdd_Click;
            panel.Controls.Add(btnAdd);

            // بخش ویرایش
            panel.Controls.Add(new Label
            {
                Text = "✏️ ویرایش رکورد انتخابی:",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 9, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(650, 60),
                Size = new Size(180, 25),
                TextAlign = ContentAlignment.MiddleRight
            });

            txtEditValue = new TextBox
            {
                Location = new Point(350, 60),
                Size = new Size(290, 28),
                Font = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 10),
                BorderStyle = BorderStyle.FixedSingle,
                Enabled = false
            };
            panel.Controls.Add(txtEditValue);

            btnEdit = CreateModernButton("✏️ ویرایش", WarningColor, 100, 32);
            btnEdit.Location = new Point(240, 58);
            btnEdit.Enabled = false;
            btnEdit.Click += BtnEdit_Click;
            panel.Controls.Add(btnEdit);

            btnDelete = CreateModernButton("🗑️ حذف", DangerColor, 100, 32);
            btnDelete.Location = new Point(130, 58);
            btnDelete.Enabled = false;
            btnDelete.Click += BtnDelete_Click;
            panel.Controls.Add(btnDelete);

            return panel;
        }

        private Panel CreateButtonPanel()
        {
            Panel panel = new Panel
            {
                Location = new Point(15, 565),
                Size = new Size(850, 50),
                BackColor = Color.Transparent
            };

            btnClose = CreateModernButton("❌ بستن", DangerColor, 120, 38);
            btnClose.Location = new Point(730, 6);
            btnClose.Font = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            btnClose.Click += (s, e) => this.Close();
            panel.Controls.Add(btnClose);

            return panel;
        }

        private Button CreateModernButton(string text, Color backColor, int width, int height)
        {
            Button btn = new Button
            {
                Text = text,
                Size = new Size(width, height),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 9)
            };
            btn.FlatAppearance.BorderSize = 0;
            ApplyRoundedCorners(btn, 6);

            Color orig = backColor;
            btn.MouseEnter += (s, e) => btn.BackColor = ControlPaint.Light(orig, 0.1f);
            btn.MouseLeave += (s, e) => btn.BackColor = orig;
            return btn;
        }

        private void LoadData(string searchTerm = "")
        {
            try
            {
                string query = string.IsNullOrEmpty(searchTerm)
                    ? $"SELECT {idColumnName}, {nameColumnName} FROM {tableName} ORDER BY {nameColumnName}"
                    : $"SELECT {idColumnName}, {nameColumnName} FROM {tableName} WHERE {nameColumnName} LIKE ? ORDER BY {nameColumnName}";

                OleDbParameter[]? parameters = string.IsNullOrEmpty(searchTerm)
                    ? null
                    : new OleDbParameter[] { new OleDbParameter("?", $"%{searchTerm}%") };

                DataTable? dt = dbHelper.ExecuteQuery(query, parameters);

                if (dt != null && dt.Rows.Count > 0)
                {
                    dgvData.DataSource = dt;
                    
                    if (dgvData.Columns.Count >= 2)
                    {
                        // Temporarily disable Fill mode to set custom widths
                        dgvData.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                        
                        dgvData.Columns[0].HeaderText = "شناسه";
                        dgvData.Columns[0].Width = 100;
                        dgvData.Columns[1].HeaderText = tableDisplayName;
                        
                        // Re-enable Fill mode for the second column
                        dgvData.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    }

                    lblRecordCount.Text = $"📊 تعداد رکوردها: {dt.Rows.Count}";
                }
                else
                {
                    dgvData.DataSource = null;
                    lblRecordCount.Text = "📊 تعداد رکوردها: 0";
                    if (!string.IsNullOrEmpty(searchTerm))
                    {
                        MessageBox.Show("❌ رکوردی یافت نشد.", "جستجو", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در بارگذاری داده‌ها:\n\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            LoadData(txtSearch.Text.Trim());
        }

        private void DgvData_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvData.SelectedRows.Count > 0)
            {
                var selectedRow = dgvData.SelectedRows[0];
                txtEditValue.Text = selectedRow.Cells[nameColumnName].Value?.ToString() ?? "";
                txtEditValue.Enabled = true;
                btnEdit.Enabled = true;
                btnDelete.Enabled = true;
            }
            else
            {
                txtEditValue.Text = "";
                txtEditValue.Enabled = false;
                btnEdit.Enabled = false;
                btnDelete.Enabled = false;
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string newValue = txtNewValue.Text.Trim();

            if (string.IsNullOrEmpty(newValue))
            {
                MessageBox.Show("⚠️ لطفاً مقدار جدید را وارد کنید.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtNewValue.Focus();
                return;
            }

            // بررسی تکراری بودن
            string checkQuery = $"SELECT COUNT(*) FROM {tableName} WHERE {nameColumnName} = ?";
            DataTable? checkResult = dbHelper.ExecuteQuery(checkQuery, new OleDbParameter[] { new OleDbParameter("?", newValue) });

            if (checkResult != null && Convert.ToInt32(checkResult.Rows[0][0]) > 0)
            {
                MessageBox.Show($"⚠️ رکورد '{newValue}' قبلاً وجود دارد.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show($"آیا از افزودن '{newValue}' اطمینان دارید؟", "تأیید", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    string insertQuery = $"INSERT INTO {tableName} ({nameColumnName}) VALUES (?)";
                    int result = dbHelper.ExecuteNonQuery(insertQuery, new OleDbParameter[] { new OleDbParameter("?", newValue) });

                    if (result > 0)
                    {
                        MessageBox.Show($"✅ رکورد '{newValue}' با موفقیت اضافه شد.", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtNewValue.Clear();
                        LoadData();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"❌ خطا در افزودن رکورد:\n\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (dgvData.SelectedRows.Count == 0)
            {
                MessageBox.Show("⚠️ لطفاً یک رکورد را انتخاب کنید.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedRow = dgvData.SelectedRows[0];
            int recordId = Convert.ToInt32(selectedRow.Cells[idColumnName].Value);
            string oldValue = selectedRow.Cells[nameColumnName].Value?.ToString() ?? "";
            string newValue = txtEditValue.Text.Trim();

            if (string.IsNullOrEmpty(newValue))
            {
                MessageBox.Show("⚠️ لطفاً مقدار جدید را وارد کنید.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtEditValue.Focus();
                return;
            }

            if (oldValue == newValue)
            {
                MessageBox.Show("⚠️ مقدار جدید با مقدار قبلی یکسان است.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // بررسی تکراری بودن
            string checkQuery = $"SELECT COUNT(*) FROM {tableName} WHERE {nameColumnName} = ? AND {idColumnName} <> ?";
            DataTable? checkResult = dbHelper.ExecuteQuery(checkQuery, new OleDbParameter[] {
                new OleDbParameter("?", newValue),
                new OleDbParameter("?", recordId)
            });

            if (checkResult != null && Convert.ToInt32(checkResult.Rows[0][0]) > 0)
            {
                MessageBox.Show($"⚠️ رکورد '{newValue}' قبلاً وجود دارد.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show($"آیا از ویرایش '{oldValue}' به '{newValue}' اطمینان دارید؟", "تأیید", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    string updateQuery = $"UPDATE {tableName} SET {nameColumnName} = ? WHERE {idColumnName} = ?";
                    int result = dbHelper.ExecuteNonQuery(updateQuery, new OleDbParameter[] {
                        new OleDbParameter("?", newValue),
                        new OleDbParameter("?", recordId)
                    });

                    if (result > 0)
                    {
                        MessageBox.Show($"✅ رکورد با موفقیت به '{newValue}' تغییر یافت.", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadData();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"❌ خطا در ویرایش رکورد:\n\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (dgvData.SelectedRows.Count == 0)
            {
                MessageBox.Show("⚠️ لطفاً یک رکورد را انتخاب کنید.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedRow = dgvData.SelectedRows[0];
            int recordId = Convert.ToInt32(selectedRow.Cells[idColumnName].Value);
            string recordValue = selectedRow.Cells[nameColumnName].Value?.ToString() ?? "";

            if (MessageBox.Show(
                $"⚠️ آیا از حذف '{recordValue}' اطمینان دارید?\n\n" +
                $"توجه: اگر این رکورد در جداول دیگر استفاده شده باشد، ممکن است خطا رخ دهد.",
                "تأیید حذف",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                try
                {
                    string deleteQuery = $"DELETE FROM {tableName} WHERE {idColumnName} = ?";
                    int result = dbHelper.ExecuteNonQuery(deleteQuery, new OleDbParameter[] { new OleDbParameter("?", recordId) });

                    if (result > 0)
                    {
                        MessageBox.Show($"✅ رکورد '{recordValue}' با موفقیت حذف شد.", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtEditValue.Clear();
                        LoadData();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"❌ خطا در حذف رکورد:\n\n{ex.Message}\n\n" +
                        $"احتمالاً این رکورد در جداول دیگر استفاده شده است.",
                        "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ApplyRoundedCorners(Control control, int radius)
        {
            try
            {
                GraphicsPath path = new GraphicsPath();
                path.AddArc(0, 0, radius, radius, 180, 90);
                path.AddArc(control.Width - radius, 0, radius, radius, 270, 90);
                path.AddArc(control.Width - radius, control.Height - radius, radius, radius, 0, 90);
                path.AddArc(0, control.Height - radius, radius, radius, 90, 90);
                path.CloseFigure();
                control.Region = new Region(path);
            }
            catch { }
        }

        private void ApplyCardShadow(Panel panel)
        {
            panel.Paint += (s, e) =>
            {
                using (SolidBrush shadowBrush = new SolidBrush(Color.FromArgb(10, 0, 0, 0)))
                    e.Graphics.FillRectangle(shadowBrush, new Rectangle(3, 3, panel.Width - 3, panel.Height - 3));
            };
        }
    }
}