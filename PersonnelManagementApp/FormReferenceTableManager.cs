using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    public partial class FormReferenceTableManager : Form
    {
        private readonly string tableName;
        private readonly string idColumnName;
        private readonly string nameColumnName;
        private readonly string displayTitle;
        private readonly DbHelper dbHelper;

        private DataGridView dgvData;
        private TextBox txtNewItem;
        private TextBox txtSearch;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnDelete;
        private Button btnRefresh;
        private Button btnClose;
        private Label lblTitle;
        private Label lblCount;

        // رنگ‌های مدرن
        private readonly Color PrimaryColor = Color.FromArgb(33, 150, 243);
        private readonly Color AccentColor = Color.FromArgb(76, 175, 80);
        private readonly Color DangerColor = Color.FromArgb(244, 67, 54);
        private readonly Color BackgroundColor = Color.FromArgb(250, 250, 250);
        private readonly Color CardBackground = Color.White;
        private readonly Color TextPrimary = Color.FromArgb(33, 33, 33);
        private readonly Color TextSecondary = Color.FromArgb(117, 117, 117);

        public FormReferenceTableManager(string tableName, string idColumnName, string nameColumnName, string displayTitle)
        {
            this.tableName = tableName;
            this.idColumnName = idColumnName;
            this.nameColumnName = nameColumnName;
            this.displayTitle = displayTitle;
            this.dbHelper = new DbHelper();

            InitializeComponent();
            LoadData();
            FontSettings.ApplyFontToForm(this);
        }

        private Font GetSafeFont(string familyName, float size, FontStyle style = FontStyle.Regular)
        {
            try { return new Font(familyName, size, style); }
            catch { return new Font("Tahoma", size, style); }
        }

        private void InitializeComponent()
        {
            this.Text = $"🗂️ مدیریت {displayTitle}";
            this.Size = new Size(900, 650);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.RightToLeft = RightToLeft.Yes;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = BackgroundColor;
            this.Padding = new Padding(15);

            // Header Panel
            Panel headerPanel = new Panel
            {
                Location = new Point(15, 15),
                Size = new Size(850, 70),
                BackColor = PrimaryColor
            };
            ApplyRoundedCorners(headerPanel, 12);

            lblTitle = new Label
            {
                Text = $"📋 {displayTitle}",
                Font = GetSafeFont(FontSettings.TitleFont?.FontFamily.Name ?? "Tahoma", 16, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(20, 10),
                Size = new Size(400, 35),
                TextAlign = ContentAlignment.MiddleRight
            };
            headerPanel.Controls.Add(lblTitle);

            lblCount = new Label
            {
                Text = "تعداد رکوردها: 0",
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9),
                ForeColor = Color.FromArgb(230, 240, 255),
                Location = new Point(20, 45),
                Size = new Size(400, 20),
                TextAlign = ContentAlignment.TopRight
            };
            headerPanel.Controls.Add(lblCount);

            this.Controls.Add(headerPanel);

            // Search Panel
            Panel searchPanel = new Panel
            {
                Location = new Point(15, 100),
                Size = new Size(850, 50),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(searchPanel, 10);

            Label lblSearch = new Label
            {
                Text = "🔍 جستجو:",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(740, 13),
                Size = new Size(90, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            searchPanel.Controls.Add(lblSearch);

            txtSearch = new TextBox
            {
                Location = new Point(380, 13),
                Size = new Size(350, 28),
                Font = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 10),
                BorderStyle = BorderStyle.FixedSingle
            };
            txtSearch.TextChanged += TxtSearch_TextChanged;
            searchPanel.Controls.Add(txtSearch);

            btnRefresh = CreateModernButton("🔄 بروزرسانی", AccentColor, 120, 30);
            btnRefresh.Location = new Point(250, 12);
            btnRefresh.Click += (s, e) => LoadData();
            searchPanel.Controls.Add(btnRefresh);

            this.Controls.Add(searchPanel);

            // DataGridView
            dgvData = new DataGridView
            {
                Location = new Point(15, 165),
                Size = new Size(850, 320),
                BackgroundColor = CardBackground,
                BorderStyle = BorderStyle.None,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                Font = GetSafeFont(FontSettings.BodyFont?.FontFamily.Name ?? "Tahoma", 9.5f)
            };
            dgvData.ColumnHeadersDefaultCellStyle.BackColor = PrimaryColor;
            dgvData.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvData.ColumnHeadersDefaultCellStyle.Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            dgvData.ColumnHeadersHeight = 35;
            dgvData.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            ApplyRoundedCorners(dgvData, 10);

            this.Controls.Add(dgvData);

            // Add/Edit Panel
            Panel editPanel = new Panel
            {
                Location = new Point(15, 500),
                Size = new Size(850, 60),
                BackColor = CardBackground
            };
            ApplyRoundedCorners(editPanel, 10);

            Label lblNewItem = new Label
            {
                Text = "➕ افزودن/ویرایش:",
                Font = GetSafeFont(FontSettings.LabelFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(700, 18),
                Size = new Size(130, 25),
                TextAlign = ContentAlignment.MiddleRight
            };
            editPanel.Controls.Add(lblNewItem);

            txtNewItem = new TextBox
            {
                Location = new Point(340, 18),
                Size = new Size(350, 28),
                Font = GetSafeFont(FontSettings.TextBoxFont?.FontFamily.Name ?? "Tahoma", 10),
                BorderStyle = BorderStyle.FixedSingle
            };
            editPanel.Controls.Add(txtNewItem);

            btnAdd = CreateModernButton("➕ افزودن", AccentColor, 110, 30);
            btnAdd.Location = new Point(220, 16);
            btnAdd.Click += BtnAdd_Click;
            editPanel.Controls.Add(btnAdd);

            btnEdit = CreateModernButton("✏️ ویرایش", Color.FromArgb(255, 152, 0), 110, 30);
            btnEdit.Location = new Point(100, 16);
            btnEdit.Click += BtnEdit_Click;
            editPanel.Controls.Add(btnEdit);

            btnDelete = CreateModernButton("🗑️ حذف", DangerColor, 90, 30);
            btnDelete.Location = new Point(5, 16);
            btnDelete.Click += BtnDelete_Click;
            editPanel.Controls.Add(btnDelete);

            this.Controls.Add(editPanel);

            // Close Button
            btnClose = CreateModernButton("❌ بستن", Color.FromArgb(158, 158, 158), 120, 38);
            btnClose.Location = new Point(745, 575);
            btnClose.Font = GetSafeFont(FontSettings.ButtonFont?.FontFamily.Name ?? "Tahoma", 10, FontStyle.Bold);
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
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
            ApplyRoundedCorners(btn, 8);

            Color orig = backColor;
            btn.MouseEnter += (s, e) => btn.BackColor = ControlPaint.Light(orig, 0.1f);
            btn.MouseLeave += (s, e) => btn.BackColor = orig;
            return btn;
        }

        private void LoadData()
        {
            try
            {
                string query = $"SELECT {idColumnName}, {nameColumnName} FROM {tableName} ORDER BY {nameColumnName}";
                DataTable dt = dbHelper.ExecuteQuery(query);

                if (dt != null && dt.Rows.Count > 0)
                {
                    dgvData.DataSource = dt;
                    dgvData.Columns[0].HeaderText = "شناسه";
                    dgvData.Columns[1].HeaderText = "نام";
                    dgvData.Columns[0].Width = 100;
                    lblCount.Text = $"تعداد رکوردها: {dt.Rows.Count}";
                }
                else
                {
                    dgvData.DataSource = null;
                    lblCount.Text = "تعداد رکوردها: 0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در بارگذاری اطلاعات:\n\n{ex.Message}",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                string searchTerm = txtSearch.Text.Trim();
                if (string.IsNullOrEmpty(searchTerm))
                {
                    LoadData();
                    return;
                }

                string query = $"SELECT {idColumnName}, {nameColumnName} FROM {tableName} WHERE {nameColumnName} LIKE ? ORDER BY {nameColumnName}";
                OleDbParameter[] parameters = new OleDbParameter[]
                {
                    new OleDbParameter("?", "%" + searchTerm + "%")
                };
                DataTable dt = dbHelper.ExecuteQuery(query, parameters);

                if (dt != null && dt.Rows.Count > 0)
                {
                    dgvData.DataSource = dt;
                    dgvData.Columns[0].HeaderText = "شناسه";
                    dgvData.Columns[1].HeaderText = "نام";
                    dgvData.Columns[0].Width = 100;
                    lblCount.Text = $"تعداد یافت شده: {dt.Rows.Count}";
                }
                else
                {
                    dgvData.DataSource = null;
                    lblCount.Text = "موردی یافت نشد";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در جستجو:\n\n{ex.Message}",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            try
            {
                string newName = txtNewItem.Text.Trim();
                if (string.IsNullOrEmpty(newName))
                {
                    MessageBox.Show("⚠️ لطفاً نام مورد جدید را وارد کنید.",
                        "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Check if already exists
                string checkQuery = $"SELECT COUNT(*) FROM {tableName} WHERE {nameColumnName} = ?";
                OleDbParameter[] checkParams = new OleDbParameter[] { new OleDbParameter("?", newName) };
                DataTable checkDt = dbHelper.ExecuteQuery(checkQuery, checkParams);

                if (checkDt != null && Convert.ToInt32(checkDt.Rows[0][0]) > 0)
                {
                    MessageBox.Show("⚠️ این مورد قبلاً وجود دارد.",
                        "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string insertQuery = $"INSERT INTO {tableName} ({nameColumnName}) VALUES (?)";
                OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", newName) };
                int result = dbHelper.ExecuteNonQuery(insertQuery, parameters);

                if (result > 0)
                {
                    MessageBox.Show("✅ مورد جدید با موفقیت اضافه شد!",
                        "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    txtNewItem.Clear();
                    LoadData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در افزودن:\n\n{ex.Message}",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvData.SelectedRows.Count == 0)
                {
                    MessageBox.Show("⚠️ لطفاً یک ردیف را انتخاب کنید.",
                        "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string newName = txtNewItem.Text.Trim();
                if (string.IsNullOrEmpty(newName))
                {
                    MessageBox.Show("⚠️ لطفاً نام جدید را وارد کنید.",
                        "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int selectedId = Convert.ToInt32(dgvData.SelectedRows[0].Cells[0].Value);
                string oldName = dgvData.SelectedRows[0].Cells[1].Value.ToString();

                if (MessageBox.Show($"⚠️ آیا مطمئن هستید که می‌خواهید:\n\n'{oldName}'\n\nرا به:\n\n'{newName}'\n\nتغییر دهید؟",
                    "تأیید ویرایش", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string updateQuery = $"UPDATE {tableName} SET {nameColumnName} = ? WHERE {idColumnName} = ?";
                    OleDbParameter[] parameters = new OleDbParameter[]
                    {
                        new OleDbParameter("?", newName),
                        new OleDbParameter("?", selectedId)
                    };
                    int result = dbHelper.ExecuteNonQuery(updateQuery, parameters);

                    if (result > 0)
                    {
                        MessageBox.Show("✅ ویرایش با موفقیت انجام شد!",
                            "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtNewItem.Clear();
                        LoadData();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در ویرایش:\n\n{ex.Message}",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgvData.SelectedRows.Count == 0)
                {
                    MessageBox.Show("⚠️ لطفاً یک ردیف را انتخاب کنید.",
                        "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int selectedId = Convert.ToInt32(dgvData.SelectedRows[0].Cells[0].Value);
                string selectedName = dgvData.SelectedRows[0].Cells[1].Value.ToString();

                if (MessageBox.Show($"⚠️ آیا مطمئن هستید که می‌خواهید:\n\n'{selectedName}'\n\nرا حذف کنید؟\n\n⚠️ توجه: اگر این مورد در جداول دیگر استفاده شده باشد، حذف انجام نمی‌شود.",
                    "تأیید حذف", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    string deleteQuery = $"DELETE FROM {tableName} WHERE {idColumnName} = ?";
                    OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", selectedId) };
                    int result = dbHelper.ExecuteNonQuery(deleteQuery, parameters);

                    if (result > 0)
                    {
                        MessageBox.Show("✅ حذف با موفقیت انجام شد!",
                            "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        txtNewItem.Clear();
                        LoadData();
                    }
                    else
                    {
                        MessageBox.Show("❌ حذف انجام نشد. احتمالاً این مورد در جای دیگری استفاده شده است.",
                            "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در حذف:\n\n{ex.Message}\n\nاحتمالاً این مورد در جداول دیگر استفاده شده است.",
                    "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
    }
}