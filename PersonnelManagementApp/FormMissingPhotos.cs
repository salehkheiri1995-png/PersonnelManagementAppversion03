using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Data.OleDb;

namespace PersonnelManagementApp
{
    public partial class FormMissingPhotos : Form
    {
        private readonly DbHelper dbHelper;
        private DataGridView dgvMissingPhotos = null!;
        private Label lblTitle = null!;
        private Label lblCount = null!;
        private Button btnExportExcel = null!;
        private Button btnRefresh = null!;
        private Button btnClose = null!;
        private TableLayoutPanel mainLayout = null!;
        private Panel buttonPanel = null!; // تغییر از FlowLayoutPanel به Panel
        private DataTable currentData = null!;

        // رنگ‌های مدرن
        private readonly Color PrimaryColor = Color.FromArgb(33, 150, 243);
        private readonly Color AccentColor = Color.FromArgb(76, 175, 80);
        private readonly Color DangerColor = Color.FromArgb(244, 67, 54);
        private readonly Color WarningColor = Color.FromArgb(255, 152, 0);
        private readonly Color BackgroundColor = Color.FromArgb(240, 248, 255);
        private readonly Color HeaderColor = Color.FromArgb(33, 150, 243);

        public FormMissingPhotos()
        {
            dbHelper = new DbHelper();
            InitializeComponent();
            FontSettings.ApplyFontToForm(this);
            LoadMissingPhotos();
        }

        private void InitializeComponent()
        {
            this.Text = "📸 پرسنل بدون عکس";
            this.Size = new Size(1400, 800);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.RightToLeft = RightToLeft.Yes;
            this.BackColor = BackgroundColor;
            this.WindowState = FormWindowState.Maximized;
            this.MinimumSize = new Size(1000, 600);

            // ایجاد ساختار اصلی صفحه با TableLayoutPanel
            mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.ColumnCount = 1;
            mainLayout.RowCount = 3;
            mainLayout.Padding = new Padding(10);
            // ردیف اول: هدر (ثابت)
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 110F));
            // ردیف دوم: لیست (پر کردن فضا)
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            // ردیف سوم: دکمه‌ها (ثابت)
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 80F));
            this.Controls.Add(mainLayout);

            // ========== 1. پنل هدر (ردیف اول) ==========
            Panel headerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = HeaderColor,
                Margin = new Padding(0, 0, 0, 10)
            };

            lblTitle = new Label
            {
                Text = "📸 لیست پرسنل بدون عکس",
                Font = new Font(FontSettings.TitleFont?.FontFamily ?? FontFamily.GenericSansSerif, 18, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = true,
                Location = new Point(20, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            headerPanel.Controls.Add(lblTitle);

            lblCount = new Label
            {
                Text = "🔍 در حال بارگذاری...",
                Font = FontSettings.SubtitleFont,
                ForeColor = Color.FromArgb(230, 240, 255),
                AutoSize = true,
                Location = new Point(20, 65),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            headerPanel.Controls.Add(lblCount);
            
            mainLayout.Controls.Add(headerPanel, 0, 0);

            // ========== 2. لیست داده‌ها (ردیف دوم) ==========
            dgvMissingPhotos = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, 
                ReadOnly = false,
                AllowUserToDeleteRows = false,
                RightToLeft = RightToLeft.Yes,
                BackgroundColor = Color.White,
                EnableHeadersVisualStyles = false,
                AllowUserToAddRows = false,
                ColumnHeadersHeight = 50,
                RowTemplate = { Height = 45 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 0, 0, 10)
            };

            dgvMissingPhotos.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
            dgvMissingPhotos.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMissingPhotos.ColumnHeadersDefaultCellStyle.Font = new Font(FontSettings.SubtitleFont.FontFamily, 11, FontStyle.Bold);
            dgvMissingPhotos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            
            dgvMissingPhotos.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgvMissingPhotos.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvMissingPhotos.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            mainLayout.Controls.Add(dgvMissingPhotos, 0, 1);

            // ========== 3. دکمه‌ها (ردیف سوم) ==========
            buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(0)
            };

            int buttonWidth = 160;
            int buttonHeight = 45;
            int spacing = 15;

            btnExportExcel = CreateStyledButton("📊 خروجی اکسل", AccentColor, buttonWidth, buttonHeight);
            btnExportExcel.Click += BtnExportExcel_Click;

            btnRefresh = CreateStyledButton("🔄 بروزرسانی", PrimaryColor, buttonWidth, buttonHeight);
            btnRefresh.Click += BtnRefresh_Click;

            btnClose = CreateStyledButton("❌ بستن", DangerColor, buttonWidth, buttonHeight);
            btnClose.Click += (s, e) => this.Close();

            // محاسبه موقعیت مرکز برای دکمه‌ها
            buttonPanel.Resize += (s, e) =>
            {
                int totalWidth = (buttonWidth * 3) + (spacing * 2);
                int startX = (buttonPanel.Width - totalWidth) / 2;
                int y = (buttonPanel.Height - buttonHeight) / 2;

                btnExportExcel.Location = new Point(startX, y);
                btnRefresh.Location = new Point(startX + buttonWidth + spacing, y);
                btnClose.Location = new Point(startX + (buttonWidth + spacing) * 2, y);
            };

            buttonPanel.Controls.Add(btnExportExcel);
            buttonPanel.Controls.Add(btnRefresh);
            buttonPanel.Controls.Add(btnClose);

            mainLayout.Controls.Add(buttonPanel, 0, 2);
        }

        private Button CreateStyledButton(string text, Color backColor, int width, int height)
        {
            Button btn = new Button
            {
                Text = text,
                Size = new Size(width, height),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = FontSettings.ButtonFont
            };
            btn.FlatAppearance.BorderSize = 0;

            Color originalColor = backColor;
            btn.MouseEnter += (s, e) => btn.BackColor = ControlPaint.Light(originalColor, 0.1f);
            btn.MouseLeave += (s, e) => btn.BackColor = originalColor;

            return btn;
        }

        private void LoadMissingPhotos()
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                // کوئری برای دریافت اطلاعات کامل پرسنل و مشخصات اداری
                string query = @"SELECT Personnel.PersonnelID, Personnel.FirstName, Personnel.LastName,
                               Personnel.PersonnelNumber, Personnel.NationalID, Personnel.MobileNumber,
                               Personnel.HireDate,
                               OperationDepartments.DeptName,
                               Districts.DistrictName,
                               PostsNames.PostName
                               FROM (((Personnel
                               LEFT JOIN OperationDepartments ON Personnel.DeptID = OperationDepartments.DeptID)
                               LEFT JOIN Districts ON Personnel.DistrictID = Districts.DistrictID)
                               LEFT JOIN PostsNames ON Personnel.PostNameID = PostsNames.PostNameID)
                               ORDER BY Personnel.LastName, Personnel.FirstName";

                DataTable? dt = dbHelper.ExecuteQuery(query);
                if (dt == null || dt.Rows.Count == 0)
                {
                    dgvMissingPhotos.Columns.Clear();
                    dgvMissingPhotos.Rows.Clear();
                    lblCount.Text = "ℹ️ هیچ داده‌ای یافت نشد.";
                    return;
                }

                // فیلتر کردن پرسنل‌هایی که عکس ندارند
                DataTable missing = dt.Clone();
                foreach (DataRow row in dt.Rows)
                {
                    string nationalId = row["NationalID"]?.ToString() ?? string.Empty;

                    // اگر کد ملی خالی است یا عکس ندارد
                    if (string.IsNullOrWhiteSpace(nationalId) || !ImageHelper.ImageExists(nationalId))
                    {
                        missing.ImportRow(row);
                    }
                }

                currentData = missing;

                if (currentData.Rows.Count > 0)
                {
                    SetupDataGridView();
                    PopulateDataGridView();
                    lblCount.Text = $"📊 تعداد پرسنل بدون عکس: {currentData.Rows.Count} نفر";
                }
                else
                {
                    dgvMissingPhotos.Columns.Clear();
                    dgvMissingPhotos.Rows.Clear();
                    lblCount.Text = "✅ همه پرسنل دارای عکس هستند!";
                    MessageBox.Show("✅ تمام پرسنل دارای عکس پرسنلی می‌باشند.", "اطلاعات", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در بارگذاری اطلاعات:\n\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void SetupDataGridView()
        {
            dgvMissingPhotos.Columns.Clear();
            dgvMissingPhotos.AutoGenerateColumns = false;

            // 1. ستون پنهان (ID)
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "PersonnelID", 
                DataPropertyName = "PersonnelID",
                Visible = false 
            });

            // 2. ردیف
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "RowNumber", 
                HeaderText = "ردیف", 
                Width = 60,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                ReadOnly = true
            });

            // 3. نام
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "FirstName", 
                DataPropertyName = "FirstName",
                HeaderText = "نام", 
                FillWeight = 15,
                ReadOnly = true
            });

            // 4. نام خانوادگی
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "LastName", 
                DataPropertyName = "LastName",
                HeaderText = "نام خانوادگی", 
                FillWeight = 20,
                ReadOnly = true
            });

            // 5. شماره پرسنلی
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "PersonnelNumber", 
                DataPropertyName = "PersonnelNumber",
                HeaderText = "ش.پرسنلی", 
                Width = 90,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                ReadOnly = true
            });

            // 6. کد ملی
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "NationalID", 
                DataPropertyName = "NationalID",
                HeaderText = "کد ملی", 
                Width = 110,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                ReadOnly = true
            });

            // 7. اداره
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "DeptName", 
                DataPropertyName = "DeptName",
                HeaderText = "اداره", 
                FillWeight = 20,
                ReadOnly = true
            });

            // 8. ناحیه
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "DistrictName", 
                DataPropertyName = "DistrictName",
                HeaderText = "ناحیه", 
                FillWeight = 15,
                ReadOnly = true
            });

            // 9. پست
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "PostName", 
                DataPropertyName = "PostName",
                HeaderText = "پست", 
                FillWeight = 20,
                ReadOnly = true
            });

            // 10. موبایل
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn 
            { 
                Name = "MobileNumber", 
                DataPropertyName = "MobileNumber",
                HeaderText = "موبایل", 
                Width = 110,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                ReadOnly = true
            });

            // 11. دکمه ویرایش
            DataGridViewButtonColumn editColumn = new DataGridViewButtonColumn
            {
                Name = "Edit",
                HeaderText = "ویرایش",
                Text = "✏️ ویرایش",
                UseColumnTextForButtonValue = true,
                Width = 90,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                DefaultCellStyle = new DataGridViewCellStyle 
                { 
                    BackColor = Color.FromArgb(40, 167, 69), 
                    ForeColor = Color.White,
                    SelectionBackColor = Color.FromArgb(30, 140, 50),
                    SelectionForeColor = Color.White
                }
            };
            dgvMissingPhotos.Columns.Add(editColumn);

            dgvMissingPhotos.CellClick += DgvMissingPhotos_CellClick;
        }

        private void PopulateDataGridView()
        {
            dgvMissingPhotos.Rows.Clear();
            
            if (currentData == null || currentData.Rows.Count == 0)
                return;

            int rowNumber = 1;
            foreach (DataRow dataRow in currentData.Rows)
            {
                int rowIndex = dgvMissingPhotos.Rows.Add();
                DataGridViewRow gridRow = dgvMissingPhotos.Rows[rowIndex];

                gridRow.Cells["PersonnelID"].Value = dataRow["PersonnelID"];
                gridRow.Cells["RowNumber"].Value = rowNumber++;
                gridRow.Cells["FirstName"].Value = dataRow["FirstName"]?.ToString() ?? "";
                gridRow.Cells["LastName"].Value = dataRow["LastName"]?.ToString() ?? "";
                gridRow.Cells["PersonnelNumber"].Value = dataRow["PersonnelNumber"]?.ToString() ?? "";
                gridRow.Cells["NationalID"].Value = dataRow["NationalID"]?.ToString() ?? "";
                gridRow.Cells["DeptName"].Value = dataRow["DeptName"]?.ToString() ?? "";
                gridRow.Cells["DistrictName"].Value = dataRow["DistrictName"]?.ToString() ?? "";
                gridRow.Cells["PostName"].Value = dataRow["PostName"]?.ToString() ?? "";
                gridRow.Cells["MobileNumber"].Value = dataRow["MobileNumber"]?.ToString() ?? "";
            }
        }

        private void DgvMissingPhotos_CellClick(object? sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            try
            {
                // بررسی اینکه آیا روی دکمه ویرایش کلیک شده
                if (e.ColumnIndex == dgvMissingPhotos.Columns["Edit"].Index)
                {
                    var cellValue = dgvMissingPhotos.Rows[e.RowIndex].Cells["PersonnelID"].Value;
                    if (cellValue != null)
                    {
                        int personnelID = Convert.ToInt32(cellValue);
                        OpenEditForm(personnelID);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenEditForm(int personnelID)
        {
            try
            {
                FormPersonnelEdit editForm = new FormPersonnelEdit();
                editForm.txtPersonnelID.Text = personnelID.ToString();
                editForm.BtnLoad_Click(null, EventArgs.Empty);

                if (editForm.ShowDialog(this) == DialogResult.OK)
                {
                    LoadMissingPhotos();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در باز کردن فرم ویرایش:\n\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnRefresh_Click(object? sender, EventArgs e)
        {
            LoadMissingPhotos();
        }

        private void BtnExportExcel_Click(object? sender, EventArgs e)
        {
            try
            {
                if (currentData == null || currentData.Rows.Count == 0)
                {
                    MessageBox.Show("❌ داده‌ای برای خروجی وجود ندارد.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = $"PersonnelWithoutPhoto_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                    Title = "ذخیره فایل اکسل"
                };

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    this.Cursor = Cursors.WaitCursor;

                    using (var workbook = new XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("پرسنل بدون عکس");

                        worksheet.Cell(1, 1).Value = "ردیف";
                        worksheet.Cell(1, 2).Value = "نام";
                        worksheet.Cell(1, 3).Value = "نام خانوادگی";
                        worksheet.Cell(1, 4).Value = "شماره پرسنلی";
                        worksheet.Cell(1, 5).Value = "کد ملی";
                        worksheet.Cell(1, 6).Value = "اداره";
                        worksheet.Cell(1, 7).Value = "ناحیه";
                        worksheet.Cell(1, 8).Value = "پست";
                        worksheet.Cell(1, 9).Value = "تلفن همراه";

                        var headerRange = worksheet.Range(1, 1, 1, 9);
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 102, 204);
                        headerRange.Style.Font.FontColor = XLColor.White;
                        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        int excelRow = 2;
                        foreach (DataRow row in currentData.Rows)
                        {
                            worksheet.Cell(excelRow, 1).Value = excelRow - 1;
                            worksheet.Cell(excelRow, 2).Value = row["FirstName"]?.ToString();
                            worksheet.Cell(excelRow, 3).Value = row["LastName"]?.ToString();
                            worksheet.Cell(excelRow, 4).Value = row["PersonnelNumber"]?.ToString();
                            worksheet.Cell(excelRow, 5).Value = row["NationalID"]?.ToString();
                            worksheet.Cell(excelRow, 6).Value = row["DeptName"]?.ToString();
                            worksheet.Cell(excelRow, 7).Value = row["DistrictName"]?.ToString();
                            worksheet.Cell(excelRow, 8).Value = row["PostName"]?.ToString();
                            worksheet.Cell(excelRow, 9).Value = row["MobileNumber"]?.ToString();
                            excelRow++;
                        }

                        worksheet.Columns().AdjustToContents();
                        workbook.SaveAs(sfd.FileName);
                    }

                    MessageBox.Show("✅ فایل اکسل ذخیره شد.", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
    }
}