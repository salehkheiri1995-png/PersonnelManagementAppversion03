using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace PersonnelManagementApp
{
    public partial class FormMissingPhotos : Form
    {
        private readonly DbHelper dbHelper;
        private DataGridView dgvMissingPhotos;
        private Label lblTitle;
        private Label lblCount;
        private Button btnExportExcel;
        private Button btnRefresh;
        private Button btnClose;
        private Panel panelHeader;
        private Panel panelButtons;
        private DataTable currentData;

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
            this.MinimumSize = new Size(1200, 600);

            // ========== پنل هدر ==========
            panelHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 100,
                BackColor = HeaderColor
            };

            lblTitle = new Label
            {
                Text = "📸 لیست پرسنل بدون عکس",
                Font = new Font(FontSettings.TitleFont?.FontFamily ?? FontFamily.GenericSansSerif, 18, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(20, 15),
                Size = new Size(600, 40),
                TextAlign = ContentAlignment.MiddleRight
            };
            panelHeader.Controls.Add(lblTitle);

            lblCount = new Label
            {
                Text = "🔍 در حال بارگذاری...",
                Font = FontSettings.SubtitleFont,
                ForeColor = Color.FromArgb(230, 240, 255),
                Location = new Point(20, 55),
                Size = new Size(600, 30),
                TextAlign = ContentAlignment.MiddleRight
            };
            panelHeader.Controls.Add(lblCount);

            this.Controls.Add(panelHeader);

            // ========== DataGridView ==========
            dgvMissingPhotos = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                ReadOnly = false,
                RightToLeft = RightToLeft.Yes,
                BackgroundColor = Color.White,
                EnableHeadersVisualStyles = false,
                AllowUserToAddRows = false,
                ColumnHeadersHeight = 45,
                RowTemplate = { Height = 40 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };

            dgvMissingPhotos.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
            dgvMissingPhotos.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvMissingPhotos.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
            dgvMissingPhotos.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvMissingPhotos.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgvMissingPhotos.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            this.Controls.Add(dgvMissingPhotos);

            // ========== پنل دکمه‌ها ==========
            panelButtons = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 80,
                BackColor = Color.White,
                Padding = new Padding(20)
            };

            int buttonWidth = 180;
            int buttonHeight = 45;
            int buttonSpacing = 15;
            int startX = (this.Width - (3 * buttonWidth + 2 * buttonSpacing)) / 2;

            btnExportExcel = CreateStyledButton("📊 خروجی اکسل", AccentColor, buttonWidth, buttonHeight);
            btnExportExcel.Location = new Point(startX, 17);
            btnExportExcel.Click += BtnExportExcel_Click;
            panelButtons.Controls.Add(btnExportExcel);

            btnRefresh = CreateStyledButton("🔄 بروزرسانی", PrimaryColor, buttonWidth, buttonHeight);
            btnRefresh.Location = new Point(startX + buttonWidth + buttonSpacing, 17);
            btnRefresh.Click += BtnRefresh_Click;
            panelButtons.Controls.Add(btnRefresh);

            btnClose = CreateStyledButton("❌ بستن", DangerColor, buttonWidth, buttonHeight);
            btnClose.Location = new Point(startX + 2 * (buttonWidth + buttonSpacing), 17);
            btnClose.Click += (s, e) => this.Close();
            panelButtons.Controls.Add(btnClose);

            this.Controls.Add(panelButtons);
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

                // 1) فقط اطلاعات پایه از دیتابیس (بدون JOIN سنگین، برای جلوگیری از خطاهای Access)
                // 2) عکس‌ها را از روی فولدر تنظیمات و نام فایل = کد ملی (NationalID) بررسی می‌کنیم
                //    (طبق ImageHelper: مسیر از AppSettings.PhotosFolder و نام فایل {NationalID}.jpg)
                // 3) هر پرسنلی که عکس ندارد در لیست نمایش داده می‌شود.
                string query = @"SELECT Personnel.PersonnelID, Personnel.FirstName, Personnel.LastName,
                               Personnel.PersonnelNumber, Personnel.NationalID, Personnel.MobileNumber,
                               Personnel.HireDate
                               FROM Personnel
                               ORDER BY Personnel.LastName, Personnel.FirstName";

                DataTable dt = dbHelper.ExecuteQuery(query);
                if (dt == null || dt.Rows.Count == 0)
                {
                    dgvMissingPhotos.Columns.Clear();
                    dgvMissingPhotos.Rows.Clear();
                    lblCount.Text = "ℹ️ هیچ داده‌ای یافت نشد.";
                    return;
                }

                // بررسی مسیر عکس‌ها
                string imagesFolder = ImageHelper.ImagesFolderPath;
                if (string.IsNullOrWhiteSpace(imagesFolder) || !Directory.Exists(imagesFolder))
                {
                    MessageBox.Show(
                        $"❌ مسیر پوشه عکس‌ها معتبر نیست یا وجود ندارد:\n\n{imagesFolder}\n\nلطفاً از تنظیمات، مسیر پوشه عکس‌ها را درست کنید.",
                        "خطا",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // فیلتر کردن پرسنل‌هایی که عکس ندارند
                DataTable missing = dt.Clone();
                foreach (DataRow row in dt.Rows)
                {
                    string nationalId = row["NationalID"]?.ToString() ?? string.Empty;

                    // اگر کد ملی خالی است، به عنوان "بدون عکس" گزارش شود
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

            // ستون‌های پنهان
            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "PersonnelID",
                HeaderText = "ID",
                Visible = false
            });

            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "RowNumber",
                HeaderText = "ردیف",
                Width = 60
            });

            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "FirstName",
                HeaderText = "نام",
                Width = 120
            });

            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "LastName",
                HeaderText = "نام‌خانوادگی",
                Width = 140
            });

            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "PersonnelNumber",
                HeaderText = "شماره پرسنلی",
                Width = 120
            });

            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "NationalID",
                HeaderText = "کد ملی",
                Width = 120
            });

            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "HireDate",
                HeaderText = "تاریخ استخدام",
                Width = 120
            });

            dgvMissingPhotos.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "MobileNumber",
                HeaderText = "تلفن همراه",
                Width = 120
            });

            // دکمه ویرایش
            DataGridViewButtonColumn editColumn = new DataGridViewButtonColumn
            {
                Name = "Edit",
                HeaderText = "✏️ ویرایش",
                Text = "ویرایش",
                UseColumnTextForButtonValue = true,
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(40, 167, 69),
                    ForeColor = Color.White,
                    Font = FontSettings.ButtonFont,
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    Padding = new Padding(5)
                }
            };
            dgvMissingPhotos.Columns.Add(editColumn);

            // دکمه حذف
            DataGridViewButtonColumn deleteColumn = new DataGridViewButtonColumn
            {
                Name = "Delete",
                HeaderText = "🗑️ حذف",
                Text = "حذف",
                UseColumnTextForButtonValue = true,
                Width = 100,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(220, 53, 69),
                    ForeColor = Color.White,
                    Font = FontSettings.ButtonFont,
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    Padding = new Padding(5)
                }
            };
            dgvMissingPhotos.Columns.Add(deleteColumn);

            dgvMissingPhotos.CellClick += DgvMissingPhotos_CellClick;
        }

        private void PopulateDataGridView()
        {
            dgvMissingPhotos.Rows.Clear();

            int rowNumber = 1;
            foreach (DataRow row in currentData.Rows)
            {
                string hireDate = row["HireDate"] != DBNull.Value
                    ? Convert.ToDateTime(row["HireDate"]).ToString("yyyy/MM/dd")
                    : "";

                dgvMissingPhotos.Rows.Add(
                    row["PersonnelID"],
                    rowNumber++,
                    row["FirstName"],
                    row["LastName"],
                    row["PersonnelNumber"],
                    row["NationalID"],
                    hireDate,
                    row["MobileNumber"],
                    "ویرایش",
                    "حذف"
                );
            }
        }

        private void DgvMissingPhotos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            try
            {
                int personnelID = Convert.ToInt32(dgvMissingPhotos.Rows[e.RowIndex].Cells["PersonnelID"].Value);

                if (e.ColumnIndex == dgvMissingPhotos.Columns["Edit"].Index)
                {
                    OpenEditForm(personnelID);
                }
                else if (e.ColumnIndex == dgvMissingPhotos.Columns["Delete"].Index)
                {
                    DeletePersonnel(personnelID, e.RowIndex);
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

        private void DeletePersonnel(int personnelID, int rowIndex)
        {
            try
            {
                string personnelName = $"{dgvMissingPhotos.Rows[rowIndex].Cells["FirstName"].Value} {dgvMissingPhotos.Rows[rowIndex].Cells["LastName"].Value}";

                DialogResult result = MessageBox.Show(
                    $"❓ آیا مطمئن هستید که می‌خواهید '{personnelName}' را حذف کنید؟\n\n⚠️ این عملیات قابل بازگشت نیست!",
                    "تایید حذف",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    string query = $"DELETE FROM Personnel WHERE PersonnelID = {personnelID}";
                    int affectedRows = dbHelper.ExecuteNonQuery(query);

                    if (affectedRows > 0)
                    {
                        MessageBox.Show("✅ پرسنل با موفقیت حذف شد.", "موفق", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        dgvMissingPhotos.Rows.RemoveAt(rowIndex);
                        UpdateRowNumbers();
                        lblCount.Text = $"📊 تعداد پرسنل بدون عکس: {dgvMissingPhotos.Rows.Count} نفر";

                        if (dgvMissingPhotos.Rows.Count == 0)
                        {
                            lblCount.Text = "✅ همه پرسنل دارای عکس هستند!";
                            MessageBox.Show("✅ تمام پرسنل دارای عکس پرسنلی می‌باشند.", "اطلاعات", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                    else
                    {
                        MessageBox.Show("❌ خطا در حذف پرسنل.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در حذف پرسنل:\n\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateRowNumbers()
        {
            for (int i = 0; i < dgvMissingPhotos.Rows.Count; i++)
            {
                dgvMissingPhotos.Rows[i].Cells["RowNumber"].Value = i + 1;
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            LoadMissingPhotos();
        }

        private void BtnExportExcel_Click(object sender, EventArgs e)
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
                        worksheet.Cell(1, 3).Value = "نام‌خانوادگی";
                        worksheet.Cell(1, 4).Value = "شماره پرسنلی";
                        worksheet.Cell(1, 5).Value = "کد ملی";
                        worksheet.Cell(1, 6).Value = "تاریخ استخدام";
                        worksheet.Cell(1, 7).Value = "تلفن همراه";
                        worksheet.Cell(1, 8).Value = "مسیر پوشه عکس‌ها";

                        var headerRange = worksheet.Range(1, 1, 1, 8);
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 102, 204);
                        headerRange.Style.Font.FontColor = XLColor.White;
                        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        string imagesFolder = ImageHelper.ImagesFolderPath;

                        int rowNumber = 1;
                        int excelRow = 2;
                        foreach (DataRow row in currentData.Rows)
                        {
                            worksheet.Cell(excelRow, 1).Value = rowNumber++;
                            worksheet.Cell(excelRow, 2).Value = row["FirstName"]?.ToString() ?? "";
                            worksheet.Cell(excelRow, 3).Value = row["LastName"]?.ToString() ?? "";
                            worksheet.Cell(excelRow, 4).Value = row["PersonnelNumber"]?.ToString() ?? "";
                            worksheet.Cell(excelRow, 5).Value = row["NationalID"]?.ToString() ?? "";

                            string hireDate = row["HireDate"] != DBNull.Value
                                ? Convert.ToDateTime(row["HireDate"]).ToString("yyyy/MM/dd")
                                : "";
                            worksheet.Cell(excelRow, 6).Value = hireDate;

                            worksheet.Cell(excelRow, 7).Value = row["MobileNumber"]?.ToString() ?? "";
                            worksheet.Cell(excelRow, 8).Value = imagesFolder;

                            if (excelRow % 2 == 0)
                            {
                                worksheet.Range(excelRow, 1, excelRow, 8).Style.Fill.BackgroundColor = XLColor.FromArgb(240, 248, 255);
                            }

                            excelRow++;
                        }

                        worksheet.Columns().AdjustToContents();
                        worksheet.RightToLeft = true;

                        workbook.SaveAs(sfd.FileName);
                    }

                    MessageBox.Show($"✅ فایل اکسل با موفقیت ذخیره شد:\n\n{sfd.FileName}", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    DialogResult openResult = MessageBox.Show("آیا می‌خواهید فایل را باز کنید؟", "باز کردن فایل", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (openResult == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = sfd.FileName,
                            UseShellExecute = true
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در ایجاد فایل اکسل:\n\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
    }
}
