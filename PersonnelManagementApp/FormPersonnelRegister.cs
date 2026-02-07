using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    public partial class FormPersonnelRegister : Form
    {
        private DbHelper db = new DbHelper();

        // ComboBox برای فیلدهای وابسته و غیروابسته
        private ComboBox cbProvince, cbCity, cbAffair, cbDept, cbDistrict, cbPostName;
        private ComboBox cbVoltage, cbWorkShift, cbGender, cbContractType, cbJobLevel, cbCompany, cbDegree, cbDegreeField, cbStatus;
        private ComboBox cbMainJobTitle, cbCurrentActivity;

        // TextBox برای فیلدهای دستی
        private TextBox txtFirstName, txtLastName, txtFatherName, txtPersonnelNumber, txtNationalID, txtMobileNumber;
        private TextBox txtDescription;

        // DateTimePicker برای تاریخ‌ها
        private DateTimePicker dtpBirthDate, dtpHireDate, dtpStartDateOperation;

        // CheckBox برای مغایرت
        private CheckBox chkInconsistency;

        // ⭐ فیلدهای جدید برای مدیریت عکس
        private PictureBox pbPhoto;
        private string selectedPhotoPath = string.Empty;

        public FormPersonnelRegister()
        {
            InitializeComponent();

            // اعمال فونت‌های تنظیم‌شده
            FontSettings.ApplyFontToForm(this);

            try
            {
                LoadProvinces();
                LoadOtherCombos();
                UpdateInconsistency();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری داده‌ها: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.Text = "ثبت پرسنل جدید";
            this.Size = new Size(600, 900);
            this.WindowState = FormWindowState.Maximized;
            this.AutoScroll = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.RightToLeft = RightToLeft.Yes;
            this.BackColor = Color.FromArgb(240, 248, 255);

            // پس‌زمینه گرادیانت
            try
            {
                using (LinearGradientBrush brush = new LinearGradientBrush(this.ClientRectangle, Color.LightBlue, Color.SkyBlue, LinearGradientMode.Vertical))
                {
                    brush.GammaCorrection = true;
                    this.BackgroundImage = new Bitmap(this.Width, this.Height);
                    using (Graphics g = Graphics.FromImage(this.BackgroundImage))
                    {
                        g.FillRectangle(brush, this.ClientRectangle);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در رسم پس‌زمینه: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            int yHeader = 20;
            int controlWidth = 300;
            int controlHeight = 40;
            int yStep = 60;
            int formWidth = this.ClientSize.Width;
            int labelWidth = 250;
            int totalControlWidth = labelWidth + controlWidth + 10;
            int margin = 20;
            int columnWidth = (formWidth - 3 * margin) / 2;

            // موقعیت پایه برای ستون سمت راست
            int baseXRight = formWidth - margin - columnWidth;
            int xControlRight = baseXRight + (columnWidth - totalControlWidth) / 2;
            int xLabelRight = xControlRight + controlWidth + 10;

            // موقعیت پایه برای ستون سمت چپ
            int baseXLeft = margin;
            int xControlLeft = baseXLeft + (columnWidth - totalControlWidth) / 2;
            int xLabelLeft = xControlLeft + controlWidth + 10;

            // سرتیتر
            Label lblHeader = new Label
            {
                Text = "ثبت پرسنل جدید",
                Location = new Point((formWidth - 400) / 2, yHeader),
                Size = new Size(400, 50),
                Font = FontSettings.TitleFont,
                ForeColor = Color.Navy,
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblHeader);
            yHeader += 60;

            // ⭐ بخش عکس - اضافه شده
            int photoBoxSize = 200;
            pbPhoto = new PictureBox
            {
                Location = new Point((formWidth - photoBoxSize) / 2, yHeader),
                Size = new Size(photoBoxSize, photoBoxSize),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White
            };
            pbPhoto.Image = ImageHelper.CreateDefaultImage(photoBoxSize, photoBoxSize);
            ApplyRoundedCorners(pbPhoto, 15);
            this.Controls.Add(pbPhoto);
            yHeader += photoBoxSize + 10;

            // دکمه‌های مدیریت عکس
            int btnPhotoWidth = 95;
            int btnPhotoSpacing = 5;
            int totalPhotoButtonWidth = (btnPhotoWidth * 2) + btnPhotoSpacing;
            int xPhotoButtonStart = (formWidth - totalPhotoButtonWidth) / 2;

            Button btnSelectPhoto = new Button
            {
                Text = "انتخاب عکس",
                Location = new Point(xPhotoButtonStart, yHeader),
                Size = new Size(btnPhotoWidth, 35),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightBlue,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnSelectPhoto, 10);
            btnSelectPhoto.Click += BtnSelectPhoto_Click;
            this.Controls.Add(btnSelectPhoto);

            Button btnRemovePhoto = new Button
            {
                Text = "حذف عکس",
                Location = new Point(xPhotoButtonStart + btnPhotoWidth + btnPhotoSpacing, yHeader),
                Size = new Size(btnPhotoWidth, 35),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightCoral,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnRemovePhoto, 10);
            btnRemovePhoto.Click += BtnRemovePhoto_Click;
            this.Controls.Add(btnRemovePhoto);
            yHeader += 45;

            // تنظیم مجدد yRight و yLeft
            int yRight = yHeader + 10;
            int yLeft = yHeader + 10;

            // *** ستون سمت راست: فیلدهای 1 تا 14 ***
            Label lblProvince = new Label { Text = "استان:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbProvince = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cbProvince.SelectedIndexChanged += CbProvince_SelectedIndexChanged;
            ApplyRoundedCorners(cbProvince, 10);
            this.Controls.Add(lblProvince);
            this.Controls.Add(cbProvince);
            yRight += yStep;

            Label lblCity = new Label { Text = "شهر:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbCity = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbCity, 10);
            this.Controls.Add(lblCity);
            this.Controls.Add(cbCity);
            yRight += yStep;

            Label lblAffair = new Label { Text = "امور انتقال:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbAffair = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cbAffair.SelectedIndexChanged += CbAffair_SelectedIndexChanged;
            ApplyRoundedCorners(cbAffair, 10);
            this.Controls.Add(lblAffair);
            this.Controls.Add(cbAffair);
            yRight += yStep;

            Label lblDept = new Label { Text = "اداره بهره‌برداری:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbDept = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cbDept.SelectedIndexChanged += CbDept_SelectedIndexChanged;
            ApplyRoundedCorners(cbDept, 10);
            this.Controls.Add(lblDept);
            this.Controls.Add(cbDept);
            yRight += yStep;

            Label lblDistrict = new Label { Text = "ناحیه بهره‌برداری:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbDistrict = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cbDistrict.SelectedIndexChanged += CbDistrict_SelectedIndexChanged;
            ApplyRoundedCorners(cbDistrict, 10);
            this.Controls.Add(lblDistrict);
            this.Controls.Add(cbDistrict);
            yRight += yStep;

            Label lblPostName = new Label { Text = "نام پست:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbPostName = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbPostName, 10);
            this.Controls.Add(lblPostName);
            this.Controls.Add(cbPostName);
            yRight += yStep;

            Label lblVoltage = new Label { Text = "سطح ولتاژ:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbVoltage = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbVoltage, 10);
            this.Controls.Add(lblVoltage);
            this.Controls.Add(cbVoltage);
            yRight += yStep;

            Label lblWorkShift = new Label { Text = "شیفت کاری:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbWorkShift = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbWorkShift, 10);
            this.Controls.Add(lblWorkShift);
            this.Controls.Add(cbWorkShift);
            yRight += yStep;

            Label lblGender = new Label { Text = "جنسیت:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbGender = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbGender, 10);
            this.Controls.Add(lblGender);
            this.Controls.Add(cbGender);
            yRight += yStep;

            Label lblFirstName = new Label { Text = "نام:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtFirstName = new TextBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtFirstName, 10);
            this.Controls.Add(lblFirstName);
            this.Controls.Add(txtFirstName);
            yRight += yStep;

            Label lblLastName = new Label { Text = "نام خانوادگی:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtLastName = new TextBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtLastName, 10);
            this.Controls.Add(lblLastName);
            this.Controls.Add(txtLastName);
            yRight += yStep;

            Label lblFatherName = new Label { Text = "نام پدر:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtFatherName = new TextBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtFatherName, 10);
            this.Controls.Add(lblFatherName);
            this.Controls.Add(txtFatherName);
            yRight += yStep;

            Label lblPersonnelNumber = new Label { Text = "شماره پرسنلی:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtPersonnelNumber = new TextBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtPersonnelNumber, 10);
            this.Controls.Add(lblPersonnelNumber);
            this.Controls.Add(txtPersonnelNumber);
            yRight += yStep;

            Label lblNationalID = new Label { Text = "کد ملی:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtNationalID = new TextBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtNationalID, 10);
            this.Controls.Add(lblNationalID);
            this.Controls.Add(txtNationalID);
            yRight += yStep;

            // *** ستون سمت چپ: فیلدهای 15 تا 28 ***
            Label lblMobileNumber = new Label { Text = "شماره موبایل:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtMobileNumber = new TextBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtMobileNumber, 10);
            this.Controls.Add(lblMobileNumber);
            this.Controls.Add(txtMobileNumber);
            yLeft += yStep;

            Label lblBirthDate = new Label { Text = "تاریخ تولد:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            dtpBirthDate = new DateTimePicker { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(dtpBirthDate, 10);
            this.Controls.Add(lblBirthDate);
            this.Controls.Add(dtpBirthDate);
            yLeft += yStep;

            Label lblHireDate = new Label { Text = "تاریخ استخدام:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            dtpHireDate = new DateTimePicker { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(dtpHireDate, 10);
            this.Controls.Add(lblHireDate);
            this.Controls.Add(dtpHireDate);
            yLeft += yStep;

            Label lblStartDateOperation = new Label { Text = "تاریخ شروع بکار:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            dtpStartDateOperation = new DateTimePicker { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(dtpStartDateOperation, 10);
            this.Controls.Add(lblStartDateOperation);
            this.Controls.Add(dtpStartDateOperation);
            yLeft += yStep;

            Label lblContractType = new Label { Text = "نوع قرارداد:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbContractType = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbContractType, 10);
            this.Controls.Add(lblContractType);
            this.Controls.Add(cbContractType);
            yLeft += yStep;

            Label lblJobLevel = new Label { Text = "سطح شغل:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbJobLevel = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbJobLevel, 10);
            this.Controls.Add(lblJobLevel);
            this.Controls.Add(cbJobLevel);
            yLeft += yStep;

            Label lblCompany = new Label { Text = "شرکت:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbCompany = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbCompany, 10);
            this.Controls.Add(lblCompany);
            this.Controls.Add(cbCompany);
            yLeft += yStep;

            Label lblDegree = new Label { Text = "مدرک تحصیلی:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbDegree = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbDegree, 10);
            this.Controls.Add(lblDegree);
            this.Controls.Add(cbDegree);
            yLeft += yStep;

            Label lblDegreeField = new Label { Text = "رشته تحصیلی:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbDegreeField = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbDegreeField, 10);
            this.Controls.Add(lblDegreeField);
            this.Controls.Add(cbDegreeField);
            yLeft += yStep;

            Label lblMainJobTitle = new Label { Text = "عنوان شغلی اصلی:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbMainJobTitle = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cbMainJobTitle.SelectedIndexChanged += (s, e) => UpdateInconsistency();
            ApplyRoundedCorners(cbMainJobTitle, 10);
            this.Controls.Add(lblMainJobTitle);
            this.Controls.Add(cbMainJobTitle);
            yLeft += yStep;

            Label lblCurrentActivity = new Label { Text = "فعالیت فعلی:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbCurrentActivity = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cbCurrentActivity.SelectedIndexChanged += (s, e) => UpdateInconsistency();
            ApplyRoundedCorners(cbCurrentActivity, 10);
            this.Controls.Add(lblCurrentActivity);
            this.Controls.Add(cbCurrentActivity);
            yLeft += yStep;

            Label lblInconsistency = new Label { Text = "مغایرت:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            chkInconsistency = new CheckBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.BodyFont, Text = "ندارد", Checked = false, Enabled = false };
            this.Controls.Add(lblInconsistency);
            this.Controls.Add(chkInconsistency);
            yLeft += yStep;

            Label lblStatus = new Label { Text = "وضعیت حضور:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cbStatus = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cbStatus, 10);
            this.Controls.Add(lblStatus);
            this.Controls.Add(cbStatus);
            yLeft += yStep;

            Label lblDescription = new Label { Text = "توضیحات:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtDescription = new TextBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, 100), Multiline = true, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtDescription, 10);
            this.Controls.Add(lblDescription);
            this.Controls.Add(txtDescription);
            yLeft += yStep + 60;

            // دکمه‌ها
            int maxY = Math.Max(yRight, yLeft) + 20;
            int buttonWidth = 150;
            int buttonSpace = 10;
            int totalButtonWidth = (buttonWidth * 2) + buttonSpace;
            int xButtonStart = (formWidth - totalButtonWidth) / 2;

            Button btnSave = new Button
            {
                Text = "ذخیره",
                Location = new Point(xButtonStart, maxY),
                Size = new Size(buttonWidth, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightGreen,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnSave, 15);
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);

            Button btnCancel = new Button
            {
                Text = "لغو",
                Location = new Point(xButtonStart + buttonWidth + buttonSpace, maxY),
                Size = new Size(buttonWidth, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightCoral,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnCancel, 15);
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);
        }

        // ⭐ Event Handlers جدید برای مدیریت عکس
        private void BtnSelectPhoto_Click(object sender, EventArgs e)
        {
            try
            {
                string photoPath = ImageHelper.OpenImageDialog();
                if (!string.IsNullOrEmpty(photoPath))
                {
                    selectedPhotoPath = photoPath;
                    Image img = Image.FromFile(photoPath);
                    ImageHelper.DrawImageInPictureBox(pbPhoto, img);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در انتخاب عکس: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnRemovePhoto_Click(object sender, EventArgs e)
        {
            try
            {
                selectedPhotoPath = string.Empty;
                pbPhoto.Image = ImageHelper.CreateDefaultImage(pbPhoto.Width, pbPhoto.Height);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در حذف عکس: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در تطبیق گوشه‌های گرد: {ex.Message}");
            }
        }

        private void LoadProvinces()
        {
            try
            {
                DataTable dt = db.GetProvinces();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbProvince.DataSource = dt;
                    cbProvince.DisplayMember = "ProvinceName";
                    cbProvince.ValueMember = "ProvinceID";
                }
                else
                {
                    MessageBox.Show("هیچ استانی برای بارگذاری یافت نشد.", "اخطار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری استان‌ها: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadOtherCombos()
        {
            try
            {
                DataTable dt;

                dt = db.GetVoltageLevels();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbVoltage.DataSource = dt;
                    cbVoltage.DisplayMember = "VoltageName";
                    cbVoltage.ValueMember = "VoltageID";
                }

                dt = db.GetWorkShifts();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbWorkShift.DataSource = dt;
                    cbWorkShift.DisplayMember = "WorkShiftName";
                    cbWorkShift.ValueMember = "WorkShiftID";
                }

                dt = db.GetGenders();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbGender.DataSource = dt;
                    cbGender.DisplayMember = "GenderName";
                    cbGender.ValueMember = "GenderID";
                }

                dt = db.GetContractTypes();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbContractType.DataSource = dt;
                    cbContractType.DisplayMember = "ContractTypeName";
                    cbContractType.ValueMember = "ContractTypeID";
                }

                dt = db.GetJobLevels();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbJobLevel.DataSource = dt;
                    cbJobLevel.DisplayMember = "JobLevelName";
                    cbJobLevel.ValueMember = "JobLevelID";
                }

                dt = db.GetCompanies();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbCompany.DataSource = dt;
                    cbCompany.DisplayMember = "CompanyName";
                    cbCompany.ValueMember = "CompanyID";
                }

                dt = db.GetDegrees();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbDegree.DataSource = dt;
                    cbDegree.DisplayMember = "DegreeName";
                    cbDegree.ValueMember = "DegreeID";
                }

                dt = db.GetDegreeFields();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbDegreeField.DataSource = dt;
                    cbDegreeField.DisplayMember = "DegreeFieldName";
                    cbDegreeField.ValueMember = "DegreeFieldID";
                }

                dt = db.GetStatusPresence();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbStatus.DataSource = dt;
                    cbStatus.DisplayMember = "StatusName";
                    cbStatus.ValueMember = "StatusID";
                }

                dt = db.GetChartAffairs();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbMainJobTitle.DataSource = dt.Copy();
                    cbMainJobTitle.DisplayMember = "ChartName";
                    cbMainJobTitle.ValueMember = "ChartID";
                }

                dt = db.GetChartAffairs1();
                if (dt != null && dt.Rows.Count > 0)
                {
                    cbCurrentActivity.DataSource = dt.Copy();
                    cbCurrentActivity.DisplayMember = "ChartName";
                    cbCurrentActivity.ValueMember = "ChartID";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری کمبوباکس‌ها: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CbProvince_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbProvince.SelectedValue != null && int.TryParse(cbProvince.SelectedValue.ToString(), out int provinceID))
                {
                    DataTable dtCity = db.GetCitiesByProvince(provinceID);
                    if (dtCity != null && dtCity.Rows.Count > 0)
                    {
                        cbCity.DataSource = dtCity;
                        cbCity.DisplayMember = "CityName";
                        cbCity.ValueMember = "CityID";
                    }
                    else
                    {
                        cbCity.DataSource = null;
                    }

                    DataTable dtAffair = db.GetAffairsByProvince(provinceID);
                    if (dtAffair != null && dtAffair.Rows.Count > 0)
                    {
                        cbAffair.DataSource = dtAffair;
                        cbAffair.DisplayMember = "AffairName";
                        cbAffair.ValueMember = "AffairID";
                    }
                    else
                    {
                        cbAffair.DataSource = null;
                    }

                    cbDept.DataSource = null;
                    cbDistrict.DataSource = null;
                    cbPostName.DataSource = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری شهر و امور: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CbAffair_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbAffair.SelectedValue != null && int.TryParse(cbAffair.SelectedValue.ToString(), out int affairID))
                {
                    DataTable dt = db.GetDeptsByAffair(affairID);
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        cbDept.DataSource = dt;
                        cbDept.DisplayMember = "DeptName";
                        cbDept.ValueMember = "DeptID";
                    }
                    else
                    {
                        cbDept.DataSource = null;
                    }

                    cbDistrict.DataSource = null;
                    cbPostName.DataSource = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری اداره بهره‌برداری: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CbDept_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbDept.SelectedValue != null && int.TryParse(cbDept.SelectedValue.ToString(), out int deptID))
                {
                    DataTable dt = db.GetDistrictsByDept(deptID);
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        cbDistrict.DataSource = dt;
                        cbDistrict.DisplayMember = "DistrictName";
                        cbDistrict.ValueMember = "DistrictID";
                    }
                    else
                    {
                        cbDistrict.DataSource = null;
                    }

                    cbPostName.DataSource = null;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری ناحیه بهره‌برداری: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CbDistrict_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (cbDistrict.SelectedValue != null && int.TryParse(cbDistrict.SelectedValue.ToString(), out int districtID))
                {
                    DataTable dt = db.GetPostNamesByDistrict(districtID);
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        cbPostName.DataSource = dt;
                        cbPostName.DisplayMember = "PostName";
                        cbPostName.ValueMember = "PostNameID";
                    }
                    else
                    {
                        cbPostName.DataSource = null;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری نام پست: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateInconsistency()
        {
            try
            {
                if (cbMainJobTitle.SelectedIndex != -1 && cbCurrentActivity.SelectedIndex != -1)
                {
                    string mainJob = cbMainJobTitle.Text?.Trim() ?? string.Empty;
                    string currentActivity = cbCurrentActivity.Text?.Trim() ?? string.Empty;
                    chkInconsistency.Checked = !mainJob.Equals(currentActivity, StringComparison.OrdinalIgnoreCase);
                    chkInconsistency.Text = chkInconsistency.Checked ? "دارد" : "ندارد";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در بروزرسانی مغایرت: {ex.Message}");
            }
        }

        private ValidationResult ValidateFormData()
        {
            if (cbProvince.SelectedIndex == -1) return new ValidationResult(false, "لطفاً استان را انتخاب کنید.");
            if (cbCity.SelectedIndex == -1) return new ValidationResult(false, "لطفاً شهر را انتخاب کنید.");
            if (cbAffair.SelectedIndex == -1) return new ValidationResult(false, "لطفاً امور انتقال را انتخاب کنید.");
            if (cbDept.SelectedIndex == -1) return new ValidationResult(false, "لطفاً اداره بهره‌برداری را انتخاب کنید.");
            if (cbDistrict.SelectedIndex == -1) return new ValidationResult(false, "لطفاً ناحیه بهره‌برداری را انتخاب کنید.");
            if (cbPostName.SelectedIndex == -1) return new ValidationResult(false, "لطفاً نام پست را انتخاب کنید.");
            if (cbVoltage.SelectedIndex == -1) return new ValidationResult(false, "لطفاً سطح ولتاژ را انتخاب کنید.");
            if (cbWorkShift.SelectedIndex == -1) return new ValidationResult(false, "لطفاً شیفت کاری را انتخاب کنید.");
            if (cbGender.SelectedIndex == -1) return new ValidationResult(false, "لطفاً جنسیت را انتخاب کنید.");
            if (cbContractType.SelectedIndex == -1) return new ValidationResult(false, "لطفاً نوع قرارداد را انتخاب کنید.");
            if (cbJobLevel.SelectedIndex == -1) return new ValidationResult(false, "لطفاً سطح شغل را انتخاب کنید.");
            if (cbCompany.SelectedIndex == -1) return new ValidationResult(false, "لطفاً شرکت را انتخاب کنید.");
            if (cbDegree.SelectedIndex == -1) return new ValidationResult(false, "لطفاً مدرک تحصیلی را انتخاب کنید.");
            if (cbDegreeField.SelectedIndex == -1) return new ValidationResult(false, "لطفاً رشته تحصیلی را انتخاب کنید.");
            if (cbMainJobTitle.SelectedIndex == -1) return new ValidationResult(false, "لطفاً عنوان شغلی اصلی را انتخاب کنید.");
            if (cbCurrentActivity.SelectedIndex == -1) return new ValidationResult(false, "لطفاً فعالیت فعلی را انتخاب کنید.");
            if (cbStatus.SelectedIndex == -1) return new ValidationResult(false, "لطفاً وضعیت حضور را انتخاب کنید.");

            if (string.IsNullOrWhiteSpace(txtFirstName.Text)) return new ValidationResult(false, "لطفاً نام را وارد کنید.");
            if (string.IsNullOrWhiteSpace(txtLastName.Text)) return new ValidationResult(false, "لطفاً نام خانوادگی را وارد کنید.");
            if (string.IsNullOrWhiteSpace(txtFatherName.Text)) return new ValidationResult(false, "لطفاً نام پدر را وارد کنید.");
            if (string.IsNullOrWhiteSpace(txtPersonnelNumber.Text)) return new ValidationResult(false, "لطفاً شماره پرسنلی را وارد کنید.");
            if (string.IsNullOrWhiteSpace(txtNationalID.Text)) return new ValidationResult(false, "لطفاً کد ملی را وارد کنید.");
            if (string.IsNullOrWhiteSpace(txtMobileNumber.Text)) return new ValidationResult(false, "لطفاً شماره موبایل را وارد کنید.");

            string nationalID = txtNationalID.Text.Trim();
            if (!System.Text.RegularExpressions.Regex.IsMatch(nationalID, @"^\d{10,11}$"))
                return new ValidationResult(false, "کد ملی باید 10 یا 11 رقم باشد.");

            string mobileNumber = txtMobileNumber.Text.Trim();
            if (!System.Text.RegularExpressions.Regex.IsMatch(mobileNumber, @"^09\d{9}$"))
                return new ValidationResult(false, "شماره موبایل باید 11 رقم و شروع با 09 باشد.");

            if (!System.Text.RegularExpressions.Regex.IsMatch(txtPersonnelNumber.Text.Trim(), @"^\d+$"))
                return new ValidationResult(false, "شماره پرسنلی باید فقط شامل اعداد باشد.");

            if (dtpBirthDate.Value >= DateTime.Now)
                return new ValidationResult(false, "تاریخ تولد نمی‌تواند در آینده باشد.");

            if (dtpBirthDate.Value.AddYears(18) > DateTime.Now)
                return new ValidationResult(false, "پرسنل باید حداقل 18 سال سن داشته باشد.");

            if (dtpHireDate.Value > DateTime.Now)
                return new ValidationResult(false, "تاریخ استخدام نمی‌تواند در آینده باشد.");

            if (dtpStartDateOperation.Value > DateTime.Now)
                return new ValidationResult(false, "تاریخ شروع بکار نمی‌تواند در آینده باشد.");

            if (dtpHireDate.Value > dtpStartDateOperation.Value)
                return new ValidationResult(false, "تاریخ استخدام نمی‌تواند بعد از تاریخ شروع بکار باشد.");

            return new ValidationResult(true, "");
        }

        private void ClearForm()
        {
            try
            {
                cbProvince.SelectedIndex = -1;
                cbCity.SelectedIndex = -1;
                cbAffair.SelectedIndex = -1;
                cbDept.SelectedIndex = -1;
                cbDistrict.SelectedIndex = -1;
                cbPostName.SelectedIndex = -1;
                cbVoltage.SelectedIndex = -1;
                cbWorkShift.SelectedIndex = -1;
                cbGender.SelectedIndex = -1;
                cbContractType.SelectedIndex = -1;
                cbJobLevel.SelectedIndex = -1;
                cbCompany.SelectedIndex = -1;
                cbDegree.SelectedIndex = -1;
                cbDegreeField.SelectedIndex = -1;
                cbMainJobTitle.SelectedIndex = -1;
                cbCurrentActivity.SelectedIndex = -1;
                cbStatus.SelectedIndex = -1;

                txtFirstName.Clear();
                txtLastName.Clear();
                txtFatherName.Clear();
                txtPersonnelNumber.Clear();
                txtNationalID.Clear();
                txtMobileNumber.Clear();
                txtDescription.Clear();

                dtpBirthDate.Value = DateTime.Now.AddYears(-25);
                dtpHireDate.Value = DateTime.Now;
                dtpStartDateOperation.Value = DateTime.Now;

                chkInconsistency.Checked = false;
                chkInconsistency.Text = "ندارد";

                // ⭐ پاک کردن عکس
                selectedPhotoPath = string.Empty;
                pbPhoto.Image = ImageHelper.CreateDefaultImage(pbPhoto.Width, pbPhoto.Height);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در پاک کردن فرم: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                ValidationResult validationResult = ValidateFormData();
                if (!validationResult.IsValid)
                {
                    MessageBox.Show(validationResult.Message, "خطای اعتبارسنجی", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string query = @"INSERT INTO Personnel 
                    (ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, VoltageID, WorkShiftID, GenderID, 
                     FirstName, LastName, FatherName, PersonnelNumber, NationalID, MobileNumber, 
                     BirthDate, HireDate, StartDateOperation, ContractTypeID, JobLevelID, CompanyID, 
                     DegreeID, DegreeFieldID, MainJobTitle, CurrentActivity, Inconsistency, Description, StatusID) 
                    VALUES 
                    (@ProvinceID, @CityID, @AffairID, @DeptID, @DistrictID, @PostNameID, @VoltageID, @WorkShiftID, @GenderID, 
                     @FirstName, @LastName, @FatherName, @PersonnelNumber, @NationalID, @MobileNumber, 
                     @BirthDate, @HireDate, @StartDateOperation, @ContractTypeID, @JobLevelID, @CompanyID, 
                     @DegreeID, @DegreeFieldID, @MainJobTitle, @CurrentActivity, @Inconsistency, @Description, @StatusID)";

                OleDbParameter[] parameters = new OleDbParameter[]
                {
                    new OleDbParameter("@ProvinceID", GetComboBoxValue(cbProvince)),
                    new OleDbParameter("@CityID", GetComboBoxValue(cbCity)),
                    new OleDbParameter("@AffairID", GetComboBoxValue(cbAffair)),
                    new OleDbParameter("@DeptID", GetComboBoxValue(cbDept)),
                    new OleDbParameter("@DistrictID", GetComboBoxValue(cbDistrict)),
                    new OleDbParameter("@PostNameID", GetComboBoxValue(cbPostName)),
                    new OleDbParameter("@VoltageID", GetComboBoxValue(cbVoltage)),
                    new OleDbParameter("@WorkShiftID", GetComboBoxValue(cbWorkShift)),
                    new OleDbParameter("@GenderID", GetComboBoxValue(cbGender)),
                    new OleDbParameter("@FirstName", txtFirstName.Text.Trim()),
                    new OleDbParameter("@LastName", txtLastName.Text.Trim()),
                    new OleDbParameter("@FatherName", txtFatherName.Text.Trim()),
                    new OleDbParameter("@PersonnelNumber", txtPersonnelNumber.Text.Trim()),
                    new OleDbParameter("@NationalID", txtNationalID.Text.Trim()),
                    new OleDbParameter("@MobileNumber", txtMobileNumber.Text.Trim()),
                    new OleDbParameter("@BirthDate", dtpBirthDate.Value.ToString("yyyy-MM-dd")),
                    new OleDbParameter("@HireDate", dtpHireDate.Value.ToString("yyyy-MM-dd")),
                    new OleDbParameter("@StartDateOperation", dtpStartDateOperation.Value.ToString("yyyy-MM-dd")),
                    new OleDbParameter("@ContractTypeID", GetComboBoxValue(cbContractType)),
                    new OleDbParameter("@JobLevelID", GetComboBoxValue(cbJobLevel)),
                    new OleDbParameter("@CompanyID", GetComboBoxValue(cbCompany)),
                    new OleDbParameter("@DegreeID", GetComboBoxValue(cbDegree)),
                    new OleDbParameter("@DegreeFieldID", GetComboBoxValue(cbDegreeField)),
                    new OleDbParameter("@MainJobTitle", GetComboBoxValue(cbMainJobTitle)),
                    new OleDbParameter("@CurrentActivity", GetComboBoxValue(cbCurrentActivity)),
                    new OleDbParameter("@Inconsistency", chkInconsistency.Checked ? 1 : 0),
                    new OleDbParameter("@Description", string.IsNullOrWhiteSpace(txtDescription.Text) ? DBNull.Value : (object)txtDescription.Text.Trim()),
                    new OleDbParameter("@StatusID", GetComboBoxValue(cbStatus))
                };

                db.ExecuteNonQuery(query, parameters);

                // ⭐ ذخیره عکس اگر انتخاب شده است
                if (!string.IsNullOrEmpty(selectedPhotoPath))
                {
                    string nationalID = txtNationalID.Text.Trim();
                    bool photoSaved = ImageHelper.SaveImage(selectedPhotoPath, nationalID);
                    
                    if (!photoSaved)
                    {
                        MessageBox.Show("پرسنل ثبت شد اما عکس ذخیره نشد.", "اخطار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                MessageBox.Show("پرسنل با موفقیت ثبت شد!", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ClearForm();
            }
            catch (OleDbException oleEx)
            {
                MessageBox.Show($"خطای پایگاه داده: {oleEx.Message}\n\nجزئیات: {oleEx.InnerException?.Message}", "خطای پایگاه داده", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در ثبت پرسنل:\n{ex.Message}\n\nStackTrace: {ex.StackTrace}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private object GetComboBoxValue(ComboBox comboBox)
        {
            try
            {
                if (comboBox.SelectedIndex != -1 && comboBox.SelectedValue != null)
                {
                    if (int.TryParse(comboBox.SelectedValue.ToString(), out int value))
                    {
                        return value;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"خطا در دریافت مقدار کمبوباکس: {ex.Message}");
            }
            return DBNull.Value;
        }

        private class ValidationResult
        {
            public bool IsValid { get; set; }
            public string Message { get; set; }

            public ValidationResult(bool isValid, string message)
            {
                IsValid = isValid;
                Message = message;
            }
        }
    }
}