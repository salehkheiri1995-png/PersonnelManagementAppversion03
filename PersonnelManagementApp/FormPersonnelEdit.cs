using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Linq;

namespace PersonnelManagementApp
{
    public partial class FormPersonnelEdit : Form
    {
        private DbHelper db = new DbHelper();
        public TextBox txtPersonnelID;
        private TextBox txtFreeSearch;
        private DataTable searchTablePersonnel;
        private DataTable provincesTable, citiesTable, transferAffairsTable, operationDepartmentsTable, districtsTable, postsNamesTable;
        private DataTable voltageLevelsTable, workShiftsTable, gendersTable, contractTypesTable, jobLevelsTable, companiesTable;
        private DataTable degreesTable, degreeFieldsTable, chartAffairsTable, statusPresenceTable;
        private ComboBox cmbProvince, cmbCity, cmbAffair, cmbDept, cmbDistrict, cmbPostName;
        private ComboBox cmbVoltage, cmbWorkShift, cmbGender, cmbContractType, cmbJobLevel, cmbCompany;
        private ComboBox cmbDegree, cmbDegreeField, cmbMainJobTitle, cmbCurrentActivity, cmbStatus;
        private TextBox txtFirstName, txtLastName, txtFatherName, txtPersonnelNumber, txtNationalID;
        private TextBox txtMobileNumber, txtBirthDate, txtHireDate, txtStartDateOperation;
        private CheckBox chkInconsistency;
        private TextBox txtDescription;

        // ⭐ فیلدهای جدید برای مدیریت عکس
        private PictureBox pbPhoto;
        private string selectedPhotoPath = string.Empty;

        public FormPersonnelEdit()
        {
            InitializeComponent();

            // اعمال فونت‌های تنظیم‌شده
            FontSettings.ApplyFontToForm(this);

            LoadData();
            CreateSearchTable();
            LoadProvinces();
            LoadOtherCombos();
        }

        private void InitializeComponent()
        {
            this.Text = "ویرایش پرسنل";
            this.Size = new Size(600, 900);
            this.WindowState = FormWindowState.Maximized;
            this.AutoScroll = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.RightToLeft = RightToLeft.Yes;
            this.BackColor = Color.FromArgb(240, 248, 255);

            // پس‌زمینه گرادیانت
            using (LinearGradientBrush brush = new LinearGradientBrush(this.ClientRectangle, Color.LightBlue, Color.SkyBlue, LinearGradientMode.Vertical))
            {
                brush.GammaCorrection = true;
                this.BackgroundImage = new Bitmap(this.Width, this.Height);
                using (Graphics g = Graphics.FromImage(this.BackgroundImage))
                {
                    g.FillRectangle(brush, this.ClientRectangle);
                }
            }

            int yHeader = 20;
            int ySearch = 100; // موقعیت بخش جستجو (زیر سرتیتر)
            int yRight = 550; // شروع ستون راست بعد از گرید
            int yLeft = 550;  // شروع ستون چپ بعد از گرید
            int controlWidth = 300;
            int controlHeight = 40;
            int yStep = 60;
            int formWidth = this.ClientSize.Width;
            int labelWidth = 250; // افزایش عرض لیبل برای جایگیری بهتر متن
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
                Text = "ویرایش اطلاعات پرسنل",
                Location = new Point((formWidth - 400) / 2, yHeader),
                Size = new Size(400, 50),
                Font = FontSettings.TitleFont,
                ForeColor = Color.Navy,
                TextAlign = ContentAlignment.MiddleCenter
            };
            this.Controls.Add(lblHeader);
            yHeader += 60;

            // TextBox برای PersonnelID
            Label lblPersonnelID = new Label { Text = "شناسه پرسنل:", Location = new Point((formWidth - labelWidth) / 2 + controlWidth + 95, ySearch), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtPersonnelID = new TextBox { Location = new Point((formWidth - controlWidth) / 2, ySearch), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            Button btnLoad = new Button { Text = "بارگذاری", Location = new Point((formWidth - controlWidth) / 2 + controlWidth + 20, ySearch - 5), Size = new Size(100, controlHeight), Font = FontSettings.ButtonFont, BackColor = Color.LightGreen, ForeColor = Color.White };
            btnLoad.Click += BtnLoad_Click;
            ApplyRoundedCorners(txtPersonnelID, 10);
            ApplyRoundedCorners(btnLoad, 15);
            btnLoad.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) btnLoad.BackColor = Color.ForestGreen; else btnLoad.BackColor = Color.LightGreen; };
            this.Controls.Add(lblPersonnelID);
            this.Controls.Add(txtPersonnelID);
            this.Controls.Add(btnLoad);
            ySearch += yStep;

            // جستجوی آزاد
            Label lblFreeSearch = new Label { Text = "جستجوی آزاد:", Location = new Point((formWidth - labelWidth) / 2 + controlWidth + 95, ySearch), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtFreeSearch = new TextBox { Location = new Point((formWidth - controlWidth) / 2, ySearch), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont, PlaceholderText = "جستجوی پرسنل..." };
            Button btnSearch = new Button { Text = "جستجو", Location = new Point((formWidth - controlWidth) / 2 + controlWidth + 20, ySearch - 5), Size = new Size(100, controlHeight), Font = FontSettings.ButtonFont, BackColor = Color.LightGreen, ForeColor = Color.White };
            btnSearch.Click += BtnSearch_Click;
            ApplyRoundedCorners(txtFreeSearch, 10);
            ApplyRoundedCorners(btnSearch, 15);
            btnSearch.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) btnSearch.BackColor = Color.RoyalBlue; else btnSearch.BackColor = Color.LightBlue; };
            this.Controls.Add(lblFreeSearch);
            this.Controls.Add(txtFreeSearch);
            this.Controls.Add(btnSearch);
            ySearch += yStep;

            // گرید نمایش نتایج جستجو
            DataGridView dgv = new DataGridView
            {
                Name = "dataGridView1",
                Location = new Point((formWidth - (controlWidth + 120)) / 2, ySearch),
                Size = new Size(controlWidth + 120, 200),
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                AllowUserToAddRows = false,
                ReadOnly = true,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None
            };
            ApplyRoundedCorners(dgv, 20);
            dgv.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) dgv.BackColor = Color.LightYellow; else dgv.BackColor = Color.White; };
            dgv.CellClick += Dgv_CellClick;
            this.Controls.Add(dgv);
            ySearch += 220;

            // ⭐ بخش عکس - اضافه شده
            int photoBoxSize = 200;
            pbPhoto = new PictureBox
            {
                Location = new Point((formWidth - photoBoxSize) / 2, ySearch),
                Size = new Size(photoBoxSize, photoBoxSize),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White
            };
            pbPhoto.Image = ImageHelper.CreateDefaultImage(photoBoxSize, photoBoxSize);
            ApplyRoundedCorners(pbPhoto, 15);
            this.Controls.Add(pbPhoto);
            ySearch += photoBoxSize + 10;

            // دکمه‌های مدیریت عکس
            int btnPhotoWidth = 95;
            int btnPhotoSpacing = 5;
            int totalPhotoButtonWidth = (btnPhotoWidth * 2) + btnPhotoSpacing;
            int xPhotoButtonStart = (formWidth - totalPhotoButtonWidth) / 2;

            Button btnSelectPhoto = new Button
            {
                Text = "انتخاب عکس",
                Location = new Point(xPhotoButtonStart, ySearch),
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
                Location = new Point(xPhotoButtonStart + btnPhotoWidth + btnPhotoSpacing, ySearch),
                Size = new Size(btnPhotoWidth, 35),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightCoral,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnRemovePhoto, 10);
            btnRemovePhoto.Click += BtnRemovePhoto_Click;
            this.Controls.Add(btnRemovePhoto);
            ySearch += 45;

            // تنظیم مجدد yRight و yLeft
            yRight = ySearch + 10;
            yLeft = ySearch + 10;

            // *** ستون سمت راست: فیلدهای 1 تا 13 ***
            // ComboBox برای استان
            Label lblProvince = new Label { Text = "استان:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbProvince = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cmbProvince.SelectedIndexChanged += CmbProvince_SelectedIndexChanged;
            ApplyRoundedCorners(cmbProvince, 10);
            this.Controls.Add(lblProvince);
            this.Controls.Add(cmbProvince);
            yRight += yStep;

            // ComboBox برای شهر
            Label lblCity = new Label { Text = "شهر:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbCity = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbCity, 10);
            this.Controls.Add(lblCity);
            this.Controls.Add(cmbCity);
            yRight += yStep;

            // ComboBox برای امور انتقال
            Label lblAffair = new Label { Text = "امور انتقال:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbAffair = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cmbAffair.SelectedIndexChanged += CmbAffair_SelectedIndexChanged;
            ApplyRoundedCorners(cmbAffair, 10);
            this.Controls.Add(lblAffair);
            this.Controls.Add(cmbAffair);
            yRight += yStep;

            // ComboBox برای اداره بهره‌برداری
            Label lblDept = new Label { Text = "اداره بهره‌برداری:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbDept = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cmbDept.SelectedIndexChanged += CmbDept_SelectedIndexChanged;
            ApplyRoundedCorners(cmbDept, 10);
            this.Controls.Add(lblDept);
            this.Controls.Add(cmbDept);
            yRight += yStep;

            // ComboBox برای ناحیه بهره‌برداری
            Label lblDistrict = new Label { Text = "ناحیه بهره‌برداری:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbDistrict = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            cmbDistrict.SelectedIndexChanged += CmbDistrict_SelectedIndexChanged;
            ApplyRoundedCorners(cmbDistrict, 10);
            this.Controls.Add(lblDistrict);
            this.Controls.Add(cmbDistrict);
            yRight += yStep;

            // ComboBox برای نام پست
            Label lblPostName = new Label { Text = "نام پست:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbPostName = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbPostName, 10);
            this.Controls.Add(lblPostName);
            this.Controls.Add(cmbPostName);
            yRight += yStep;

            // ComboBox برای سطح ولتاژ
            Label lblVoltage = new Label { Text = "سطح ولتاژ:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbVoltage = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbVoltage, 10);
            this.Controls.Add(lblVoltage);
            this.Controls.Add(cmbVoltage);
            yRight += yStep;

            // ComboBox برای شیفت کاری
            Label lblWorkShift = new Label { Text = "شیفت کاری:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbWorkShift = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbWorkShift, 10);
            this.Controls.Add(lblWorkShift);
            this.Controls.Add(cmbWorkShift);
            yRight += yStep;

            // ComboBox برای جنسیت
            Label lblGender = new Label { Text = "جنسیت:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbGender = new ComboBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbGender, 10);
            this.Controls.Add(lblGender);
            this.Controls.Add(cmbGender);
            yRight += yStep;

            // TextBox برای نام
            Label lblFirstName = new Label { Text = "نام:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtFirstName = new TextBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtFirstName, 10);
            this.Controls.Add(lblFirstName);
            this.Controls.Add(txtFirstName);
            yRight += yStep;

            // TextBox برای نام خانوادگی
            Label lblLastName = new Label { Text = "نام خانوادگی:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtLastName = new TextBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtLastName, 10);
            this.Controls.Add(lblLastName);
            this.Controls.Add(txtLastName);
            yRight += yStep;

            // TextBox برای نام پدر
            Label lblFatherName = new Label { Text = "نام پدر:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtFatherName = new TextBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtFatherName, 10);
            this.Controls.Add(lblFatherName);
            this.Controls.Add(txtFatherName);
            yRight += yStep;

            // TextBox برای شماره پرسنلی
            Label lblPersonnelNumber = new Label { Text = "شماره پرسنلی:", Location = new Point(xLabelRight, yRight), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtPersonnelNumber = new TextBox { Location = new Point(xControlRight, yRight), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtPersonnelNumber, 10);
            this.Controls.Add(lblPersonnelNumber);
            this.Controls.Add(txtPersonnelNumber);
            yRight += yStep;

            // *** ستون سمت چپ: فیلدهای 14 تا 28 ***
            // TextBox برای کد ملی
            Label lblNationalID = new Label { Text = "کد ملی:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtNationalID = new TextBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtNationalID, 10);
            this.Controls.Add(lblNationalID);
            this.Controls.Add(txtNationalID);
            yLeft += yStep;

            // TextBox برای شماره موبایل
            Label lblMobileNumber = new Label { Text = "شماره موبایل:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtMobileNumber = new TextBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtMobileNumber, 10);
            this.Controls.Add(lblMobileNumber);
            this.Controls.Add(txtMobileNumber);
            yLeft += yStep;

            // TextBox برای تاریخ تولد
            Label lblBirthDate = new Label { Text = "تاریخ تولد:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtBirthDate = new TextBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtBirthDate, 10);
            this.Controls.Add(lblBirthDate);
            this.Controls.Add(txtBirthDate);
            yLeft += yStep;

            // TextBox برای تاریخ استخدام
            Label lblHireDate = new Label { Text = "تاریخ استخدام:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtHireDate = new TextBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtHireDate, 10);
            this.Controls.Add(lblHireDate);
            this.Controls.Add(txtHireDate);
            yLeft += yStep;

            // TextBox برای تاریخ شروع به کار
            Label lblStartDateOperation = new Label { Text = "تاریخ شروع به کار:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtStartDateOperation = new TextBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtStartDateOperation, 10);
            this.Controls.Add(lblStartDateOperation);
            this.Controls.Add(txtStartDateOperation);
            yLeft += yStep;

            // ComboBox برای نوع قرارداد
            Label lblContractType = new Label { Text = "نوع قرارداد:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbContractType = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbContractType, 10);
            this.Controls.Add(lblContractType);
            this.Controls.Add(cmbContractType);
            yLeft += yStep;

            // ComboBox برای سطح شغل
            Label lblJobLevel = new Label { Text = "سطح شغل:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbJobLevel = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbJobLevel, 10);
            this.Controls.Add(lblJobLevel);
            this.Controls.Add(cmbJobLevel);
            yLeft += yStep;

            // ComboBox برای شرکت
            Label lblCompany = new Label { Text = "شرکت:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbCompany = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbCompany, 10);
            this.Controls.Add(lblCompany);
            this.Controls.Add(cmbCompany);
            yLeft += yStep;

            // ComboBox برای مدرک تحصیلی
            Label lblDegree = new Label { Text = "مدرک تحصیلی:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbDegree = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbDegree, 10);
            this.Controls.Add(lblDegree);
            this.Controls.Add(cmbDegree);
            yLeft += yStep;

            // ComboBox برای رشته تحصیلی
            Label lblDegreeField = new Label { Text = "رشته تحصیلی:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbDegreeField = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbDegreeField, 10);
            this.Controls.Add(lblDegreeField);
            this.Controls.Add(cmbDegreeField);
            yLeft += yStep;

            // ComboBox برای عنوان شغلی اصلی
            Label lblMainJobTitle = new Label { Text = "عنوان شغلی اصلی:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbMainJobTitle = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbMainJobTitle, 10);
            this.Controls.Add(lblMainJobTitle);
            this.Controls.Add(cmbMainJobTitle);
            yLeft += yStep;

            // ComboBox برای فعالیت فعلی
            Label lblCurrentActivity = new Label { Text = "فعالیت فعلی:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbCurrentActivity = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbCurrentActivity, 10);
            this.Controls.Add(lblCurrentActivity);
            this.Controls.Add(cmbCurrentActivity);
            yLeft += yStep;

            // CheckBox برای مغایرت
            Label lblInconsistency = new Label { Text = "مغایرت:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            chkInconsistency = new CheckBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), Font = FontSettings.BodyFont };
            this.Controls.Add(lblInconsistency);
            this.Controls.Add(chkInconsistency);
            yLeft += yStep;

            // ComboBox برای وضعیت حضور
            Label lblStatus = new Label { Text = "وضعیت حضور:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            cmbStatus = new ComboBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, controlHeight), DropDownStyle = ComboBoxStyle.DropDownList, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(cmbStatus, 10);
            this.Controls.Add(lblStatus);
            this.Controls.Add(cmbStatus);
            yLeft += yStep;

            // TextBox برای توضیحات
            Label lblDescription = new Label { Text = "توضیحات:", Location = new Point(xLabelLeft, yLeft), Size = new Size(labelWidth, 30), Font = FontSettings.LabelFont };
            txtDescription = new TextBox { Location = new Point(xControlLeft, yLeft), Size = new Size(controlWidth, 100), Multiline = true, Font = FontSettings.TextBoxFont };
            ApplyRoundedCorners(txtDescription, 10);
            this.Controls.Add(lblDescription);
            this.Controls.Add(txtDescription);
            yLeft += yStep + 60;

            // دکمه‌های به‌روزرسانی و لغو
            int maxY = Math.Max(yRight, yLeft) + 20;
            int buttonWidth = 150;
            int buttonSpace = 10;
            int totalButtonWidth = (buttonWidth * 2) + buttonSpace;
            int xButtonStart = (formWidth - totalButtonWidth) / 2;

            Button btnUpdate = new Button
            {
                Text = "به‌روزرسانی",
                Location = new Point(xButtonStart, maxY),
                Size = new Size(buttonWidth, 50),
                Font = FontSettings.ButtonFont,
                BackColor = Color.LightGreen,
                ForeColor = Color.White
            };
            ApplyRoundedCorners(btnUpdate, 15);
            btnUpdate.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) btnUpdate.BackColor = Color.ForestGreen; else btnUpdate.BackColor = Color.LightGreen; };
            btnUpdate.Click += BtnUpdate_Click;
            this.Controls.Add(btnUpdate);

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
            btnCancel.MouseClick += (s, e) => { if (e.Button == MouseButtons.Left) btnCancel.BackColor = Color.Red; else btnCancel.BackColor = Color.LightCoral; };
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
            GraphicsPath path = new GraphicsPath();
            path.AddArc(0, 0, radius, radius, 180, 90);
            path.AddArc(control.Width - radius, 0, radius, radius, 270, 90);
            path.AddArc(control.Width - radius, control.Height - radius, radius, radius, 0, 90);
            path.AddArc(0, control.Height - radius, radius, radius, 90, 90);
            path.CloseFigure();
            control.Region = new Region(path);
        }

        private void LoadData()
        {
            try
            {
                provincesTable = db.GetProvinces();
                citiesTable = db.GetCitiesByProvince(0);
                transferAffairsTable = db.GetAffairsByProvince(0);
                operationDepartmentsTable = db.GetDeptsByAffair(0);
                districtsTable = db.GetDistrictsByDept(0);
                postsNamesTable = db.GetPostNamesByDistrict(0);
                voltageLevelsTable = db.GetVoltageLevels();
                workShiftsTable = db.GetWorkShifts();
                gendersTable = db.GetGenders();
                contractTypesTable = db.GetContractTypes();
                jobLevelsTable = db.GetJobLevels();
                companiesTable = db.GetCompanies();
                degreesTable = db.GetDegrees();
                degreeFieldsTable = db.GetDegreeFields();
                chartAffairsTable = db.GetChartAffairs();
                statusPresenceTable = db.GetStatusPresence();
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در بارگذاری داده‌ها: " + ex.Message);
            }
        }

        private void CreateSearchTable()
        {
            searchTablePersonnel = new DataTable("SearchTablePersonnel");
            searchTablePersonnel.Columns.Add("PersonnelID", typeof(int)).ColumnName = "شناسه پرسنل";
            searchTablePersonnel.Columns.Add("FirstName", typeof(string)).ColumnName = "نام";
            searchTablePersonnel.Columns.Add("LastName", typeof(string)).ColumnName = "نام خانوادگی";
            searchTablePersonnel.Columns.Add("FatherName", typeof(string)).ColumnName = "نام پدر";
            searchTablePersonnel.Columns.Add("PersonnelNumber", typeof(string)).ColumnName = "شماره پرسنلی";
            searchTablePersonnel.Columns.Add("NationalID", typeof(string)).ColumnName = "کد ملی";
            searchTablePersonnel.Columns.Add("MobileNumber", typeof(string)).ColumnName = "شماره موبایل";
            searchTablePersonnel.Columns.Add("ProvinceName", typeof(string)).ColumnName = "استان";
            searchTablePersonnel.Columns.Add("CityName", typeof(string)).ColumnName = "شهر";
            searchTablePersonnel.Columns.Add("AffairName", typeof(string)).ColumnName = "امور";
            searchTablePersonnel.Columns.Add("DeptName", typeof(string)).ColumnName = "اداره";
            searchTablePersonnel.Columns.Add("DistrictName", typeof(string)).ColumnName = "ناحیه";
            searchTablePersonnel.Columns.Add("PostName", typeof(string)).ColumnName = "نام پست";
            searchTablePersonnel.Columns.Add("VoltageName", typeof(string)).ColumnName = "سطح ولتاژ";
            searchTablePersonnel.Columns.Add("WorkShiftName", typeof(string)).ColumnName = "شیفت کاری";
            searchTablePersonnel.Columns.Add("GenderName", typeof(string)).ColumnName = "جنسیت";
            searchTablePersonnel.Columns.Add("ContractTypeName", typeof(string)).ColumnName = "نوع قرارداد";
            searchTablePersonnel.Columns.Add("JobLevelName", typeof(string)).ColumnName = "سطح شغل";
            searchTablePersonnel.Columns.Add("CompanyName", typeof(string)).ColumnName = "شرکت";
            searchTablePersonnel.Columns.Add("DegreeName", typeof(string)).ColumnName = "مدرک تحصیلی";
            searchTablePersonnel.Columns.Add("DegreeFieldName", typeof(string)).ColumnName = "رشته تحصیلی";
            searchTablePersonnel.Columns.Add("MainJobTitle", typeof(string)).ColumnName = "عنوان شغلی اصلی";
            searchTablePersonnel.Columns.Add("CurrentActivity", typeof(string)).ColumnName = "فعالیت فعلی";
            searchTablePersonnel.Columns.Add("Inconsistency", typeof(bool)).ColumnName = "مغایرت";
            searchTablePersonnel.Columns.Add("Description", typeof(string)).ColumnName = "توضیحات";
            searchTablePersonnel.Columns.Add("StatusName", typeof(string)).ColumnName = "وضعیت حضور";
            searchTablePersonnel.Columns.Add("BirthDate", typeof(string)).ColumnName = "تاریخ تولد";
            searchTablePersonnel.Columns.Add("HireDate", typeof(string)).ColumnName = "تاریخ استخدام";
            searchTablePersonnel.Columns.Add("StartDateOperation", typeof(string)).ColumnName = "تاریخ شروع به کار";

            DataTable personnelTable = db.ExecuteQuery("SELECT PersonnelID, ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, VoltageID, " +
                                                      "WorkShiftID, GenderID, FirstName, LastName, FatherName, PersonnelNumber, NationalID, " +
                                                      "MobileNumber, BirthDate, HireDate, StartDateOperation, ContractTypeID, JobLevelID, " +
                                                      "CompanyID, DegreeID, DegreeFieldID, MainJobTitle, CurrentActivity, Inconsistency, " +
                                                      "Description, StatusID FROM Personnel");

            foreach (DataRow personnelRow in personnelTable.Rows)
            {
                DataRow newRow = searchTablePersonnel.NewRow();
                newRow["شناسه پرسنل"] = personnelRow["PersonnelID"];
                newRow["نام"] = personnelRow["FirstName"];
                newRow["نام خانوادگی"] = personnelRow["LastName"];
                newRow["نام پدر"] = personnelRow["FatherName"];
                newRow["شماره پرسنلی"] = personnelRow["PersonnelNumber"];
                newRow["کد ملی"] = personnelRow["NationalID"];
                newRow["شماره موبایل"] = personnelRow["MobileNumber"];
                newRow["تاریخ تولد"] = personnelRow["BirthDate"];
                newRow["تاریخ استخدام"] = personnelRow["HireDate"];
                newRow["تاریخ شروع به کار"] = personnelRow["StartDateOperation"];
                newRow["مغایرت"] = personnelRow["Inconsistency"];
                newRow["توضیحات"] = personnelRow["Description"];

                int? provinceId = personnelRow["ProvinceID"] as int?;
                newRow["استان"] = provinceId.HasValue ? provincesTable.Select($"ProvinceID = {provinceId.Value}").FirstOrDefault()?["ProvinceName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? cityId = personnelRow["CityID"] as int?;
                newRow["شهر"] = cityId.HasValue ? citiesTable.Select($"CityID = {cityId.Value}").FirstOrDefault()?["CityName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? affairId = personnelRow["AffairID"] as int?;
                newRow["امور"] = affairId.HasValue ? transferAffairsTable.Select($"AffairID = {affairId.Value}").FirstOrDefault()?["AffairName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? deptId = personnelRow["DeptID"] as int?;
                newRow["اداره"] = deptId.HasValue ? operationDepartmentsTable.Select($"DeptID = {deptId.Value}").FirstOrDefault()?["DeptName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? districtId = personnelRow["DistrictID"] as int?;
                newRow["ناحیه"] = districtId.HasValue ? districtsTable.Select($"DistrictID = {districtId.Value}").FirstOrDefault()?["DistrictName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? postNameId = personnelRow["PostNameID"] as int?;
                newRow["نام پست"] = postNameId.HasValue ? postsNamesTable.Select($"PostNameID = {postNameId.Value}").FirstOrDefault()?["PostName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? voltageId = personnelRow["VoltageID"] as int?;
                newRow["سطح ولتاژ"] = voltageId.HasValue ? voltageLevelsTable.Select($"VoltageID = {voltageId.Value}").FirstOrDefault()?["VoltageName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? workShiftId = personnelRow["WorkShiftID"] as int?;
                newRow["شیفت کاری"] = workShiftId.HasValue ? workShiftsTable.Select($"WorkShiftID = {workShiftId.Value}").FirstOrDefault()?["WorkShiftName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? genderId = personnelRow["GenderID"] as int?;
                newRow["جنسیت"] = genderId.HasValue ? gendersTable.Select($"GenderID = {genderId.Value}").FirstOrDefault()?["GenderName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? contractTypeId = personnelRow["ContractTypeID"] as int?;
                newRow["نوع قرارداد"] = contractTypeId.HasValue ? contractTypesTable.Select($"ContractTypeID = {contractTypeId.Value}").FirstOrDefault()?["ContractTypeName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? jobLevelId = personnelRow["JobLevelID"] as int?;
                newRow["سطح شغل"] = jobLevelId.HasValue ? jobLevelsTable.Select($"JobLevelID = {jobLevelId.Value}").FirstOrDefault()?["JobLevelName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? companyId = personnelRow["CompanyID"] as int?;
                newRow["شرکت"] = companyId.HasValue ? companiesTable.Select($"CompanyID = {companyId.Value}").FirstOrDefault()?["CompanyName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? degreeId = personnelRow["DegreeID"] as int?;
                newRow["مدرک تحصیلی"] = degreeId.HasValue ? degreesTable.Select($"DegreeID = {degreeId.Value}").FirstOrDefault()?["DegreeName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? degreeFieldId = personnelRow["DegreeFieldID"] as int?;
                newRow["رشته تحصیلی"] = degreeFieldId.HasValue ? degreeFieldsTable.Select($"DegreeFieldID = {degreeFieldId.Value}").FirstOrDefault()?["DegreeFieldName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? mainJobTitleId = personnelRow["MainJobTitle"] as int?;
                newRow["عنوان شغلی اصلی"] = mainJobTitleId.HasValue ? chartAffairsTable.Select($"ChartID = {mainJobTitleId.Value}").FirstOrDefault()?["ChartName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? currentActivityId = personnelRow["CurrentActivity"] as int?;
                newRow["فعالیت فعلی"] = currentActivityId.HasValue ? chartAffairsTable.Select($"ChartID = {currentActivityId.Value}").FirstOrDefault()?["ChartName"]?.ToString() ?? "نامشخص" : "نامشخص";

                int? statusId = personnelRow["StatusID"] as int?;
                newRow["وضعیت حضور"] = statusId.HasValue ? statusPresenceTable.Select($"StatusID = {statusId.Value}").FirstOrDefault()?["StatusName"]?.ToString() ?? "نامشخص" : "نامشخص";

                searchTablePersonnel.Rows.Add(newRow);
            }
        }

        private void LoadProvinces()
        {
            DataTable dt = db.GetProvinces();
            cmbProvince.DataSource = dt;
            cmbProvince.DisplayMember = "ProvinceName";
            cmbProvince.ValueMember = "ProvinceID";
        }

        private void LoadOtherCombos()
        {
            cmbVoltage.DataSource = db.GetVoltageLevels();
            cmbVoltage.DisplayMember = "VoltageName";
            cmbVoltage.ValueMember = "VoltageID";

            cmbWorkShift.DataSource = db.GetWorkShifts();
            cmbWorkShift.DisplayMember = "WorkShiftName";
            cmbWorkShift.ValueMember = "WorkShiftID";

            cmbGender.DataSource = db.GetGenders();
            cmbGender.DisplayMember = "GenderName";
            cmbGender.ValueMember = "GenderID";

            cmbContractType.DataSource = db.GetContractTypes();
            cmbContractType.DisplayMember = "ContractTypeName";
            cmbContractType.ValueMember = "ContractTypeID";

            cmbJobLevel.DataSource = db.GetJobLevels();
            cmbJobLevel.DisplayMember = "JobLevelName";
            cmbJobLevel.ValueMember = "JobLevelID";

            cmbCompany.DataSource = db.GetCompanies();
            cmbCompany.DisplayMember = "CompanyName";
            cmbCompany.ValueMember = "CompanyID";

            cmbDegree.DataSource = db.GetDegrees();
            cmbDegree.DisplayMember = "DegreeName";
            cmbDegree.ValueMember = "DegreeID";

            cmbDegreeField.DataSource = db.GetDegreeFields();
            cmbDegreeField.DisplayMember = "DegreeFieldName";
            cmbDegreeField.ValueMember = "DegreeFieldID";

            cmbMainJobTitle.DataSource = db.GetChartAffairs();
            cmbMainJobTitle.DisplayMember = "ChartName";
            cmbMainJobTitle.ValueMember = "ChartID";

            cmbCurrentActivity.DataSource = db.GetChartAffairs();
            cmbCurrentActivity.DisplayMember = "ChartName";
            cmbCurrentActivity.ValueMember = "ChartID";

            cmbStatus.DataSource = db.GetStatusPresence();
            cmbStatus.DisplayMember = "StatusName";
            cmbStatus.ValueMember = "StatusID";
        }

        private void CmbProvince_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbProvince.SelectedValue != null && int.TryParse(cmbProvince.SelectedValue.ToString(), out int provinceID))
            {
                cmbCity.DataSource = db.GetCitiesByProvince(provinceID);
                cmbCity.DisplayMember = "CityName";
                cmbCity.ValueMember = "CityID";

                cmbAffair.DataSource = db.GetAffairsByProvince(provinceID);
                cmbAffair.DisplayMember = "AffairName";
                cmbAffair.ValueMember = "AffairID";

                cmbDept.DataSource = null;
                cmbDistrict.DataSource = null;
                cmbPostName.DataSource = null;
            }
        }

        private void CmbAffair_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbAffair.SelectedValue != null && int.TryParse(cmbAffair.SelectedValue.ToString(), out int affairID))
            {
                cmbDept.DataSource = db.GetDeptsByAffair(affairID);
                cmbDept.DisplayMember = "DeptName";
                cmbDept.ValueMember = "DeptID";

                cmbDistrict.DataSource = null;
                cmbPostName.DataSource = null;
            }
        }

        private void CmbDept_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbDept.SelectedValue != null && int.TryParse(cmbDept.SelectedValue.ToString(), out int deptID))
            {
                cmbDistrict.DataSource = db.GetDistrictsByDept(deptID);
                cmbDistrict.DisplayMember = "DistrictName";
                cmbDistrict.ValueMember = "DistrictID";

                cmbPostName.DataSource = null;
            }
        }

        private void CmbDistrict_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbDistrict.SelectedValue != null && int.TryParse(cmbDistrict.SelectedValue.ToString(), out int districtID))
            {
                cmbPostName.DataSource = db.GetPostNamesByDistrict(districtID);
                cmbPostName.DisplayMember = "PostName";
                cmbPostName.ValueMember = "PostNameID";
            }
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            try
            {
                string searchText = txtFreeSearch.Text.Trim().ToLower();
                var filteredTable = searchTablePersonnel.Clone();
                foreach (DataRow row in searchTablePersonnel.Rows)
                {
                    if (row["نام"].ToString().ToLower().Contains(searchText) ||
                        row["نام خانوادگی"].ToString().ToLower().Contains(searchText) ||
                        row["شماره پرسنلی"].ToString().ToLower().Contains(searchText) ||
                        row["کد ملی"].ToString().ToLower().Contains(searchText))
                    {
                        filteredTable.ImportRow(row);
                    }
                }

                DataGridView dgv = this.Controls.Find("dataGridView1", true)[0] as DataGridView;
                dgv.DataSource = filteredTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در جستجوی آزاد: " + ex.Message);
            }
        }

        private void Dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridView dgv = sender as DataGridView;
                txtPersonnelID.Text = dgv.Rows[e.RowIndex].Cells["شناسه پرسنل"].Value.ToString();
                BtnLoad_Click(sender, EventArgs.Empty);
            }
        }

        public void BtnLoad_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPersonnelID.Text) || !int.TryParse(txtPersonnelID.Text, out int personnelID))
            {
                MessageBox.Show("لطفاً یک شناسه پرسنل معتبر وارد کنید.");
                return;
            }

            string query = "SELECT * FROM Personnel WHERE PersonnelID = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", personnelID) };
            DataTable dt = db.ExecuteQuery(query, parameters);

            if (dt.Rows.Count > 0)
            {
                DataRow row = dt.Rows[0];

                cmbProvince.SelectedValue = row["ProvinceID"];
                CmbProvince_SelectedIndexChanged(null, null);
                cmbCity.SelectedValue = row["CityID"];

                cmbAffair.SelectedValue = row["AffairID"];
                CmbAffair_SelectedIndexChanged(null, null);
                cmbDept.SelectedValue = row["DeptID"];
                CmbDept_SelectedIndexChanged(null, null);
                cmbDistrict.SelectedValue = row["DistrictID"];
                CmbDistrict_SelectedIndexChanged(null, null);
                cmbPostName.SelectedValue = row["PostNameID"];

                cmbVoltage.SelectedValue = row["VoltageID"];
                cmbWorkShift.SelectedValue = row["WorkShiftID"];
                cmbGender.SelectedValue = row["GenderID"];
                txtFirstName.Text = row["FirstName"] != DBNull.Value ? row["FirstName"].ToString() : "";
                txtLastName.Text = row["LastName"] != DBNull.Value ? row["LastName"].ToString() : "";
                txtFatherName.Text = row["FatherName"] != DBNull.Value ? row["FatherName"].ToString() : "";
                txtPersonnelNumber.Text = row["PersonnelNumber"] != DBNull.Value ? row["PersonnelNumber"].ToString() : "";
                txtNationalID.Text = row["NationalID"] != DBNull.Value ? row["NationalID"].ToString() : "";
                txtMobileNumber.Text = row["MobileNumber"] != DBNull.Value ? row["MobileNumber"].ToString() : "";
                txtBirthDate.Text = row["BirthDate"] != DBNull.Value ? row["BirthDate"].ToString() : "";
                txtHireDate.Text = row["HireDate"] != DBNull.Value ? row["HireDate"].ToString() : "";
                txtStartDateOperation.Text = row["StartDateOperation"] != DBNull.Value ? row["StartDateOperation"].ToString() : "";
                cmbContractType.SelectedValue = row["ContractTypeID"];
                cmbJobLevel.SelectedValue = row["JobLevelID"];
                cmbCompany.SelectedValue = row["CompanyID"];
                cmbDegree.SelectedValue = row["DegreeID"];
                cmbDegreeField.SelectedValue = row["DegreeFieldID"];
                cmbMainJobTitle.SelectedValue = row["MainJobTitle"];
                cmbCurrentActivity.SelectedValue = row["CurrentActivity"];
                chkInconsistency.Checked = row["Inconsistency"] != DBNull.Value && Convert.ToBoolean(row["Inconsistency"]);
                txtDescription.Text = row["Description"] != DBNull.Value ? row["Description"].ToString() : "";
                cmbStatus.SelectedValue = row["StatusID"];

                // ⭐ بارگذاری عکس از پوشه اگر وجود دارد
                string nationalID = row["NationalID"] != DBNull.Value ? row["NationalID"].ToString() : "";
                if (!string.IsNullOrEmpty(nationalID))
                {
                    Image photoImage = ImageHelper.LoadImage(nationalID);
                    if (photoImage != null)
                    {
                        ImageHelper.DrawImageInPictureBox(pbPhoto, photoImage);
                        selectedPhotoPath = ""; // عکس از پوشه بارگذاری شده
                    }
                    else
                    {
                        pbPhoto.Image = ImageHelper.CreateDefaultImage(pbPhoto.Width, pbPhoto.Height);
                        selectedPhotoPath = "";
                    }
                }
            }
            else
            {
                MessageBox.Show("رکوردی با این شناسه یافت نشد!");
            }
        }

        private void BtnUpdate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPersonnelID.Text) || !int.TryParse(txtPersonnelID.Text, out int personnelID))
            {
                MessageBox.Show("لطفاً یک شناسه پرسنل معتبر وارد کنید.");
                return;
            }

            string query = "UPDATE Personnel SET ProvinceID = ?, CityID = ?, AffairID = ?, DeptID = ?, DistrictID = ?, PostNameID = ?, VoltageID = ?, WorkShiftID = ?, GenderID = ?, FirstName = ?, LastName = ?, FatherName = ?, PersonnelNumber = ?, NationalID = ?, MobileNumber = ?, BirthDate = ?, HireDate = ?, StartDateOperation = ?, ContractTypeID = ?, JobLevelID = ?, CompanyID = ?, DegreeID = ?, DegreeFieldID = ?, MainJobTitle = ?, CurrentActivity = ?, Inconsistency = ?, Description = ?, StatusID = ? WHERE PersonnelID = ?";

            OleDbParameter[] parameters = new OleDbParameter[]
            {
                new OleDbParameter("?", cmbProvince.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbCity.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbAffair.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbDept.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbDistrict.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbPostName.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbVoltage.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbWorkShift.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbGender.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", string.IsNullOrEmpty(txtFirstName.Text) ? DBNull.Value : txtFirstName.Text),
                new OleDbParameter("?", string.IsNullOrEmpty(txtLastName.Text) ? DBNull.Value : txtLastName.Text),
                new OleDbParameter("?", string.IsNullOrEmpty(txtFatherName.Text) ? DBNull.Value : txtFatherName.Text),
                new OleDbParameter("?", string.IsNullOrEmpty(txtPersonnelNumber.Text) ? DBNull.Value : txtPersonnelNumber.Text),
                new OleDbParameter("?", string.IsNullOrEmpty(txtNationalID.Text) ? DBNull.Value : txtNationalID.Text),
                new OleDbParameter("?", string.IsNullOrEmpty(txtMobileNumber.Text) ? DBNull.Value : txtMobileNumber.Text),
                new OleDbParameter("?", string.IsNullOrEmpty(txtBirthDate.Text) ? DBNull.Value : txtBirthDate.Text),
                new OleDbParameter("?", string.IsNullOrEmpty(txtHireDate.Text) ? DBNull.Value : txtHireDate.Text),
                new OleDbParameter("?", string.IsNullOrEmpty(txtStartDateOperation.Text) ? DBNull.Value : txtStartDateOperation.Text),
                new OleDbParameter("?", cmbContractType.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbJobLevel.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbCompany.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbDegree.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbDegreeField.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbMainJobTitle.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", cmbCurrentActivity.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", chkInconsistency.Checked),
                new OleDbParameter("?", string.IsNullOrEmpty(txtDescription.Text) ? DBNull.Value : txtDescription.Text),
                new OleDbParameter("?", cmbStatus.SelectedValue ?? DBNull.Value),
                new OleDbParameter("?", personnelID)
            };

            try
            {
                db.ExecuteNonQuery(query, parameters);

                // ⭐ ذخیره عکس جدید اگر انتخاب شده است
                if (!string.IsNullOrEmpty(selectedPhotoPath))
                {
                    string nationalID = txtNationalID.Text.Trim();
                    bool photoSaved = ImageHelper.SaveImage(selectedPhotoPath, nationalID);

                    if (!photoSaved)
                    {
                        MessageBox.Show("پرسنل به‌روزرسانی شد اما عکس ذخیره نشد.", "اخطار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                MessageBox.Show("رکورد پرسنل با موفقیت به‌روزرسانی شد!");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در به‌روزرسانی رکورد: " + ex.Message);
            }
        }
    }
}