using System;
using System.Drawing;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    public partial class FormFontSettings : Form
    {
        private ComboBox cmbFontFamily;
        private NumericUpDown nudTitleSize;
        private NumericUpDown nudLabelSize;
        private NumericUpDown nudTextBoxSize;
        private NumericUpDown nudButtonSize;
        private NumericUpDown nudBodySize;
        private NumericUpDown nudChartLabelSize;
        private CheckBox chkTitleBold;
        private CheckBox chkLabelBold;
        private CheckBox chkButtonBold;
        private CheckBox chkChartLabelBold;
        private Button btnSave;
        private Button btnCancel;
        private Button btnReset;
        private Label lblPreview;

        public FormFontSettings()
        {
            InitializeComponent();
            BuildUI();
            LoadCurrentSettings();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(700, 650);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormFontSettings";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.RightToLeftLayout = true;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "⚙️ تنظیمات فونت";
            this.BackColor = Color.FromArgb(240, 248, 255);
            this.ResumeLayout(false);
        }

        private void BuildUI()
        {
            int yPos = 20;
            int labelWidth = 150;
            int controlWidth = 200;
            int xLabel = 500;
            int xControl = 280;
            int rowHeight = 50;

            // عنوان فرم
            Label lblTitle = new Label
            {
                Text = "🎨 تنظیمات فونت برنامه",
                Location = new Point(20, yPos),
                Size = new Size(660, 35),
                Font = new Font(FontSettings.FontFamilyName, 14F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.FromArgb(230, 240, 250)
            };
            Controls.Add(lblTitle);
            yPos += 50;

            // انتخاب فونت
            Label lblFontFamily = new Label
            {
                Text = "🔤 نوع فونت:",
                Location = new Point(xLabel, yPos),
                Size = new Size(labelWidth, 25),
                Font = FontSettings.LabelFont,
                TextAlign = ContentAlignment.MiddleRight
            };
            Controls.Add(lblFontFamily);

            cmbFontFamily = new ComboBox
            {
                Location = new Point(xControl, yPos),
                Size = new Size(controlWidth, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = FontSettings.TextBoxFont
            };
            cmbFontFamily.Items.AddRange(FontSettings.GetPersianFonts());
            cmbFontFamily.SelectedIndexChanged += (s, e) => UpdatePreview();
            Controls.Add(cmbFontFamily);
            yPos += rowHeight;

            // اندازه فونت عنوان
            Label lblTitleSize = new Label
            {
                Text = "📏 اندازه عنوان:",
                Location = new Point(xLabel, yPos),
                Size = new Size(labelWidth, 25),
                Font = FontSettings.LabelFont,
                TextAlign = ContentAlignment.MiddleRight
            };
            Controls.Add(lblTitleSize);

            nudTitleSize = new NumericUpDown
            {
                Location = new Point(xControl, yPos),
                Size = new Size(100, 25),
                Minimum = 8,
                Maximum = 30,
                DecimalPlaces = 0,
                Font = FontSettings.TextBoxFont
            };
            nudTitleSize.ValueChanged += (s, e) => UpdatePreview();
            Controls.Add(nudTitleSize);

            chkTitleBold = new CheckBox
            {
                Text = "ضخیم",
                Location = new Point(xControl + 110, yPos),
                Size = new Size(80, 25),
                Font = FontSettings.LabelFont
            };
            chkTitleBold.CheckedChanged += (s, e) => UpdatePreview();
            Controls.Add(chkTitleBold);
            yPos += rowHeight;

            // اندازه فونت برچسب
            Label lblLabelSize = new Label
            {
                Text = "📋 اندازه برچسب:",
                Location = new Point(xLabel, yPos),
                Size = new Size(labelWidth, 25),
                Font = FontSettings.LabelFont,
                TextAlign = ContentAlignment.MiddleRight
            };
            Controls.Add(lblLabelSize);

            nudLabelSize = new NumericUpDown
            {
                Location = new Point(xControl, yPos),
                Size = new Size(100, 25),
                Minimum = 8,
                Maximum = 24,
                DecimalPlaces = 0,
                Font = FontSettings.TextBoxFont
            };
            nudLabelSize.ValueChanged += (s, e) => UpdatePreview();
            Controls.Add(nudLabelSize);

            chkLabelBold = new CheckBox
            {
                Text = "ضخیم",
                Location = new Point(xControl + 110, yPos),
                Size = new Size(80, 25),
                Font = FontSettings.LabelFont
            };
            chkLabelBold.CheckedChanged += (s, e) => UpdatePreview();
            Controls.Add(chkLabelBold);
            yPos += rowHeight;

            // اندازه فونت TextBox
            Label lblTextBoxSize = new Label
            {
                Text = "✍️ اندازه جعبه‌متن:",
                Location = new Point(xLabel, yPos),
                Size = new Size(labelWidth, 25),
                Font = FontSettings.LabelFont,
                TextAlign = ContentAlignment.MiddleRight
            };
            Controls.Add(lblTextBoxSize);

            nudTextBoxSize = new NumericUpDown
            {
                Location = new Point(xControl, yPos),
                Size = new Size(100, 25),
                Minimum = 8,
                Maximum = 24,
                DecimalPlaces = 0,
                Font = FontSettings.TextBoxFont
            };
            nudTextBoxSize.ValueChanged += (s, e) => UpdatePreview();
            Controls.Add(nudTextBoxSize);
            yPos += rowHeight;

            // اندازه فونت Button
            Label lblButtonSize = new Label
            {
                Text = "🔘 اندازه دکمه:",
                Location = new Point(xLabel, yPos),
                Size = new Size(labelWidth, 25),
                Font = FontSettings.LabelFont,
                TextAlign = ContentAlignment.MiddleRight
            };
            Controls.Add(lblButtonSize);

            nudButtonSize = new NumericUpDown
            {
                Location = new Point(xControl, yPos),
                Size = new Size(100, 25),
                Minimum = 8,
                Maximum = 24,
                DecimalPlaces = 0,
                Font = FontSettings.TextBoxFont
            };
            nudButtonSize.ValueChanged += (s, e) => UpdatePreview();
            Controls.Add(nudButtonSize);

            chkButtonBold = new CheckBox
            {
                Text = "ضخیم",
                Location = new Point(xControl + 110, yPos),
                Size = new Size(80, 25),
                Font = FontSettings.LabelFont
            };
            chkButtonBold.CheckedChanged += (s, e) => UpdatePreview();
            Controls.Add(chkButtonBold);
            yPos += rowHeight;

            // اندازه فونت Body
            Label lblBodySize = new Label
            {
                Text = "📄 اندازه متن عادی:",
                Location = new Point(xLabel, yPos),
                Size = new Size(labelWidth, 25),
                Font = FontSettings.LabelFont,
                TextAlign = ContentAlignment.MiddleRight
            };
            Controls.Add(lblBodySize);

            nudBodySize = new NumericUpDown
            {
                Location = new Point(xControl, yPos),
                Size = new Size(100, 25),
                Minimum = 8,
                Maximum = 20,
                DecimalPlaces = 0,
                Font = FontSettings.TextBoxFont
            };
            nudBodySize.ValueChanged += (s, e) => UpdatePreview();
            Controls.Add(nudBodySize);
            yPos += rowHeight;

            // اندازه فونت متن نمودار - **جدید**
            Label lblChartLabelSize = new Label
            {
                Text = "📊 اندازه متن نمودار:",
                Location = new Point(xLabel, yPos),
                Size = new Size(labelWidth, 25),
                Font = FontSettings.LabelFont,
                ForeColor = Color.FromArgb(0, 102, 204),
                TextAlign = ContentAlignment.MiddleRight
            };
            Controls.Add(lblChartLabelSize);

            nudChartLabelSize = new NumericUpDown
            {
                Location = new Point(xControl, yPos),
                Size = new Size(100, 25),
                Minimum = 7,
                Maximum = 20,
                DecimalPlaces = 0,
                Font = FontSettings.TextBoxFont
            };
            nudChartLabelSize.ValueChanged += (s, e) => UpdatePreview();
            Controls.Add(nudChartLabelSize);

            chkChartLabelBold = new CheckBox
            {
                Text = "ضخیم",
                Location = new Point(xControl + 110, yPos),
                Size = new Size(80, 25),
                Font = FontSettings.LabelFont
            };
            chkChartLabelBold.CheckedChanged += (s, e) => UpdatePreview();
            Controls.Add(chkChartLabelBold);
            yPos += rowHeight;

            // نمایش پیش‌نمایش
            Label lblPreviewTitle = new Label
            {
                Text = "👁️ پیش‌نمایش:",
                Location = new Point(20, yPos),
                Size = new Size(660, 25),
                Font = new Font(FontSettings.FontFamilyName, 11F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                TextAlign = ContentAlignment.MiddleCenter
            };
            Controls.Add(lblPreviewTitle);
            yPos += 35;

            lblPreview = new Label
            {
                Text = "این یک متن نمونه است برای پیش‌نمایش فونت\nاعداد: 1234567890\nEnglish: Sample Text",
                Location = new Point(50, yPos),
                Size = new Size(600, 80),
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.White
            };
            Controls.Add(lblPreview);
            yPos += 100;

            // دکمه‌های عملیات
            btnSave = new Button
            {
                Text = "💾 ذخیره",
                Location = new Point(500, yPos),
                Size = new Size(150, 40),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                Font = FontSettings.ButtonFont,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;
            Controls.Add(btnSave);

            btnReset = new Button
            {
                Text = "🔄 بازگشت به پیش‌فرض",
                Location = new Point(280, yPos),
                Size = new Size(200, 40),
                BackColor = Color.FromArgb(255, 193, 7),
                ForeColor = Color.Black,
                Font = FontSettings.ButtonFont,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnReset.FlatAppearance.BorderSize = 0;
            btnReset.Click += BtnReset_Click;
            Controls.Add(btnReset);

            btnCancel = new Button
            {
                Text = "❌ انصراف",
                Location = new Point(50, yPos),
                Size = new Size(150, 40),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                Font = FontSettings.ButtonFont,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(btnCancel);
        }

        private void LoadCurrentSettings()
        {
            cmbFontFamily.SelectedItem = FontSettings.FontFamilyName;
            nudTitleSize.Value = (decimal)FontSettings.TitleFontSize;
            nudLabelSize.Value = (decimal)FontSettings.LabelFontSize;
            nudTextBoxSize.Value = (decimal)FontSettings.TextBoxFontSize;
            nudButtonSize.Value = (decimal)FontSettings.ButtonFontSize;
            nudBodySize.Value = (decimal)FontSettings.BodyFontSize;
            nudChartLabelSize.Value = (decimal)FontSettings.ChartLabelFontSize;
            chkTitleBold.Checked = FontSettings.TitleFontBold;
            chkLabelBold.Checked = FontSettings.LabelFontBold;
            chkButtonBold.Checked = FontSettings.ButtonFontBold;
            chkChartLabelBold.Checked = FontSettings.ChartLabelFontBold;

            UpdatePreview();
        }

        private void UpdatePreview()
        {
            try
            {
                string fontName = cmbFontFamily.SelectedItem?.ToString() ?? "Tahoma";
                float size = (float)nudBodySize.Value;
                lblPreview.Font = new Font(fontName, size, FontStyle.Regular);
            }
            catch { }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                FontSettings.FontFamilyName = cmbFontFamily.SelectedItem?.ToString() ?? "Tahoma";
                FontSettings.TitleFontSize = (float)nudTitleSize.Value;
                FontSettings.LabelFontSize = (float)nudLabelSize.Value;
                FontSettings.TextBoxFontSize = (float)nudTextBoxSize.Value;
                FontSettings.ButtonFontSize = (float)nudButtonSize.Value;
                FontSettings.BodyFontSize = (float)nudBodySize.Value;
                FontSettings.ChartLabelFontSize = (float)nudChartLabelSize.Value;
                FontSettings.TitleFontBold = chkTitleBold.Checked;
                FontSettings.LabelFontBold = chkLabelBold.Checked;
                FontSettings.ButtonFontBold = chkButtonBold.Checked;
                FontSettings.ChartLabelFontBold = chkChartLabelBold.Checked;

                FontSettings.SaveSettings();

                MessageBox.Show(
                    "✅ تنظیمات فونت با موفقیت ذخیره شد.\n\nبرای اعمال تغییرات، برنامه را مجدداً راه‌اندازی کنید.",
                    "موفق",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در ذخیره تنظیمات: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnReset_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "⚠️ آیا مطمئن هستید که می‌خواهید تنظیمات را به حالت پیش‌فرض بازگردانید؟",
                "تأیید بازنشانی",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result == DialogResult.Yes)
            {
                FontSettings.ResetToDefaults();
                LoadCurrentSettings();
                MessageBox.Show("✅ تنظیمات به حالت پیش‌فرض بازگشت.", "موفق", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}