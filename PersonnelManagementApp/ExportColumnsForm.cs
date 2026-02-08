using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    public partial class ExportColumnsForm : Form
    {
        public List<string> SelectedColumns { get; private set; }
        private CheckedListBox checkedListBox;
        private Button btnOK;
        private Button btnCancel;
        private Button btnSelectAll;
        private Button btnDeselectAll;

        private readonly Dictionary<string, string> columnMappings = new Dictionary<string, string>
        {
            { "PersonnelID", "🆔 شناسه" },
            { "FirstName", "👤 نام" },
            { "LastName", "👤 نام‌خانوادگی" },
            { "PersonnelNumber", "🔢 شماره پرسنلی" },
            { "NationalID", "🆔 کد ملی" },
            { "PostName", "💼 پست" },
            { "DeptName", "🏛️ اداره" },
            { "Province", "🗺️ استان" },
            { "City", "🏙️ شهر" },
            { "Affair", "📋 امور" },
            { "District", "🔺 ناحیه" },
            { "ContractType", "📄 نوع قرارداد" },
            { "HireDate", "📅 تاریخ استخدام" },
            { "MobileNumber", "📱 تلفن همراه" },
            { "Gender", "👥 جنسیت" },
            { "Education", "📚 تحصیلات" },
            { "JobLevel", "📊 سطح شغلی" },
            { "Company", "🏢 شرکت" },
            { "WorkShift", "⏰ شیفت کاری" },
            { "Salary", "💰 حقوق" },
            { "Email", "✉️ ایمیل" },
            { "BirthDate", "🎂 تاریخ تولد" },
            { "Address", "🏠 آدرس" }
        };

        public ExportColumnsForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "📊 انتخاب ستون‌های خروجی اکسل";
            Size = new Size(500, 700);
            StartPosition = FormStartPosition.CenterParent;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.FromArgb(240, 248, 255);
            Font = FontSettings.BodyFont;

            // ========== عنوان ==========
            Label lblTitle = new Label
            {
                Text = "لطفاً ستون‌هایی که می‌خواهید در فایل اکسل داشته باشید را انتخاب کنید:",
                Location = new Point(20, 20),
                Size = new Size(460, 50),
                Font = FontSettings.SubtitleFont,
                ForeColor = Color.FromArgb(0, 102, 204),
                TextAlign = ContentAlignment.TopRight
            };
            Controls.Add(lblTitle);

            // ========== CheckedListBox ==========
            checkedListBox = new CheckedListBox
            {
                Location = new Point(20, 80),
                Size = new Size(460, 450),
                CheckOnClick = true,
                Font = FontSettings.BodyFont,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            // اضافه کردن آیتم‌ها
            foreach (var item in columnMappings)
            {
                checkedListBox.Items.Add(new ColumnItem { Key = item.Key, Display = item.Value });
                // به صورت پیش‌فرض فیلدهای مهم انتخاب شده‌اند
                if (IsImportantField(item.Key))
                {
                    checkedListBox.SetItemChecked(checkedListBox.Items.Count - 1, true);
                }
            }

            checkedListBox.DisplayMember = "Display";
            Controls.Add(checkedListBox);

            // ========== دکمه‌های انتخاب همه / هیچ‌کدام ==========
            btnSelectAll = new Button
            {
                Text = "✅ انتخاب همه",
                Location = new Point(20, 545),
                Size = new Size(220, 40),
                BackColor = Color.FromArgb(0, 123, 255),
                ForeColor = Color.White,
                Font = FontSettings.ButtonFont,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnSelectAll.FlatAppearance.BorderSize = 0;
            btnSelectAll.Click += BtnSelectAll_Click;
            Controls.Add(btnSelectAll);

            btnDeselectAll = new Button
            {
                Text = "❌ حذف انتخاب همه",
                Location = new Point(260, 545),
                Size = new Size(220, 40),
                BackColor = Color.FromArgb(108, 117, 125),
                ForeColor = Color.White,
                Font = FontSettings.ButtonFont,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnDeselectAll.FlatAppearance.BorderSize = 0;
            btnDeselectAll.Click += BtnDeselectAll_Click;
            Controls.Add(btnDeselectAll);

            // ========== دکمه‌های تایید / لغو ==========
            btnOK = new Button
            {
                Text = "✅ تایید و خروجی گرفتن",
                Location = new Point(20, 600),
                Size = new Size(220, 50),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                Font = FontSettings.ButtonFont,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.OK
            };
            btnOK.FlatAppearance.BorderSize = 0;
            btnOK.Click += BtnOK_Click;
            Controls.Add(btnOK);

            btnCancel = new Button
            {
                Text = "❌ لغو",
                Location = new Point(260, 600),
                Size = new Size(220, 50),
                BackColor = Color.FromArgb(220, 53, 69),
                ForeColor = Color.White,
                Font = FontSettings.ButtonFont,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            Controls.Add(btnCancel);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        private bool IsImportantField(string fieldName)
        {
            // فیلدهایی که به صورت پیش‌فرض انتخاب می‌شوند
            return fieldName == "FirstName" || fieldName == "LastName" ||
                   fieldName == "PersonnelNumber" || fieldName == "NationalID" ||
                   fieldName == "PostName" || fieldName == "DeptName" ||
                   fieldName == "Province" || fieldName == "MobileNumber";
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                checkedListBox.SetItemChecked(i, true);
            }
        }

        private void BtnDeselectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox.Items.Count; i++)
            {
                checkedListBox.SetItemChecked(i, false);
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            SelectedColumns = new List<string>();

            foreach (var item in checkedListBox.CheckedItems)
            {
                if (item is ColumnItem columnItem)
                {
                    SelectedColumns.Add(columnItem.Key);
                }
            }

            if (SelectedColumns.Count == 0)
            {
                MessageBox.Show("❌ لطفاً حداقل یک ستون را انتخاب کنید!",
                    "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        // کلاس برای نگهداری اطلاعات ستون
        private class ColumnItem
        {
            public string Key { get; set; }
            public string Display { get; set; }

            public override string ToString()
            {
                return Display;
            }
        }
    }
}