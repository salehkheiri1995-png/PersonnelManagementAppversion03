using System;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    public partial class FormPersonnelDelete : Form
    {
        private DbHelper db = new DbHelper();
        private ComboBox cbPersonnel;

        public FormPersonnelDelete()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "حذف پرسنل";
            this.WindowState = FormWindowState.Normal;
            this.RightToLeft = RightToLeft.Yes;
            this.BackColor = Color.FromArgb(240, 248, 255);
            this.Size = new Size(400, 200);

            // پس‌زمینه گرادیانت
            using (LinearGradientBrush brush = new LinearGradientBrush(this.ClientRectangle, Color.LightBlue, Color.White, LinearGradientMode.Vertical))
            {
                this.BackgroundImage = new Bitmap(this.Width, this.Height);
                using (Graphics g = Graphics.FromImage(this.BackgroundImage))
                {
                    g.FillRectangle(brush, this.ClientRectangle);
                }
            }

            // انتخاب پرسنل
            Label lblPersonnel = new Label { Text = "انتخاب پرسنل:", Location = new Point(50, 50), Size = new Size(100, 20), Font = new Font("Tahoma", 10) };
            cbPersonnel = new ComboBox { Location = new Point(160, 50), Size = new Size(200, 20), DropDownStyle = ComboBoxStyle.DropDownList };

            // دکمه حذف
            Button btnDelete = new Button
            {
                Text = "حذف",
                Location = new Point(160, 100),
                Size = new Size(100, 30),
                Font = new Font("Tahoma", 10),
                BackColor = Color.LightCoral,
                ForeColor = Color.White
            };
            btnDelete.Click += BtnDelete_Click;

            // دکمه بازگشت
            Button btnBack = new Button
            {
                Text = "بازگشت",
                Location = new Point(270, 100),
                Size = new Size(100, 30),
                Font = new Font("Tahoma", 10),
                BackColor = Color.LightGray,
                ForeColor = Color.Black
            };
            btnBack.Click += (s, e) => { this.Close(); };

            this.Controls.Add(lblPersonnel);
            this.Controls.Add(cbPersonnel);
            this.Controls.Add(btnDelete);
            this.Controls.Add(btnBack);

            LoadPersonnelList();
        }

        private void LoadPersonnelList()
        {
            cbPersonnel.DataSource = db.ExecuteQuery("SELECT PersonnelID, FirstName + ' ' + LastName AS FullName FROM Personnel").DefaultView;
            cbPersonnel.DisplayMember = "FullName";
            cbPersonnel.ValueMember = "PersonnelID";
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (cbPersonnel.SelectedIndex >= 0)
            {
                if (MessageBox.Show("آیا مطمئن هستید که می‌خواهید این پرسنل را حذف کنید؟", "تأیید حذف", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    try
                    {
                        int personnelId = (int)cbPersonnel.SelectedValue;
                        string query = "DELETE FROM Personnel WHERE PersonnelID = ?";
                        OleDbParameter[] parameters = new OleDbParameter[]
                        {
                            new OleDbParameter("?", personnelId)
                        };
                        int rowsAffected = db.ExecuteNonQuery(query, parameters);
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("پرسنل با موفقیت حذف شد!", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadPersonnelList();
                        }
                        else
                        {
                            MessageBox.Show("هیچ پرسنلی حذف نشد!", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("خطا در حذف پرسنل: " + ex.Message, "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("لطفاً یک پرسنل انتخاب کنید.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}