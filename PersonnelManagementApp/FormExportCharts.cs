using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace PersonnelManagementApp
{
    public partial class FormExportCharts : Form
    {
        private Panel pnlChartList;
        private Panel pnlPreview;
        private Chart previewChart;
        private RichTextBox txtStats;
        private ComboBox cmbChartType;
        private Button btnExportPDF;
        private Button btnExportImage;
        private Button btnPrint;

        // رنگ‌های مدرن
        private readonly Color PrimaryColor = Color.FromArgb(33, 150, 243);
        private readonly Color AccentColor = Color.FromArgb(76, 175, 80);
        private readonly Color WarningColor = Color.FromArgb(255, 152, 0);
        private readonly Color BackgroundColor = Color.FromArgb(250, 250, 250);
        private readonly Color CardBackground = Color.White;
        private readonly Color TextPrimary = Color.FromArgb(33, 33, 33);
        private readonly Color TextSecondary = Color.FromArgb(117, 117, 117);

        private string selectedChartType = "";

        public FormExportCharts()
        {
            InitializeComponent();
            FontSettings.ApplyFontToForm(this);
            LoadChartTypes();
        }

        private void InitializeComponent()
        {
            this.Text = "📊 خروجی نمودارها";
            this.Size = new Size(1200, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.RightToLeft = RightToLeft.Yes;
            this.BackColor = BackgroundColor;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // ========== Panel چپ: انتخاب نمودار ==========
            Panel leftPanel = new Panel
            {
                Location = new Point(20, 20),
                Size = new Size(300, 680),
                BackColor = CardBackground
            };
            this.Controls.Add(leftPanel);

            Label lblTitle = new Label
            {
                Text = "📋 انتخاب نمودار",
                Location = new Point(20, 20),
                Size = new Size(260, 35),
                Font = new Font(FontSettings.FontFamilyName, 14, FontStyle.Bold),
                ForeColor = PrimaryColor,
                TextAlign = ContentAlignment.MiddleRight
            };
            leftPanel.Controls.Add(lblTitle);

            Label lblDesc = new Label
            {
                Text = "نمودار مورد نظر را انتخاب کنید:",
                Location = new Point(20, 60),
                Size = new Size(260, 25),
                Font = new Font(FontSettings.FontFamilyName, 9),
                ForeColor = TextSecondary,
                TextAlign = ContentAlignment.MiddleRight
            };
            leftPanel.Controls.Add(lblDesc);

            cmbChartType = new ComboBox
            {
                Location = new Point(20, 95),
                Size = new Size(260, 30),
                Font = new Font(FontSettings.FontFamilyName, 10),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbChartType.SelectedIndexChanged += CmbChartType_SelectedIndexChanged;
            leftPanel.Controls.Add(cmbChartType);

            // لیست نمودارها
            pnlChartList = new Panel
            {
                Location = new Point(20, 140),
                Size = new Size(260, 430),
                BackColor = Color.FromArgb(248, 249, 250),
                AutoScroll = true
            };
            leftPanel.Controls.Add(pnlChartList);

            // دکمه‌های Export
            int btnY = 585;
            btnExportPDF = CreateActionButton("📄 خروجی PDF", 20, btnY, AccentColor);
            btnExportPDF.Click += BtnExportPDF_Click;
            leftPanel.Controls.Add(btnExportPDF);

            btnExportImage = CreateActionButton("🖼️ ذخیره عکس", 20, btnY + 40, PrimaryColor);
            btnExportImage.Click += BtnExportImage_Click;
            leftPanel.Controls.Add(btnExportImage);

            // ========== Panel راست: پیش‌نمایش ==========
            Panel rightPanel = new Panel
            {
                Location = new Point(340, 20),
                Size = new Size(840, 680),
                BackColor = CardBackground
            };
            this.Controls.Add(rightPanel);

            Label lblPreview = new Label
            {
                Text = "👁️ پیش‌نمایش",
                Location = new Point(20, 20),
                Size = new Size(800, 35),
                Font = new Font(FontSettings.FontFamilyName, 14, FontStyle.Bold),
                ForeColor = PrimaryColor,
                TextAlign = ContentAlignment.MiddleRight
            };
            rightPanel.Controls.Add(lblPreview);

            // نمودار پیش‌نمایش
            previewChart = new Chart
            {
                Location = new Point(20, 65),
                Size = new Size(520, 400),
                BackColor = Color.White
            };
            previewChart.ChartAreas.Add(new ChartArea("MainArea")
            {
                BackColor = Color.White
            });
            rightPanel.Controls.Add(previewChart);

            // آمار نمودار
            Label lblStats = new Label
            {
                Text = "📈 آمار نمودار:",
                Location = new Point(560, 65),
                Size = new Size(260, 30),
                Font = new Font(FontSettings.FontFamilyName, 11, FontStyle.Bold),
                ForeColor = TextPrimary,
                TextAlign = ContentAlignment.MiddleRight
            };
            rightPanel.Controls.Add(lblStats);

            txtStats = new RichTextBox
            {
                Location = new Point(560, 100),
                Size = new Size(260, 365),
                Font = new Font(FontSettings.FontFamilyName, 9),
                ReadOnly = true,
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };
            rightPanel.Controls.Add(txtStats);

            // دکمه چاپ
            btnPrint = CreateActionButton("🖨️ چاپ", 20, 485, WarningColor);
            btnPrint.Click += BtnPrint_Click;
            rightPanel.Controls.Add(btnPrint);

            // دکمه بستن
            Button btnClose = CreateActionButton("❌ بستن", 180, 485, Color.FromArgb(244, 67, 54));
            btnClose.Click += (s, e) => this.Close();
            rightPanel.Controls.Add(btnClose);
        }

        private Button CreateActionButton(string text, int x, int y, Color backColor)
        {
            Button btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(150, 45),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(FontSettings.FontFamilyName, 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void LoadChartTypes()
        {
            var chartTypes = new Dictionary<string, string>
            {
                { "department", "📊 نمودار اداره" },
                { "education", "🎓 نمودار تحصیلات" },
                { "employment", "💼 نمودار وضعیت استخدام" },
                { "jobtype", "👔 نمودار نوع شغل" },
                { "military", "🪖 نمودار وضعیت نظام وظیفه" },
                { "age", "📅 نمودار سنی" },
                { "gender", "👤 نمودار جنسیت" },
                { "marital", "💑 نمودار وضعیت تاهل" }
            };

            cmbChartType.Items.Clear();
            foreach (var item in chartTypes)
            {
                cmbChartType.Items.Add(item.Value);
            }

            if (cmbChartType.Items.Count > 0)
                cmbChartType.SelectedIndex = 0;
        }

        private void CmbChartType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbChartType.SelectedIndex < 0) return;

            string selected = cmbChartType.SelectedItem.ToString();
            
            if (selected.Contains("اداره"))
                selectedChartType = "department";
            else if (selected.Contains("تحصیلات"))
                selectedChartType = "education";
            else if (selected.Contains("استخدام"))
                selectedChartType = "employment";
            else if (selected.Contains("نوع شغل"))
                selectedChartType = "jobtype";
            else if (selected.Contains("نظام"))
                selectedChartType = "military";
            else if (selected.Contains("سنی"))
                selectedChartType = "age";
            else if (selected.Contains("جنسیت"))
                selectedChartType = "gender";
            else if (selected.Contains("تاهل"))
                selectedChartType = "marital";

            LoadChartPreview();
        }

        private void LoadChartPreview()
        {
            if (string.IsNullOrEmpty(selectedChartType)) return;

            try
            {
                previewChart.Series.Clear();
                previewChart.Titles.Clear();

                string query = "";
                string chartTitle = "";
                string fieldName = "";
                string displayName = "";

                switch (selectedChartType)
                {
                    case "department":
                        query = "SELECT [نام اداره], COUNT(*) as تعداد FROM Personnel GROUP BY [نام اداره]";
                        chartTitle = "توزیع پرسنل بر اساس اداره";
                        fieldName = "نام اداره";
                        displayName = "اداره";
                        break;
                    case "education":
                        query = "SELECT [مدرک تحصیلی], COUNT(*) as تعداد FROM Personnel GROUP BY [مدرک تحصیلی]";
                        chartTitle = "توزیع پرسنل بر اساس تحصیلات";
                        fieldName = "مدرک تحصیلی";
                        displayName = "تحصیلات";
                        break;
                    case "employment":
                        query = "SELECT [وضعیت استخدام], COUNT(*) as تعداد FROM Personnel GROUP BY [وضعیت استخدام]";
                        chartTitle = "توزیع پرسنل بر اساس وضعیت استخدام";
                        fieldName = "وضعیت استخدام";
                        displayName = "وضعیت";
                        break;
                    case "jobtype":
                        query = "SELECT [نوع شغل], COUNT(*) as تعداد FROM Personnel GROUP BY [نوع شغل]";
                        chartTitle = "توزیع پرسنل بر اساس نوع شغل";
                        fieldName = "نوع شغل";
                        displayName = "نوع شغل";
                        break;
                    case "military":
                        query = "SELECT [وضعیت نظام وظیفه], COUNT(*) as تعداد FROM Personnel GROUP BY [وضعیت نظام وظیفه]";
                        chartTitle = "توزیع پرسنل بر اساس وضعیت نظام وظیفه";
                        fieldName = "وضعیت نظام وظیفه";
                        displayName = "وضعیت";
                        break;
                    case "age":
                        query = @"SELECT 
                                    IIF(Age < 25, 'زیر 25 سال',
                                    IIF(Age >= 25 AND Age < 35, '25-34 سال',
                                    IIF(Age >= 35 AND Age < 45, '35-44 سال',
                                    IIF(Age >= 45 AND Age < 55, '45-54 سال', '55 سال به بالا')))) as [گروه سنی],
                                    COUNT(*) as تعداد
                                  FROM (SELECT YEAR(Date()) - YEAR([تاریخ تولد]) as Age FROM Personnel)
                                  GROUP BY [گروه سنی]";
                        chartTitle = "توزیع پرسنل بر اساس گروه‌های سنی";
                        fieldName = "گروه سنی";
                        displayName = "گروه سنی";
                        break;
                    case "gender":
                        query = "SELECT [جنسیت], COUNT(*) as تعداد FROM Personnel GROUP BY [جنسیت]";
                        chartTitle = "توزیع پرسنل بر اساس جنسیت";
                        fieldName = "جنسیت";
                        displayName = "جنسیت";
                        break;
                    case "marital":
                        query = "SELECT [وضعیت تاهل], COUNT(*) as تعداد FROM Personnel GROUP BY [وضعیت تاهل]";
                        chartTitle = "توزیع پرسنل بر اساس وضعیت تاهل";
                        fieldName = "وضعیت تاهل";
                        displayName = "وضعیت";
                        break;
                }

                using (OleDbConnection conn = new OleDbConnection(AppSettings.ConnectionString))
                {
                    conn.Open();
                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        Series series = new Series("Data")
                        {
                            ChartType = SeriesChartType.Column,
                            Font = FontSettings.ChartLabelFont,
                            IsValueShownAsLabel = true,
                            LabelForeColor = Color.Black
                        };

                        int totalCount = 0;
                        var statsData = new List<Tuple<string, int>>();

                        while (reader.Read())
                        {
                            string label = reader[0]?.ToString() ?? "نامشخص";
                            int count = Convert.ToInt32(reader[1]);
                            series.Points.AddXY(label, count);
                            totalCount += count;
                            statsData.Add(new Tuple<string, int>(label, count));
                        }

                        previewChart.Series.Add(series);
                        previewChart.Titles.Add(new Title(chartTitle)
                        {
                            Font = new Font(FontSettings.FontFamilyName, 12, FontStyle.Bold),
                            ForeColor = PrimaryColor
                        });

                        // نمایش آمار
                        DisplayStats(chartTitle, displayName, statsData, totalCount);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری نمودار:\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DisplayStats(string title, string categoryName, List<Tuple<string, int>> data, int total)
        {
            txtStats.Clear();
            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 11, FontStyle.Bold);
            txtStats.SelectionColor = PrimaryColor;
            txtStats.AppendText($"{title}\n\n");

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 9, FontStyle.Bold);
            txtStats.SelectionColor = TextPrimary;
            txtStats.AppendText($"📊 تعداد کل: {total} نفر\n\n");

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 9, FontStyle.Bold);
            txtStats.AppendText($"📋 تفکیک {categoryName}:\n\n");

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 9);
            foreach (var item in data.OrderByDescending(x => x.Item2))
            {
                double percentage = (double)item.Item2 / total * 100;
                txtStats.SelectionColor = TextSecondary;
                txtStats.AppendText($"• {item.Item1}:\n");
                txtStats.SelectionColor = AccentColor;
                txtStats.AppendText($"   {item.Item2} نفر ({percentage:F1}%)\n\n");
            }

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 8, FontStyle.Italic);
            txtStats.SelectionColor = TextSecondary;
            txtStats.AppendText($"\n\nتاریخ تولید: {DateTime.Now:yyyy/MM/dd - HH:mm}");
        }

        private void BtnExportPDF_Click(object sender, EventArgs e)
        {
            if (previewChart.Series.Count == 0)
            {
                MessageBox.Show("لطفاً ابتدا یک نمودار انتخاب کنید!", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            MessageBox.Show(
                "⚠️ برای تولید PDF نیاز به نصب کتابخانه‌های اضافی است.\n\n" +
                "در حال حاضر می‌توانید از 'ذخیره عکس' استفاده کنید.",
                "اطلاعات",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        private void BtnExportImage_Click(object sender, EventArgs e)
        {
            if (previewChart.Series.Count == 0)
            {
                MessageBox.Show("لطفاً ابتدا یک نمودار انتخاب کنید!", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "PNG Image|*.png|JPEG Image|*.jpg|BMP Image|*.bmp";
                    sfd.Title = "ذخیره نمودار به عنوان عکس";
                    sfd.FileName = $"نمودار_{selectedChartType}_{DateTime.Now:yyyyMMdd_HHmmss}";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        // ایجاد یک Bitmap بزرگ‌تر برای نمودار + آمار
                        int width = 1200;
                        int height = 800;
                        using (Bitmap bmp = new Bitmap(width, height))
                        using (Graphics g = Graphics.FromImage(bmp))
                        {
                            g.Clear(Color.White);

                            // رسم نمودار
                            previewChart.Printing.PrintPaint(g, new Rectangle(50, 50, 700, 600));

                            // رسم آمار
                            g.DrawString(txtStats.Text, new Font(FontSettings.FontFamilyName, 9), Brushes.Black, new RectangleF(780, 50, 380, 700));

                            // ذخیره
                            ImageFormat format = ImageFormat.Png;
                            if (sfd.FileName.EndsWith(".jpg"))
                                format = ImageFormat.Jpeg;
                            else if (sfd.FileName.EndsWith(".bmp"))
                                format = ImageFormat.Bmp;

                            bmp.Save(sfd.FileName, format);
                        }

                        MessageBox.Show("✅ نمودار با موفقیت ذخیره شد!", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در ذخیره عکس:\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (previewChart.Series.Count == 0)
            {
                MessageBox.Show("لطفاً ابتدا یک نمودار انتخاب کنید!", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                previewChart.Printing.Print(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در چاپ:\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}