using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace PersonnelManagementApp
{
    public partial class FormExportCharts : Form
    {
        private ComboBox cmbChartType = null!;
        private Chart previewChart = null!;
        private RichTextBox txtStats = null!;
        private Button btnExportPDF = null!;
        private Button btnExportImage = null!;
        private Button btnPrint = null!;

        private readonly Color PrimaryColor = Color.FromArgb(33, 150, 243);
        private readonly Color AccentColor = Color.FromArgb(76, 175, 80);
        private readonly Color WarningColor = Color.FromArgb(255, 152, 0);
        private readonly Color BackgroundColor = Color.FromArgb(250, 250, 250);
        private readonly Color CardBackground = Color.White;
        private readonly Color TextPrimary = Color.FromArgb(33, 33, 33);
        private readonly Color TextSecondary = Color.FromArgb(117, 117, 117);

        private readonly DbHelper dbHelper;
        private readonly AnalyticsDataModel analyticsModel;

        public FormExportCharts()
        {
            dbHelper = new DbHelper();
            analyticsModel = new AnalyticsDataModel();
            
            InitializeComponent();
            FontSettings.ApplyFontToForm(this);
            LoadAnalyticsData();
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

            // ========== Panel چپ ==========
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

            // دکمه‌های Export
            btnExportImage = CreateActionButton("🖼️ ذخیره عکس", 20, 200, PrimaryColor);
            btnExportImage.Click += BtnExportImage_Click;
            leftPanel.Controls.Add(btnExportImage);

            btnPrint = CreateActionButton("🖨️ چاپ نمودار", 20, 260, WarningColor);
            btnPrint.Click += BtnPrint_Click;
            leftPanel.Controls.Add(btnPrint);

            Button btnClose = CreateActionButton("❌ بستن", 20, 320, Color.FromArgb(244, 67, 54));
            btnClose.Click += (s, e) => this.Close();
            leftPanel.Controls.Add(btnClose);

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
                Text = "👁️ پیش‌نمایش نمودار",
                Location = new Point(20, 20),
                Size = new Size(800, 35),
                Font = new Font(FontSettings.FontFamilyName, 14, FontStyle.Bold),
                ForeColor = PrimaryColor,
                TextAlign = ContentAlignment.MiddleRight
            };
            rightPanel.Controls.Add(lblPreview);

            previewChart = new Chart
            {
                Location = new Point(20, 65),
                Size = new Size(500, 400),
                BackColor = Color.White
            };
            previewChart.ChartAreas.Add(new ChartArea("MainArea")
            {
                BackColor = Color.White,
                Area3DStyle = { Enable3D = true, Inclination = 15, Rotation = 45 }
            });
            rightPanel.Controls.Add(previewChart);

            Label lblStats = new Label
            {
                Text = "📈 آمار نمودار:",
                Location = new Point(540, 65),
                Size = new Size(280, 30),
                Font = new Font(FontSettings.FontFamilyName, 11, FontStyle.Bold),
                ForeColor = TextPrimary,
                TextAlign = ContentAlignment.MiddleRight
            };
            rightPanel.Controls.Add(lblStats);

            txtStats = new RichTextBox
            {
                Location = new Point(540, 100),
                Size = new Size(280, 365),
                Font = new Font(FontSettings.FontFamilyName, 9),
                ReadOnly = true,
                BackColor = Color.FromArgb(248, 249, 250),
                BorderStyle = BorderStyle.FixedSingle
            };
            rightPanel.Controls.Add(txtStats);
        }

        private Button CreateActionButton(string text, int x, int y, Color backColor)
        {
            Button btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(260, 45),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font(FontSettings.FontFamilyName, 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void LoadAnalyticsData()
        {
            try
            {
                if (!dbHelper.TestConnection())
                {
                    MessageBox.Show("❌ اتصال به دیتابیس ناموفق بود.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!analyticsModel.LoadData(dbHelper))
                {
                    MessageBox.Show("❌ خطا در بارگذاری داده‌ها.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadChartTypes()
        {
            var chartTypes = new Dictionary<string, string>
            {
                { "department", "📊 نمودار ادارات" },
                { "position", "💼 نمودار پستها" },
                { "gender", "👥 نمودار جنسیت" },
                { "joblevel", "📈 نمودار سطح شغلی" },
                { "contract", "📋 نمودار نوع قرارداد" },
                { "province", "🗺️ نمودار استان" },
                { "education", "📚 نمودار تحصیلات" },
                { "company", "🏢 نمودار شرکت" },
                { "workshift", "⏰ نمودار شیفت کاری" },
                { "age", "🎂 نمودار سن" },
                { "experience", "💼 نمودار سابقه کاری" }
            };

            cmbChartType.Items.Clear();
            foreach (var item in chartTypes)
            {
                cmbChartType.Items.Add(item.Value);
            }

            if (cmbChartType.Items.Count > 0)
                cmbChartType.SelectedIndex = 0;
        }

        private void CmbChartType_SelectedIndexChanged(object? sender, EventArgs e)
        {
            LoadChartPreview();
        }

        private void LoadChartPreview()
        {
            if (cmbChartType.SelectedIndex < 0) return;

            try
            {
                previewChart.Series.Clear();
                previewChart.Titles.Clear();

                string? selected = cmbChartType.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(selected)) return;

                List<StatisticItem> stats = new List<StatisticItem>();
                string chartTitle = "";

                if (selected.Contains("ادارات"))
                {
                    stats = analyticsModel.GetFilteredDepartmentStatistics();
                    chartTitle = "📊 توزیع پرسنل در ادارهها";
                }
                else if (selected.Contains("پستها"))
                {
                    stats = analyticsModel.GetFilteredPositionStatistics();
                    chartTitle = "💼 توزیع پستهای شغلی";
                }
                else if (selected.Contains("جنسیت"))
                {
                    stats = analyticsModel.GetFilteredGenderStatistics();
                    chartTitle = "👥 توزیع جنسیت";
                }
                else if (selected.Contains("سطح شغلی"))
                {
                    stats = analyticsModel.GetFilteredJobLevelStatistics();
                    chartTitle = "📈 توزیع سطح شغلی";
                }
                else if (selected.Contains("قرارداد"))
                {
                    stats = analyticsModel.GetFilteredContractTypeStatistics();
                    chartTitle = "📋 توزیع نوع قرارداد";
                }
                else if (selected.Contains("استان"))
                {
                    stats = analyticsModel.GetFilteredProvinceStatistics();
                    chartTitle = "🗺️ توزیع بر اساس استان";
                }
                else if (selected.Contains("تحصیلات"))
                {
                    stats = analyticsModel.GetFilteredEducationStatistics();
                    chartTitle = "📚 توزیع مدارک تحصیلی";
                }
                else if (selected.Contains("شرکت"))
                {
                    stats = analyticsModel.GetFilteredCompanyStatistics();
                    chartTitle = "🏢 توزیع شرکتها";
                }
                else if (selected.Contains("شیفت"))
                {
                    stats = analyticsModel.GetFilteredWorkShiftStatistics();
                    chartTitle = "⏰ توزیع شیفت‌های کاری";
                }
                else if (selected.Contains("سن"))
                {
                    stats = analyticsModel.GetFilteredAgeStatistics(10);
                    chartTitle = "🎂 توزیع بر اساس سن";
                }
                else if (selected.Contains("سابقه"))
                {
                    stats = analyticsModel.GetFilteredWorkExperienceStatistics();
                    chartTitle = "💼 توزیع بر اساس سابقه کاری";
                }

                if (stats.Count == 0)
                {
                    MessageBox.Show("❌ داده‌ای برای نمایش وجود ندارد.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // ✅ ساخت نمودار
                Series series = new Series("تعداد")
                {
                    ChartType = SeriesChartType.Pie,
                    Font = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F),
                    IsValueShownAsLabel = true,
                    LabelForeColor = Color.Black
                };
                series["PieLabelStyle"] = "Outside";

                int total = stats.Sum(x => x.Count);
                var displayStats = stats.Take(15).ToList(); // فقط 15 تای اول

                foreach (var item in displayStats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر";
                }

                previewChart.Series.Add(series);
                previewChart.Titles.Add(new Title(chartTitle)
                {
                    Font = FontSettings.HeaderFont ?? new Font("Tahoma", 12F, FontStyle.Bold),
                    ForeColor = PrimaryColor
                });

                // نمایش آمار
                DisplayStats(chartTitle, stats, total);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری نمودار:\n{ex.Message}\n\n{ex.StackTrace}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DisplayStats(string title, List<StatisticItem> data, int total)
        {
            txtStats.Clear();
            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 11, FontStyle.Bold);
            txtStats.SelectionColor = PrimaryColor;
            txtStats.AppendText($"{title}\n\n");

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 9, FontStyle.Bold);
            txtStats.SelectionColor = TextPrimary;
            txtStats.AppendText($"📊 تعداد کل: {total} نفر\n\n");

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 9, FontStyle.Bold);
            txtStats.AppendText($"📋 تفکیک:\n\n");

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 9);
            foreach (var item in data.Take(20)) // فقط 20 تای اول
            {
                double percentage = total > 0 ? (double)item.Count / total * 100 : 0;
                txtStats.SelectionColor = TextSecondary;
                txtStats.AppendText($"• {item.Name}:\n");
                txtStats.SelectionColor = AccentColor;
                txtStats.AppendText($"   {item.Count} نفر ({percentage:F1}%)\n\n");
            }

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 8, FontStyle.Italic);
            txtStats.SelectionColor = TextSecondary;
            txtStats.AppendText($"\n\nتاریخ تولید: {DateTime.Now:yyyy/MM/dd - HH:mm}");
        }

        private void BtnExportImage_Click(object? sender, EventArgs e)
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
                    sfd.FileName = $"نمودار_{DateTime.Now:yyyyMMdd_HHmmss}";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
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

        private void BtnPrint_Click(object? sender, EventArgs e)
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