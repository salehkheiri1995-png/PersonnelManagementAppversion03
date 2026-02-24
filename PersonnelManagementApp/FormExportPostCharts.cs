using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace PersonnelManagementApp
{
    public partial class FormExportPostCharts : Form
    {
        private ComboBox cmbChartType = null!;
        private Chart previewChart = null!;
        private RichTextBox txtStats = null!;
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
        private DataTable? allPostsData;

        public FormExportPostCharts()
        {
            dbHelper = new DbHelper();
            
            InitializeComponent();
            FontSettings.ApplyFontToForm(this);
            LoadPostsData();
            LoadChartTypes();
        }

        private void InitializeComponent()
        {
            this.Text = "📊 خروجی نمودارهای پست‌ها";
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
                Text = "📋 انتخاب نمودار پست",
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

        private void LoadPostsData()
        {
            try
            {
                if (!dbHelper.TestConnection())
                {
                    MessageBox.Show("❌ اتصال به دیتابیس ناموفق بود.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string query = @"
                    SELECT Posts.PostID, Posts.OperationYear, Posts.DistributedCapacity, 
                    Posts.CapacityHV, Posts.CapacityMV, 
                    Provinces.ProvinceName, Cities.CityName, TransferAffairs.AffairName, 
                    OperationDepartments.DeptName, Districts.DistrictName, PostsNames.PostName, 
                    VoltageLevels.VoltageName, PostStandards.StandardName, PostTypes.TypeName, 
                    DistributedConnections.ConnName, InsulationTypes.InsName, PostTypeTwos.PT2Name, 
                    FixedMobiles.FMName, CircuitStatuses.CircuitName, DieselGenerators.DieselName, 
                    DistributionFeeds.FeedName, WaterStatuses.WaterName, GuestHouses.GuestName 
                    FROM (((((((((((((((((Posts 
                    INNER JOIN Provinces ON Posts.ProvinceID = Provinces.ProvinceID)
                    INNER JOIN Cities ON Posts.CityID = Cities.CityID)
                    INNER JOIN TransferAffairs ON Posts.AffairID = TransferAffairs.AffairID)
                    INNER JOIN OperationDepartments ON Posts.DeptID = OperationDepartments.DeptID)
                    INNER JOIN Districts ON Posts.DistrictID = Districts.DistrictID)
                    INNER JOIN PostsNames ON Posts.PostNameID = PostsNames.PostNameID)
                    INNER JOIN VoltageLevels ON Posts.VoltageID = VoltageLevels.VoltageID)
                    INNER JOIN PostStandards ON Posts.StandardID = PostStandards.StandardID)
                    INNER JOIN PostTypes ON Posts.TypeID = PostTypes.TypeID)
                    INNER JOIN DistributedConnections ON Posts.ConnID = DistributedConnections.ConnID)
                    INNER JOIN InsulationTypes ON Posts.InsID = InsulationTypes.InsID)
                    INNER JOIN PostTypeTwos ON Posts.PT2ID = PostTypeTwos.PT2ID)
                    INNER JOIN FixedMobiles ON Posts.FMID = FixedMobiles.FMID)
                    INNER JOIN CircuitStatuses ON Posts.CircuitID = CircuitStatuses.CircuitID)
                    INNER JOIN DieselGenerators ON Posts.DieselID = DieselGenerators.DieselID)
                    INNER JOIN DistributionFeeds ON Posts.FeedID = DistributionFeeds.FeedID)
                    INNER JOIN WaterStatuses ON Posts.WaterID = WaterStatuses.WaterID)
                    INNER JOIN GuestHouses ON Posts.GuestID = GuestHouses.GuestID
                ";

                allPostsData = dbHelper.ExecuteQuery(query);

                if (allPostsData == null || allPostsData.Rows.Count == 0)
                {
                    MessageBox.Show("⚠️ داده‌ای در جدول پست‌ها یافت نشد.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در بارگذاری داده‌ها: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadChartTypes()
        {
            var chartTypes = new Dictionary<string, string>
            {
                { "province", "🗺️ نمودار استان" },
                { "department", "🏛️ نمودار ادارات" },
                { "voltage", "⚡ نمودار سطح ولتاژ" },
                { "type", "🏗️ نمودار نوع پست" },
                { "standard", "📐 نمودار استاندارد" },
                { "circuit", "🔌 نمودار وضعیت مدار" },
                { "fixedmobile", "🚗 نمودار ثابت/سیار" },
                { "connection", "🔗 نمودار اتصال توزیع" },
                { "insulation", "🔆 نمودار نوع عایق" },
                { "posttype2", "📋 نمودار نوع پست ۲" },
                { "diesel", "🔋 نمودار دیزل ژنراتور" },
                { "operationyear", "📅 نمودار سال بهره‌برداری" }
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
            if (cmbChartType.SelectedIndex < 0 || allPostsData == null || allPostsData.Rows.Count == 0) return;

            try
            {
                previewChart.Series.Clear();
                previewChart.Titles.Clear();

                string? selected = cmbChartType.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(selected)) return;

                List<(string Name, int Count)> stats = new List<(string, int)>();
                string chartTitle = "";
                string columnName = "";

                if (selected.Contains("استان"))
                {
                    columnName = "ProvinceName";
                    chartTitle = "🗺️ توزیع پست‌ها بر اساس استان";
                }
                else if (selected.Contains("ادارات"))
                {
                    columnName = "DeptName";
                    chartTitle = "🏛️ توزیع پست‌ها بر اساس اداره";
                }
                else if (selected.Contains("ولتاژ"))
                {
                    columnName = "VoltageName";
                    chartTitle = "⚡ توزیع بر اساس سطح ولتاژ";
                }
                else if (selected.Contains("نوع پست") && !selected.Contains("۲"))
                {
                    columnName = "TypeName";
                    chartTitle = "🏗️ توزیع بر اساس نوع پست";
                }
                else if (selected.Contains("استاندارد"))
                {
                    columnName = "StandardName";
                    chartTitle = "📐 توزیع بر اساس استاندارد پست";
                }
                else if (selected.Contains("مدار"))
                {
                    columnName = "CircuitName";
                    chartTitle = "🔌 توزیع بر اساس وضعیت مدار";
                }
                else if (selected.Contains("ثابت"))
                {
                    columnName = "FMName";
                    chartTitle = "🚗 توزیع بر اساس ثابت / سیار";
                }
                else if (selected.Contains("اتصال"))
                {
                    columnName = "ConnName";
                    chartTitle = "🔗 توزیع بر اساس اتصال توزیع";
                }
                else if (selected.Contains("عایق"))
                {
                    columnName = "InsName";
                    chartTitle = "🔆 توزیع بر اساس نوع عایق";
                }
                else if (selected.Contains("نوع پست ۲"))
                {
                    columnName = "PT2Name";
                    chartTitle = "📋 توزیع بر اساس نوع پست ۲";
                }
                else if (selected.Contains("دیزل"))
                {
                    columnName = "DieselName";
                    chartTitle = "🔋 توزیع بر اساس دیزل ژنراتور";
                }
                else if (selected.Contains("سال"))
                {
                    DrawOperationYearChart();
                    return;
                }

                if (!string.IsNullOrEmpty(columnName))
                {
                    stats = allPostsData.AsEnumerable()
                        .GroupBy(r => r[columnName]?.ToString() ?? "نامشخص")
                        .Select(g => (Name: g.Key, Count: g.Count()))
                        .OrderByDescending(x => x.Count)
                        .ToList();
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
                var displayStats = stats.Take(15).ToList();

                foreach (var item in displayStats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} پست";
                }

                previewChart.Series.Add(series);
                previewChart.Titles.Add(new Title(chartTitle)
                {
                    Font = FontSettings.HeaderFont ?? new Font("Tahoma", 12F, FontStyle.Bold),
                    ForeColor = PrimaryColor
                });

                DisplayStats(chartTitle, stats, total);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در بارگذاری نمودار:\n{ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DrawOperationYearChart()
        {
            if (allPostsData == null) return;

            try
            {
                previewChart.Series.Clear();
                previewChart.Titles.Clear();

                var stats = allPostsData.AsEnumerable()
                    .Where(r => r["OperationYear"] != DBNull.Value)
                    .GroupBy(r =>
                    {
                        if (int.TryParse(r["OperationYear"]?.ToString(), out int y))
                            return $"{(y / 10) * 10}–{(y / 10) * 10 + 9}";
                        return "نامشخص";
                    })
                    .Select(g => (Name: g.Key, Count: g.Count()))
                    .OrderBy(x => x.Name)
                    .ToList();

                int total = stats.Sum(x => x.Count);

                Series series = new Series("تعداد")
                {
                    ChartType = SeriesChartType.Pie,
                    Font = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F),
                    IsValueShownAsLabel = true,
                    LabelForeColor = Color.Black
                };
                series["PieLabelStyle"] = "Outside";

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} پست";
                }

                previewChart.Series.Add(series);
                previewChart.Titles.Add(new Title("📅 توزیع بر اساس دهه بهره‌برداری")
                {
                    Font = FontSettings.HeaderFont ?? new Font("Tahoma", 12F, FontStyle.Bold),
                    ForeColor = PrimaryColor
                });

                DisplayStats("📅 توزیع بر اساس دهه بهره‌برداری", stats, total);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در نمودار سال: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DisplayStats(string title, List<(string Name, int Count)> data, int total)
        {
            txtStats.Clear();
            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 11, FontStyle.Bold);
            txtStats.SelectionColor = PrimaryColor;
            txtStats.AppendText($"{title}\n\n");

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 9, FontStyle.Bold);
            txtStats.SelectionColor = TextPrimary;
            txtStats.AppendText($"📊 تعداد کل: {total} پست\n\n");

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 9, FontStyle.Bold);
            txtStats.AppendText($"📋 تفکیک:\n\n");

            txtStats.SelectionFont = new Font(FontSettings.FontFamilyName, 9);
            foreach (var item in data.Take(20))
            {
                double percentage = total > 0 ? (double)item.Count / total * 100 : 0;
                txtStats.SelectionColor = TextSecondary;
                txtStats.AppendText($"• {item.Name}:\n");
                txtStats.SelectionColor = AccentColor;
                txtStats.AppendText($"   {item.Count} پست ({percentage:F1}%)\n\n");
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
                    sfd.FileName = $"نمودار_پست_{DateTime.Now:yyyyMMdd_HHmmss}";

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
