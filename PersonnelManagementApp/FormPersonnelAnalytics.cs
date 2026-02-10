using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace PersonnelManagementApp
{
    public partial class FormPersonnelAnalytics : Form
    {
        private readonly DbHelper dbHelper;
        private readonly TabControl tabControl;
        private readonly AnalyticsDataModel analyticsModel;

        // تمام نمودارها
        private readonly Chart chartDepartmentPie;
        private readonly Chart chartPositionPie;
        private readonly Chart chartGenderPie;
        private readonly Chart chartJobLevelPie;
        private readonly Chart chartContractTypePie;
        private readonly Chart chartProvincePie;
        private readonly Chart chartEducationPie;
        private readonly Chart chartCompanyPie;
        private readonly Chart chartWorkShiftPie;

        private readonly DataGridView dgvPersonnelStats;
        private readonly DataGridView dgvDepartmentDetails;
        private readonly DataGridView dgvPositionDetails;

        // فیلترها
        private readonly CheckedListBox clbProvincesFilter;
        private readonly CheckedListBox clbCitiesFilter;
        private readonly CheckedListBox clbAffairsFilter;
        private readonly CheckedListBox clbDepartmentsFilter;
        private readonly CheckedListBox clbDistrictsFilter;
        private readonly CheckedListBox clbPositionsFilter;
        private readonly CheckedListBox clbEducationFilter;
        private readonly CheckedListBox clbJobLevelFilter;
        private readonly CheckedListBox clbContractTypeFilter;
        private readonly CheckedListBox clbCompanyFilter;
        private readonly CheckedListBox clbWorkShiftFilter;
        private readonly CheckedListBox clbGenderFilter;

        private readonly Button btnClearFilters;
        private readonly Label lblFilterInfo;

        // فیلتر تاریخ استخدام
        private DateTimePicker dtpHireDateFrom;
        private DateTimePicker dtpHireDateTo;
        private CheckBox chkHireDateFilter;

        // رادیو باتن‌ها برای انتخاب نمایش
        private RadioButton rbShowSummary;
        private RadioButton rbShowFullStats;

        public FormPersonnelAnalytics()
        {
            dbHelper = new DbHelper();
            analyticsModel = new AnalyticsDataModel();
            tabControl = new TabControl();

            chartDepartmentPie = new Chart();
            chartPositionPie = new Chart();
            chartGenderPie = new Chart();
            chartJobLevelPie = new Chart();
            chartContractTypePie = new Chart();
            chartProvincePie = new Chart();
            chartEducationPie = new Chart();
            chartCompanyPie = new Chart();
            chartWorkShiftPie = new Chart();

            dgvPersonnelStats = new DataGridView();
            dgvDepartmentDetails = new DataGridView();
            dgvPositionDetails = new DataGridView();

            clbProvincesFilter = new CheckedListBox();
            clbCitiesFilter = new CheckedListBox();
            clbAffairsFilter = new CheckedListBox();
            clbDepartmentsFilter = new CheckedListBox();
            clbDistrictsFilter = new CheckedListBox();
            clbPositionsFilter = new CheckedListBox();
            clbEducationFilter = new CheckedListBox();
            clbJobLevelFilter = new CheckedListBox();
            clbContractTypeFilter = new CheckedListBox();
            clbCompanyFilter = new CheckedListBox();
            clbWorkShiftFilter = new CheckedListBox();
            clbGenderFilter = new CheckedListBox();

            btnClearFilters = new Button();
            lblFilterInfo = new Label();

            rbShowSummary = new RadioButton();
            rbShowFullStats = new RadioButton();

            InitializeComponent();
            BuildUI();

            // اعمال فونت مرکزی
            FontSettings.ApplyFontToForm(this);

            LoadData();
        }

        private void BuildUI()
        {
            Text = "🎯 تحلیل دادههای پرسنل - سیستم پیشرفته";
            WindowState = FormWindowState.Maximized;
            RightToLeft = RightToLeft.Yes;
            BackColor = Color.FromArgb(240, 248, 255);
            MinimumSize = new Size(1200, 700);
            Font = FontSettings.BodyFont;

            // ========== پنل فیلتر (دو ردیف) ==========
            Panel panelFilter = new Panel
            {
                Dock = DockStyle.Top,
                Height = 330,
                BackColor = Color.FromArgb(230, 240, 250),
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = false,
                Padding = new Padding(8, 8, 8, 6)
            };

            // جدول 2 ردیفه برای فیلترها (6 ستون در هر ردیف)
            TableLayoutPanel filterGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 245,
                ColumnCount = 6,
                RowCount = 2,
                RightToLeft = RightToLeft.Yes,
                BackColor = Color.Transparent
            };
            for (int i = 0; i < 6; i++)
                filterGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16.66f));
            filterGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
            filterGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));

            // ردیف اول
            filterGrid.Controls.Add(CreateFilterBox("استانها 🗺️", clbProvincesFilter, ClbProvincesFilter_ItemCheck), 0, 0);
            filterGrid.Controls.Add(CreateFilterBox("شهرها 🏙️", clbCitiesFilter, ClbCitiesFilter_ItemCheck), 1, 0);
            filterGrid.Controls.Add(CreateFilterBox("امور 📋", clbAffairsFilter, ClbAffairsFilter_ItemCheck), 2, 0);
            filterGrid.Controls.Add(CreateFilterBox("ادارات 🏛️", clbDepartmentsFilter, ClbDepartmentsFilter_ItemCheck), 3, 0);
            filterGrid.Controls.Add(CreateFilterBox("نواحی 🔺", clbDistrictsFilter, ClbDistrictsFilter_ItemCheck), 4, 0);
            filterGrid.Controls.Add(CreateFilterBox("پستها ⚡", clbPositionsFilter, ClbPositionsFilter_ItemCheck), 5, 0);

            // ردیف دوم
            filterGrid.Controls.Add(CreateFilterBox("جنسیت 👥", clbGenderFilter, ClbGenderFilter_ItemCheck), 0, 1);
            filterGrid.Controls.Add(CreateFilterBox("تحصیلات 📚", clbEducationFilter, ClbEducationFilter_ItemCheck), 1, 1);
            filterGrid.Controls.Add(CreateFilterBox("سطح شغلی 📊", clbJobLevelFilter, ClbJobLevelFilter_ItemCheck), 2, 1);
            filterGrid.Controls.Add(CreateFilterBox("نوع قرارداد 📄", clbContractTypeFilter, ClbContractTypeFilter_ItemCheck), 3, 1);
            filterGrid.Controls.Add(CreateFilterBox("شرکت 🏢", clbCompanyFilter, ClbCompanyFilter_ItemCheck), 4, 1);
            filterGrid.Controls.Add(CreateFilterBox("شیفت کاری ⏰", clbWorkShiftFilter, ClbWorkShiftFilter_ItemCheck), 5, 1);

            panelFilter.Controls.Add(filterGrid);

            // پایین پنل فیلتر: تاریخ استخدام + دکمه پاک کردن + پیام وضعیت
            Panel filterBottomPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent
            };

            FlowLayoutPanel rowActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 42,
                RightToLeft = RightToLeft.Yes,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                BackColor = Color.Transparent,
                Padding = new Padding(0, 2, 0, 0)
            };

            btnClearFilters.Text = "🔄 پاک کردن فیلترها";
            btnClearFilters.Size = new Size(170, 32);
            btnClearFilters.BackColor = Color.FromArgb(220, 53, 69);
            btnClearFilters.ForeColor = Color.White;
            btnClearFilters.Font = new Font(FontSettings.ButtonFont.FontFamily, 9.5F, FontStyle.Bold);
            btnClearFilters.FlatStyle = FlatStyle.Flat;
            btnClearFilters.FlatAppearance.BorderSize = 0;
            btnClearFilters.Margin = new Padding(6, 4, 6, 4);
            btnClearFilters.Click += BtnClearFilters_Click;

            Label lblHireDate = new Label
            {
                Text = "📅 تاریخ استخدام",
                AutoSize = true,
                Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                Margin = new Padding(6, 9, 6, 4)
            };

            chkHireDateFilter = new CheckBox
            {
                Text = "فعال",
                AutoSize = true,
                Font = new Font(FontSettings.BodyFont.FontFamily, 9F),
                Margin = new Padding(6, 9, 6, 4)
            };
            chkHireDateFilter.CheckedChanged += ChkHireDateFilter_CheckedChanged;

            dtpHireDateFrom = new DateTimePicker
            {
                Size = new Size(135, 26),
                Font = new Font(FontSettings.TextBoxFont.FontFamily, 9F),
                Enabled = false,
                Value = DateTime.Now.AddYears(-10),
                Format = DateTimePickerFormat.Short,
                Margin = new Padding(6, 6, 6, 4)
            };

            Label lblTo = new Label
            {
                Text = "تا",
                AutoSize = true,
                Font = new Font(FontSettings.LabelFont.FontFamily, 9F),
                Margin = new Padding(6, 9, 6, 4)
            };

            dtpHireDateTo = new DateTimePicker
            {
                Size = new Size(135, 26),
                Font = new Font(FontSettings.TextBoxFont.FontFamily, 9F),
                Enabled = false,
                Value = DateTime.Now,
                Format = DateTimePickerFormat.Short,
                Margin = new Padding(6, 6, 6, 4)
            };

            rowActions.Controls.Add(btnClearFilters);
            rowActions.Controls.Add(dtpHireDateTo);
            rowActions.Controls.Add(lblTo);
            rowActions.Controls.Add(dtpHireDateFrom);
            rowActions.Controls.Add(chkHireDateFilter);
            rowActions.Controls.Add(lblHireDate);

            lblFilterInfo.Text = "✓ فیلتری فعال نیست";
            lblFilterInfo.Dock = DockStyle.Bottom;
            lblFilterInfo.Height = 26;
            lblFilterInfo.Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold);
            lblFilterInfo.ForeColor = Color.FromArgb(0, 102, 204);
            lblFilterInfo.TextAlign = ContentAlignment.MiddleLeft;

            filterBottomPanel.Controls.Add(rowActions);
            filterBottomPanel.Controls.Add(lblFilterInfo);
            panelFilter.Controls.Add(filterBottomPanel);

            // ========== SplitContainer: بالا نمودارها (2/3) - پایین جدول‌ها (1/3) ==========
            SplitContainer mainSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                FixedPanel = FixedPanel.None,
                SplitterWidth = 6
            };

            // ========== بالا: نمودارها (TabControl) ==========
            tabControl.Dock = DockStyle.Fill;
            tabControl.RightToLeft = RightToLeft.Yes;
            tabControl.ItemSize = new Size(120, 30);
            tabControl.Font = FontSettings.BodyFont;

            AddChartTab(tabControl, "📊 ادارات", chartDepartmentPie, dgvDepartmentDetails);
            AddChartTab(tabControl, "💼 پستها", chartPositionPie, dgvPositionDetails);
            AddChartTab(tabControl, "👥 جنسیت", chartGenderPie, null);
            AddChartTab(tabControl, "📈 سطح شغلی", chartJobLevelPie, null);
            AddChartTab(tabControl, "📋 قرارداد", chartContractTypePie, null);
            AddChartTab(tabControl, "🗺️ استان", chartProvincePie, null);
            AddChartTab(tabControl, "📚 تحصیلات", chartEducationPie, null);
            AddChartTab(tabControl, "🏢 شرکت", chartCompanyPie, null);
            AddChartTab(tabControl, "⏰ شیفت", chartWorkShiftPie, null);

            mainSplit.Panel1.Controls.Add(tabControl);

            // ========== پایین: جدول آماری/خلاصه ==========
            Panel statsPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(6)
            };

            Panel radioPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = Color.FromArgb(230, 240, 250)
            };

            rbShowSummary = new RadioButton
            {
                Text = "📊 خلاصه آماری",
                Location = new Point(10, 9),
                Size = new Size(150, 25),
                Checked = true,
                Font = FontSettings.ButtonFont
            };
            rbShowSummary.CheckedChanged += RbShowSummary_CheckedChanged;
            radioPanel.Controls.Add(rbShowSummary);

            rbShowFullStats = new RadioButton
            {
                Text = "📋 جدول کامل آمار",
                Location = new Point(170, 9),
                Size = new Size(170, 25),
                Font = FontSettings.ButtonFont
            };
            rbShowFullStats.CheckedChanged += RbShowFullStats_CheckedChanged;
            radioPanel.Controls.Add(rbShowFullStats);

            dgvPersonnelStats.Dock = DockStyle.Fill;
            dgvPersonnelStats.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvPersonnelStats.ReadOnly = true;
            dgvPersonnelStats.RightToLeft = RightToLeft.Yes;
            dgvPersonnelStats.BackgroundColor = Color.White;
            dgvPersonnelStats.EnableHeadersVisualStyles = false;
            dgvPersonnelStats.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
            dgvPersonnelStats.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPersonnelStats.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
            dgvPersonnelStats.ColumnHeadersHeight = 35;
            dgvPersonnelStats.DefaultCellStyle.BackColor = Color.White;
            dgvPersonnelStats.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgvPersonnelStats.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            statsPanel.Controls.Add(dgvPersonnelStats);
            statsPanel.Controls.Add(radioPanel);

            mainSplit.Panel2.Controls.Add(statsPanel);

            Controls.Add(mainSplit);
            Controls.Add(panelFilter);

            // جلوگیری از خطای SplitterDistance (وقتی فرم هنوز اندازه نگرفته یا خیلی کوچک می‌شود)
            Shown += (s, e) =>
            {
                BeginInvoke((MethodInvoker)delegate
                {
                    ApplyMainSplitSizing(mainSplit);
                });
            };
            Resize += (s, e) =>
            {
                ApplyMainSplitSizing(mainSplit);
            };
        }

        private void ApplyMainSplitSizing(SplitContainer mainSplit)
        {
            if (mainSplit == null || mainSplit.IsDisposed)
                return;

            int total = mainSplit.Orientation == Orientation.Horizontal ? mainSplit.Height : mainSplit.Width;
            if (total <= 0)
                return;

            const int desiredPanel1Min = 250;
            const int desiredPanel2Min = 220;

            // فقط وقتی مین‌سایزها را اعمال کن که واقعاً جا باشد؛
            // وگرنه خود WinForms هنگام ApplyPanel2MinSize ممکن است SplitterDistance نامعتبر تولید کند.
            if (total > desiredPanel1Min + desiredPanel2Min + mainSplit.SplitterWidth)
            {
                mainSplit.Panel1MinSize = desiredPanel1Min;
                mainSplit.Panel2MinSize = desiredPanel2Min;
                SetSplitDistanceSafe(mainSplit, 0.66);
            }
            else
            {
                // حالت پنجره خیلی کوچک
                mainSplit.Panel1MinSize = 50;
                mainSplit.Panel2MinSize = 50;
                SetSplitDistanceSafe(mainSplit, 0.5);
            }
        }

        private void SetSplitDistanceSafe(SplitContainer sc, double ratio)
        {
            if (sc == null || sc.IsDisposed)
                return;

            int total = sc.Orientation == Orientation.Horizontal ? sc.Height : sc.Width;
            if (total <= 0)
                return;

            int min1 = sc.Panel1MinSize;
            int max = total - sc.Panel2MinSize - sc.SplitterWidth;
            if (max < min1)
                return;

            int desired = (int)(total * ratio);
            if (desired < min1) desired = min1;
            if (desired > max) desired = max;

            try
            {
                sc.SplitterDistance = desired;
            }
            catch
            {
                // ignore
            }
        }

        private Panel CreateFilterBox(string title, CheckedListBox clb, ItemCheckEventHandler eventHandler)
        {
            Panel box = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(5),
                Padding = new Padding(4),
                BackColor = Color.FromArgb(245, 252, 255),
                BorderStyle = BorderStyle.FixedSingle
            };

            Label lbl = new Label
            {
                Text = title,
                Dock = DockStyle.Top,
                AutoSize = false,
                Height = 22,
                Font = new Font(FontSettings.SubtitleFont.FontFamily, 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                TextAlign = ContentAlignment.MiddleLeft
            };

            SetupFilterCheckedListBox(clb, eventHandler);
            clb.Dock = DockStyle.Fill;

            box.Controls.Add(clb);
            box.Controls.Add(lbl);
            return box;
        }

        private void SetupFilterCheckedListBox(CheckedListBox clb, ItemCheckEventHandler eventHandler)
        {
            clb.RightToLeft = RightToLeft.Yes;
            clb.ItemCheck -= eventHandler;
            clb.ItemCheck += eventHandler;
            clb.BackColor = Color.White;
            clb.Font = new Font(FontSettings.BodyFont.FontFamily, 9F);
            clb.IntegralHeight = false;
            clb.HorizontalScrollbar = true;
            clb.HorizontalExtent = 2000; // برای اینکه متن‌های طولانی قطع نشوند (با اسکرول افقی)
        }

        private void CreateFilterColumn(Panel parent, string title, CheckedListBox clb, int x, int y, int width, int height, ItemCheckEventHandler eventHandler)
        {
            // (این متد از نسخه قبلی باقی مانده و فعلاً استفاده نمی‌شود)
            Label lbl = new Label
            {
                Text = title,
                Location = new Point(x, y),
                Size = new Size(width, 18),
                Font = new Font(FontSettings.SubtitleFont.FontFamily, 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204)
            };
            parent.Controls.Add(lbl);

            clb.Location = new Point(x, y + 22);
            clb.Size = new Size(width, height - 22);
            clb.RightToLeft = RightToLeft.Yes;
            clb.ItemCheck += eventHandler;
            clb.BackColor = Color.White;
            clb.Font = new Font(FontSettings.BodyFont.FontFamily, 8F);
            clb.IntegralHeight = false;
            parent.Controls.Add(clb);
        }

        private void RbShowSummary_CheckedChanged(object sender, EventArgs e)
        {
            if (rbShowSummary.Checked)
            {
                LoadSummaryTable();
            }
        }

        private void RbShowFullStats_CheckedChanged(object sender, EventArgs e)
        {
            if (rbShowFullStats.Checked)
            {
                LoadStatisticalTable();
            }
        }

        private void AddChartTab(TabControl tabControl, string title, Chart chart, DataGridView detailsGrid)
        {
            TabPage tab = new TabPage(title);

            if (detailsGrid != null)
            {
                SplitContainer split = new SplitContainer
                {
                    Dock = DockStyle.Fill,
                    Orientation = Orientation.Horizontal,
                    SplitterDistance = 400
                };

                chart.Dock = DockStyle.Fill;
                chart.BackColor = Color.White;
                chart.MinimumSize = new Size(100, 100);
                chart.ChartAreas.Add(new ChartArea("ChartArea1")
                {
                    BackColor = Color.White,
                    Area3DStyle = { Enable3D = true, Inclination = 15, Rotation = 45 }
                });
                chart.MouseClick += Chart_MouseClick;
                split.Panel1.Controls.Add(chart);

                detailsGrid.Dock = DockStyle.Fill;
                detailsGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                detailsGrid.ReadOnly = true;
                detailsGrid.RightToLeft = RightToLeft.Yes;
                detailsGrid.BackgroundColor = Color.White;
                detailsGrid.EnableHeadersVisualStyles = false;
                detailsGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
                detailsGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                detailsGrid.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
                split.Panel2.Controls.Add(detailsGrid);

                tab.Controls.Add(split);
            }
            else
            {
                chart.Dock = DockStyle.Fill;
                chart.BackColor = Color.White;
                chart.MinimumSize = new Size(100, 100);
                chart.ChartAreas.Add(new ChartArea("ChartArea1")
                {
                    BackColor = Color.White,
                    Area3DStyle = { Enable3D = true, Inclination = 15, Rotation = 45 }
                });
                chart.MouseClick += Chart_MouseClick;
                tab.Controls.Add(chart);
            }

            tabControl.TabPages.Add(tab);
        }

        private void LoadData()
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
                    MessageBox.Show("❌ خطا در بارگذاری دادهها.", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                LoadFilterOptions();
                RefreshAllCharts();
                MessageBox.Show($"✅ دادهها با موفقیت بارگذاری شدند.\n👥 تعداد پرسنل: {analyticsModel.TotalPersonnel}", "موفق", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadFilterOptions()
        {
            clbProvincesFilter.Items.Clear();
            foreach (var p in analyticsModel.GetAllProvinces())
                clbProvincesFilter.Items.Add(p, false);

            clbGenderFilter.Items.Clear();
            foreach (var g in analyticsModel.GetAllGenders())
                clbGenderFilter.Items.Add(g, false);

            clbEducationFilter.Items.Clear();
            foreach (var e in analyticsModel.GetAllEducations())
                clbEducationFilter.Items.Add(e, false);

            clbJobLevelFilter.Items.Clear();
            foreach (var j in analyticsModel.GetAllJobLevels())
                clbJobLevelFilter.Items.Add(j, false);

            clbContractTypeFilter.Items.Clear();
            foreach (var c in analyticsModel.GetAllContractTypes())
                clbContractTypeFilter.Items.Add(c, false);

            clbCompanyFilter.Items.Clear();
            foreach (var co in analyticsModel.GetAllCompanies())
                clbCompanyFilter.Items.Add(co, false);

            clbWorkShiftFilter.Items.Clear();
            foreach (var ws in analyticsModel.GetAllWorkShifts())
                clbWorkShiftFilter.Items.Add(ws, false);
        }

        private void ClbProvincesFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdateCitiesAndAffairs();
                RefreshAllCharts();
            });
        }

        private void ClbCitiesFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdateDepartmentsAndDistricts();
                RefreshAllCharts();
            });
        }

        private void ClbAffairsFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdateDepartmentsAndDistricts();
                RefreshAllCharts();
            });
        }

        private void ClbDepartmentsFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdateDistrictsAndPositions();
                RefreshAllCharts();
            });
        }

        private void ClbDistrictsFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                UpdatePositions();
                RefreshAllCharts();
            });
        }

        private void ClbPositionsFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbGenderFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbEducationFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbJobLevelFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbContractTypeFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbCompanyFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ClbWorkShiftFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void ChkHireDateFilter_CheckedChanged(object sender, EventArgs e)
        {
            dtpHireDateFrom.Enabled = chkHireDateFilter.Checked;
            dtpHireDateTo.Enabled = chkHireDateFilter.Checked;
            BeginInvoke((MethodInvoker)delegate
            {
                UpdateFilters();
                RefreshAllCharts();
            });
        }

        private void UpdateFilters()
        {
            List<string> selectedProvinces = clbProvincesFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedCities = clbCitiesFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedAffairs = clbAffairsFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedDepts = clbDepartmentsFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedDistricts = clbDistrictsFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedPositions = clbPositionsFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedGenders = clbGenderFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedEducations = clbEducationFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedJobLevels = clbJobLevelFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedContractTypes = clbContractTypeFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedCompanies = clbCompanyFilter.CheckedItems.Cast<string>().ToList();
            List<string> selectedWorkShifts = clbWorkShiftFilter.CheckedItems.Cast<string>().ToList();

            DateTime? hireFromDate = chkHireDateFilter.Checked ? dtpHireDateFrom.Value : (DateTime?)null;
            DateTime? hireToDate = chkHireDateFilter.Checked ? dtpHireDateTo.Value : (DateTime?)null;

            analyticsModel.SetFilters(selectedProvinces, selectedCities, selectedAffairs, selectedDepts,
                selectedDistricts, selectedPositions, selectedGenders, selectedEducations, selectedJobLevels,
                selectedContractTypes, selectedCompanies, selectedWorkShifts, hireFromDate, hireToDate);

            int filterCount = selectedProvinces.Count + selectedCities.Count + selectedAffairs.Count +
                selectedDepts.Count + selectedDistricts.Count + selectedPositions.Count +
                selectedGenders.Count + selectedEducations.Count + selectedJobLevels.Count +
                selectedContractTypes.Count + selectedCompanies.Count + selectedWorkShifts.Count +
                (chkHireDateFilter.Checked ? 1 : 0);

            lblFilterInfo.Text = filterCount > 0 ? $"🔴 {filterCount} فیلتر فعال" : "✓ فیلتری فعال نیست";
        }

        private void UpdateCitiesAndAffairs()
        {
            clbCitiesFilter.Items.Clear();
            clbAffairsFilter.Items.Clear();
            var selectedProvinces = clbProvincesFilter.CheckedItems.Cast<string>().ToList();

            if (selectedProvinces.Count > 0)
            {
                foreach (var city in analyticsModel.GetCitiesByProvinces(selectedProvinces).Distinct().OrderBy(x => x))
                    clbCitiesFilter.Items.Add(city, false);

                foreach (var affair in analyticsModel.GetAffairsByProvinces(selectedProvinces).Distinct().OrderBy(x => x))
                    clbAffairsFilter.Items.Add(affair, false);
            }
        }

        private void UpdateDepartmentsAndDistricts()
        {
            clbDepartmentsFilter.Items.Clear();
            clbDistrictsFilter.Items.Clear();
            var selectedProvinces = clbProvincesFilter.CheckedItems.Cast<string>().ToList();
            var selectedCities = clbCitiesFilter.CheckedItems.Cast<string>().ToList();
            var selectedAffairs = clbAffairsFilter.CheckedItems.Cast<string>().ToList();

            if (selectedProvinces.Count > 0 || selectedCities.Count > 0 || selectedAffairs.Count > 0)
            {
                foreach (var dept in analyticsModel.GetDepartmentsByFilters(selectedProvinces, selectedCities, selectedAffairs).Distinct().OrderBy(x => x))
                    clbDepartmentsFilter.Items.Add(dept, false);
            }
        }

        private void UpdateDistrictsAndPositions()
        {
            clbDistrictsFilter.Items.Clear();
            var selectedDepts = clbDepartmentsFilter.CheckedItems.Cast<string>().ToList();

            if (selectedDepts.Count > 0)
            {
                foreach (var district in analyticsModel.GetDistrictsByDepartments(selectedDepts).Distinct().OrderBy(x => x))
                    clbDistrictsFilter.Items.Add(district, false);
            }
        }

        private void UpdatePositions()
        {
            clbPositionsFilter.Items.Clear();
            var selectedDistricts = clbDistrictsFilter.CheckedItems.Cast<string>().ToList();

            if (selectedDistricts.Count > 0)
            {
                foreach (var pos in analyticsModel.GetPositionsByDistricts(selectedDistricts).Distinct().OrderBy(x => x))
                    clbPositionsFilter.Items.Add(pos, false);
            }
        }

        private void BtnClearFilters_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < clbProvincesFilter.Items.Count; i++) clbProvincesFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbCitiesFilter.Items.Count; i++) clbCitiesFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbAffairsFilter.Items.Count; i++) clbAffairsFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbDepartmentsFilter.Items.Count; i++) clbDepartmentsFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbDistrictsFilter.Items.Count; i++) clbDistrictsFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbPositionsFilter.Items.Count; i++) clbPositionsFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbGenderFilter.Items.Count; i++) clbGenderFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbEducationFilter.Items.Count; i++) clbEducationFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbJobLevelFilter.Items.Count; i++) clbJobLevelFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbContractTypeFilter.Items.Count; i++) clbContractTypeFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbCompanyFilter.Items.Count; i++) clbCompanyFilter.SetItemChecked(i, false);
            for (int i = 0; i < clbWorkShiftFilter.Items.Count; i++) clbWorkShiftFilter.SetItemChecked(i, false);
            chkHireDateFilter.Checked = false;

            analyticsModel.ClearFilters();
            lblFilterInfo.Text = "✓ فیلتری فعال نیست";
            LoadFilterOptions();
            RefreshAllCharts();
        }

        private void RefreshAllCharts()
        {
            if (rbShowSummary.Checked)
                LoadSummaryTable();
            else
                LoadStatisticalTable();

            LoadDepartmentPieChart();
            LoadPositionPieChart();
            LoadGenderPieChart();
            LoadJobLevelPieChart();
            LoadContractTypePieChart();
            LoadProvincePieChart();
            LoadEducationPieChart();
            LoadCompanyPieChart();
            LoadWorkShiftPieChart();
        }

        private void LoadSummaryTable()
        {
            try
            {
                dgvPersonnelStats.DataSource = null;
                dgvPersonnelStats.Columns.Clear();
                dgvPersonnelStats.Columns.Add("Metric", "معیار");
                dgvPersonnelStats.Columns.Add("Value", "مقدار");

                dgvPersonnelStats.Rows.Add("👥 کل پرسنل", analyticsModel.GetFilteredTotal());
                dgvPersonnelStats.Rows.Add("🏛️ تعداد ادارهها", analyticsModel.GetFilteredDepartmentCount());
                dgvPersonnelStats.Rows.Add("💼 تعداد پستهای شغلی", analyticsModel.GetFilteredPositionCount());
                dgvPersonnelStats.Rows.Add("🗺️ تعداد استانها", analyticsModel.ProvinceCount);
                dgvPersonnelStats.Rows.Add("🏢 تعداد شرکتها", analyticsModel.CompanyCount);
                dgvPersonnelStats.Rows.Add("📈 تعداد سطحهای شغلی", analyticsModel.JobLevelCount);
                dgvPersonnelStats.Rows.Add("📋 تعداد انواع قرارداد", analyticsModel.ContractTypeCount);
                dgvPersonnelStats.Rows.Add("📚 تعداد مدارک تحصیلی", analyticsModel.EducationCount);
                dgvPersonnelStats.Rows.Add("⏰ تعداد شیفت‌های کاری", analyticsModel.WorkShiftCount);
                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("👩 افراد خانم", analyticsModel.GetFilteredFemaleCount());
                dgvPersonnelStats.Rows.Add("👨 افراد آقا", analyticsModel.GetFilteredMaleCount());
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadDepartmentPieChart()
        {
            try
            {
                chartDepartmentPie.Series.Clear();
                var stats = analyticsModel.GetFilteredDepartmentStatistics();
                int total = stats.Sum(x => x.Count);

                Series series = new Series("درصد")
                {
                    ChartType = SeriesChartType.Pie,
                    IsValueShownAsLabel = true,
                    CustomProperties = "PieLabelStyle=Outside"
                };

                foreach (var item in stats.Take(15))
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                }

                chartDepartmentPie.Series.Add(series);
                chartDepartmentPie.Titles.Clear();
                chartDepartmentPie.Titles.Add(new Title("📊 توزیع پرسنل در ادارهها") { Font = FontSettings.HeaderFont });

                dgvDepartmentDetails.DataSource = null;
                dgvDepartmentDetails.Columns.Clear();
                dgvDepartmentDetails.Columns.Add("Name", "اداره");
                dgvDepartmentDetails.Columns.Add("Count", "تعداد");
                dgvDepartmentDetails.Columns.Add("Percent", "درصد");
                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    dgvDepartmentDetails.Rows.Add(item.Name, item.Count, $"{pct:F1}%");
                }
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadPositionPieChart()
        {
            try
            {
                chartPositionPie.Series.Clear();
                var stats = analyticsModel.GetFilteredPositionStatistics();
                int total = stats.Sum(x => x.Count);

                Series series = new Series("درصد")
                {
                    ChartType = SeriesChartType.Pie,
                    IsValueShownAsLabel = true,
                    CustomProperties = "PieLabelStyle=Outside"
                };

                foreach (var item in stats.Take(15))
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                }

                chartPositionPie.Series.Add(series);
                chartPositionPie.Titles.Clear();
                chartPositionPie.Titles.Add(new Title("💼 توزیع پستهای شغلی") { Font = FontSettings.HeaderFont });

                dgvPositionDetails.DataSource = null;
                dgvPositionDetails.Columns.Clear();
                dgvPositionDetails.Columns.Add("Name", "پست");
                dgvPositionDetails.Columns.Add("Count", "تعداد");
                dgvPositionDetails.Columns.Add("Percent", "درصد");
                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    dgvPositionDetails.Rows.Add(item.Name, item.Count, $"{pct:F1}%");
                }
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadGenderPieChart()
        {
            try
            {
                chartGenderPie.Series.Clear();
                var stats = analyticsModel.GetFilteredGenderStatistics();
                int total = stats.Sum(x => x.Count);

                Series series = new Series("درصد")
                {
                    ChartType = SeriesChartType.Pie,
                    IsValueShownAsLabel = true,
                    CustomProperties = "PieLabelStyle=Outside"
                };

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                }

                chartGenderPie.Series.Add(series);
                chartGenderPie.Titles.Clear();
                chartGenderPie.Titles.Add(new Title("👥 توزیع جنسیت") { Font = FontSettings.HeaderFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadJobLevelPieChart()
        {
            try
            {
                chartJobLevelPie.Series.Clear();
                var stats = analyticsModel.GetFilteredJobLevelStatistics();
                int total = stats.Sum(x => x.Count);

                Series series = new Series("درصد")
                {
                    ChartType = SeriesChartType.Pie,
                    IsValueShownAsLabel = true,
                    CustomProperties = "PieLabelStyle=Outside"
                };

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                }

                chartJobLevelPie.Series.Add(series);
                chartJobLevelPie.Titles.Clear();
                chartJobLevelPie.Titles.Add(new Title("📈 توزیع سطح شغلی") { Font = FontSettings.HeaderFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadContractTypePieChart()
        {
            try
            {
                chartContractTypePie.Series.Clear();
                var stats = analyticsModel.GetFilteredContractTypeStatistics();
                int total = stats.Sum(x => x.Count);

                Series series = new Series("درصد")
                {
                    ChartType = SeriesChartType.Pie,
                    IsValueShownAsLabel = true,
                    CustomProperties = "PieLabelStyle=Outside"
                };

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                }

                chartContractTypePie.Series.Add(series);
                chartContractTypePie.Titles.Clear();
                chartContractTypePie.Titles.Add(new Title("📋 توزیع نوع قرارداد") { Font = FontSettings.HeaderFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadProvincePieChart()
        {
            try
            {
                chartProvincePie.Series.Clear();
                var stats = analyticsModel.GetFilteredProvinceStatistics();
                int total = stats.Sum(x => x.Count);

                Series series = new Series("درصد")
                {
                    ChartType = SeriesChartType.Pie,
                    IsValueShownAsLabel = true,
                    CustomProperties = "PieLabelStyle=Outside"
                };

                foreach (var item in stats.Take(20))
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                }

                chartProvincePie.Series.Add(series);
                chartProvincePie.Titles.Clear();
                chartProvincePie.Titles.Add(new Title("🗺️ توزیع بر اساس استان") { Font = FontSettings.HeaderFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadEducationPieChart()
        {
            try
            {
                chartEducationPie.Series.Clear();
                var stats = analyticsModel.GetFilteredEducationStatistics();
                int total = stats.Sum(x => x.Count);

                Series series = new Series("درصد")
                {
                    ChartType = SeriesChartType.Pie,
                    IsValueShownAsLabel = true,
                    CustomProperties = "PieLabelStyle=Outside"
                };

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                }

                chartEducationPie.Series.Add(series);
                chartEducationPie.Titles.Clear();
                chartEducationPie.Titles.Add(new Title("📚 توزیع مدارک تحصیلی") { Font = FontSettings.HeaderFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadCompanyPieChart()
        {
            try
            {
                chartCompanyPie.Series.Clear();
                var stats = analyticsModel.GetFilteredCompanyStatistics();
                int total = stats.Sum(x => x.Count);

                Series series = new Series("درصد")
                {
                    ChartType = SeriesChartType.Pie,
                    IsValueShownAsLabel = true,
                    CustomProperties = "PieLabelStyle=Outside"
                };

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                }

                chartCompanyPie.Series.Add(series);
                chartCompanyPie.Titles.Clear();
                chartCompanyPie.Titles.Add(new Title("🏢 توزیع شرکتها") { Font = FontSettings.HeaderFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void LoadWorkShiftPieChart()
        {
            try
            {
                chartWorkShiftPie.Series.Clear();
                var stats = analyticsModel.GetFilteredWorkShiftStatistics();
                int total = stats.Sum(x => x.Count);

                Series series = new Series("درصد")
                {
                    ChartType = SeriesChartType.Pie,
                    IsValueShownAsLabel = true,
                    CustomProperties = "PieLabelStyle=Outside"
                };

                foreach (var item in stats)
                {
                    double pct = total > 0 ? (item.Count * 100.0) / total : 0;
                    int idx = series.Points.AddXY(item.Name, item.Count);
                    series.Points[idx].Label = $"{item.Name}\n{item.Count} نفر ({pct:F1}%)";
                    series.Points[idx].ToolTip = $"{item.Name}: {item.Count} نفر ({pct:F1}%)";
                }

                chartWorkShiftPie.Series.Add(series);
                chartWorkShiftPie.Titles.Clear();
                chartWorkShiftPie.Titles.Add(new Title("⏰ توزیع شیفت‌های کاری") { Font = FontSettings.HeaderFont });
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void Chart_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                Chart chart = sender as Chart;
                if (chart == null) return;

                HitTestResult result = chart.HitTest(e.X, e.Y);
                if (result.ChartElementType == ChartElementType.DataPoint)
                {
                    int pointIndex = result.PointIndex;
                    DataPoint point = result.Series.Points[pointIndex];
                    string itemName = point.AxisLabel;

                    var personnel = analyticsModel.GetPersonnelByFilter(itemName, chart);
                    if (personnel.Count > 0)
                        ShowPersonnelDetails(itemName, personnel);
                    else
                        MessageBox.Show("❌ داده‌ای برای نمایش وجود ندارد.", "پیام", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }

        private void ShowPersonnelDetails(string category, List<PersonnelDetail> personnel)
        {
            Form detailsForm = new Form
            {
                Text = $"👥 جزئیات پرسنل - {category}",
                Size = new Size(1400, 800),
                StartPosition = FormStartPosition.CenterScreen,
                RightToLeft = RightToLeft.Yes,
                BackColor = Color.FromArgb(240, 248, 255),
                WindowState = FormWindowState.Maximized,
                Font = FontSettings.BodyFont
            };

            // =============== DataGridView ===============
            DataGridView dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                ReadOnly = false,
                RightToLeft = RightToLeft.Yes,
                BackgroundColor = Color.White,
                EnableHeadersVisualStyles = false,
                AllowUserToAddRows = false,
                ColumnHeadersHeight = 40,
                RowTemplate = { Height = 35 }
            };

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 102, 204);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = FontSettings.SubtitleFont;
            dgv.DefaultCellStyle.Font = FontSettings.BodyFont;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255);

            dgv.Columns.Add("PersonnelID", "ID");
            dgv.Columns["PersonnelID"].Visible = false;
            dgv.Columns.Add("FirstName", "نام");
            dgv.Columns.Add("LastName", "نام‌خانوادگی");
            dgv.Columns.Add("PersonnelNumber", "شماره پرسنلی");
            dgv.Columns.Add("NationalID", "شناسه ملی");
            dgv.Columns.Add("PostName", "پست");
            dgv.Columns.Add("DeptName", "اداره");
            dgv.Columns.Add("Province", "استان");
            dgv.Columns.Add("ContractType", "نوع قرارداد");
            dgv.Columns.Add("HireDate", "تاریخ استخدام");
            dgv.Columns.Add("MobileNumber", "تلفن");

            // سیل‌های اکشن
            DataGridViewButtonColumn editColumn = new DataGridViewButtonColumn
            {
                Name = "Edit",
                HeaderText = "✏️ ویرایش",
                Text = "ویرایش",
                UseColumnTextForButtonValue = true,
                Width = 120,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(40, 167, 69),
                    ForeColor = Color.White,
                    Font = FontSettings.ButtonFont,
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    Padding = new Padding(5)
                }
            };
            dgv.Columns.Add(editColumn);

            DataGridViewButtonColumn deleteColumn = new DataGridViewButtonColumn
            {
                Name = "Delete",
                HeaderText = "🗑️ حذف",
                Text = "حذف",
                UseColumnTextForButtonValue = true,
                Width = 120,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.FromArgb(220, 53, 69),
                    ForeColor = Color.White,
                    Font = FontSettings.ButtonFont,
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    Padding = new Padding(5)
                }
            };
            dgv.Columns.Add(deleteColumn);

            foreach (var p in personnel)
            {
                dgv.Rows.Add(p.PersonnelID, p.FirstName, p.LastName, p.PersonnelNumber, p.NationalID, p.PostName,
                    p.DeptName, p.Province, p.ContractType, p.HireDate?.ToString("yyyy/MM/dd"), p.MobileNumber, "ویرایش", "حذف");
            }

            // Event Handler برای کلیک دکمه ها
            dgv.CellClick += (sender, e) =>
            {
                if (e.ColumnIndex == dgv.Columns["Edit"].Index && e.RowIndex >= 0)
                {
                    int personnelID = Convert.ToInt32(dgv.Rows[e.RowIndex].Cells["PersonnelID"].Value);
                    OpenEditForm(personnelID, detailsForm);
                }
                else if (e.ColumnIndex == dgv.Columns["Delete"].Index && e.RowIndex >= 0)
                {
                    int personnelID = Convert.ToInt32(dgv.Rows[e.RowIndex].Cells["PersonnelID"].Value);
                    DeletePersonnel(personnelID, detailsForm, dgv, e.RowIndex);
                }
            };

            // =============== پنل دکمه‌های پایین ===============
            Panel bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 70,
                BackColor = Color.FromArgb(230, 240, 250)
            };

            Button btnExportExcel = new Button
            {
                Text = "📊 خروجی اکسل",
                Location = new Point(20, 15),
                Size = new Size(200, 40),
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                Font = FontSettings.ButtonFont,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnExportExcel.FlatAppearance.BorderSize = 0;
            btnExportExcel.Click += (s, ev) =>
            {
                using (ExportColumnsForm exportForm = new ExportColumnsForm())
                {
                    if (exportForm.ShowDialog() == DialogResult.OK)
                    {
                        var selectedColumns = exportForm.SelectedColumns;
                        if (selectedColumns != null && selectedColumns.Count > 0)
                        {
                            ExcelExportHelper.ExportToExcel(personnel, selectedColumns, $"Personnel_{category}");
                        }
                    }
                }
            };
            bottomPanel.Controls.Add(btnExportExcel);

            detailsForm.Controls.Add(dgv);
            detailsForm.Controls.Add(bottomPanel);
            detailsForm.ShowDialog();
        }

        private void OpenEditForm(int personnelID, Form parentForm)
        {
            try
            {
                FormPersonnelEdit editForm = new FormPersonnelEdit();
                editForm.txtPersonnelID.Text = personnelID.ToString();
                editForm.BtnLoad_Click(null, EventArgs.Empty);
                editForm.ShowDialog(parentForm);

                RefreshAllCharts();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در باز کردن فرم ویرایش: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeletePersonnel(int personnelID, Form parentForm, DataGridView dgv, int rowIndex)
        {
            try
            {
                DialogResult result = MessageBox.Show(
                    $"❓ آیا مطمئن هستید که می‌خواهید این پرسنل را حذف کنید؟",
                    "تایید حذف",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    string query = $"DELETE FROM Personnel WHERE PersonnelID = {personnelID}";
                    dbHelper.ExecuteNonQuery(query);

                    MessageBox.Show("✅ پرسنل با موفقیت حذف شد.", "موفق", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    dgv.Rows.RemoveAt(rowIndex);

                    LoadData();
                    RefreshAllCharts();

                    if (dgv.Rows.Count == 0)
                    {
                        parentForm.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در حذف پرسنل: {ex.Message}", "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadStatisticalTable()
        {
            try
            {
                dgvPersonnelStats.DataSource = null;
                dgvPersonnelStats.Columns.Clear();
                dgvPersonnelStats.Columns.Add("Metric", "معیار");
                dgvPersonnelStats.Columns.Add("Value", "مقدار");

                dgvPersonnelStats.Rows.Add("═══════════════════", "");
                dgvPersonnelStats.Rows.Add("👥 کل پرسنل", analyticsModel.GetFilteredTotal());
                dgvPersonnelStats.Rows.Add("🏛️ تعداد ادارهها", analyticsModel.GetFilteredDepartmentCount());
                dgvPersonnelStats.Rows.Add("💼 تعداد پستهای شغلی", analyticsModel.GetFilteredPositionCount());
                dgvPersonnelStats.Rows.Add("🗺️ تعداد استانها", analyticsModel.ProvinceCount);
                dgvPersonnelStats.Rows.Add("🏢 تعداد شرکتها", analyticsModel.CompanyCount);
                dgvPersonnelStats.Rows.Add("📈 تعداد سطحهای شغلی", analyticsModel.JobLevelCount);
                dgvPersonnelStats.Rows.Add("📋 تعداد انواع قرارداد", analyticsModel.ContractTypeCount);
                dgvPersonnelStats.Rows.Add("📚 تعداد مدارک تحصیلی", analyticsModel.EducationCount);
                dgvPersonnelStats.Rows.Add("⏰ تعداد شیفت‌های کاری", analyticsModel.WorkShiftCount);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("═════ توزیع جنسیت ═════", "");
                foreach (var g in analyticsModel.GetFilteredGenderStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {g.Name}", g.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════ توزیع سطح شغلی ════", "");
                foreach (var j in analyticsModel.GetFilteredJobLevelStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {j.Name}", j.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════ توزیع نوع قرارداد ════", "");
                foreach (var c in analyticsModel.GetFilteredContractTypeStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {c.Name}", c.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("═════════ تمام ادارات ═════════", "");
                foreach (var d in analyticsModel.GetFilteredDepartmentStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {d.Name}", d.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════════ تمام پستهای شغلی ════════", "");
                foreach (var p in analyticsModel.GetFilteredPositionStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {p.Name}", p.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════════════ تمام استان‌ها ════════════", "");
                foreach (var pr in analyticsModel.GetFilteredProvinceStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {pr.Name}", pr.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("════════════ تمام شرکتها ════════════", "");
                foreach (var co in analyticsModel.GetFilteredCompanyStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {co.Name}", co.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("═════════ تمام مدارک تحصیلی ═════════", "");
                foreach (var e in analyticsModel.GetFilteredEducationStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {e.Name}", e.Count);

                dgvPersonnelStats.Rows.Add("", "");
                dgvPersonnelStats.Rows.Add("═════════ تمام شیفت‌های کاری ═════════", "");
                foreach (var ws in analyticsModel.GetFilteredWorkShiftStatistics())
                    dgvPersonnelStats.Rows.Add($"  • {ws.Name}", ws.Count);
            }
            catch (Exception ex) { MessageBox.Show($"❌ خطا: {ex.Message}"); }
        }
    }
}
