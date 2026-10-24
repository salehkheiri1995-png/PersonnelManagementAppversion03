using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace PersonnelManagementApp
{
    /// <summary>
    /// ساخت دسته‌بندی‌ها و رندر یکپارچه نمودارهای پست‌ها.
    /// </summary>
    public static class PostChartRenderer
    {
        private static readonly Color[] Palette =
        {
            Color.FromArgb(0, 122, 204),
            Color.FromArgb(16, 185, 129),
            Color.FromArgb(245, 158, 11),
            Color.FromArgb(239, 68, 68),
            Color.FromArgb(139, 92, 246),
            Color.FromArgb(14, 165, 233),
            Color.FromArgb(236, 72, 153),
            Color.FromArgb(132, 204, 22),
            Color.FromArgb(249, 115, 22),
            Color.FromArgb(99, 102, 241),
            Color.FromArgb(20, 184, 166),
            Color.FromArgb(168, 85, 247)
        };

        private static readonly Color OtherColor = Color.FromArgb(148, 163, 184);
        private static readonly Color TextColor = Color.FromArgb(30, 41, 59);
        private static readonly Color MutedText = Color.FromArgb(100, 116, 139);
        private static readonly Color GridColor = Color.FromArgb(226, 232, 240);

        public static bool IsPie(SeriesChartType type)
            => type == SeriesChartType.Pie || type == SeriesChartType.Doughnut;

        public static bool IsHorizontalBar(SeriesChartType type)
            => type == SeriesChartType.Bar || type == SeriesChartType.StackedBar;

        /// <summary>
        /// برچسب یک ردیف در نمودار جاری.
        /// </summary>
        public static string GetLabel(PostChartSpec spec, DataRow row)
        {
            if (spec == null || row == null)
                return "نامشخص";

            switch (spec.Kind)
            {
                case PostChartKind.OperationYear:
                    return GetOperationYear(row).Label;

                case PostChartKind.Capacity:
                    return GetCapacity(row, spec.CapacityScale).Label;

                default:
                    if (!row.Table.Columns.Contains(spec.Column))
                        return "نامشخص";

                    string value = row[spec.Column]?.ToString()?.Trim() ?? string.Empty;
                    return string.IsNullOrWhiteSpace(value) ? "نامشخص" : value;
            }
        }

        /// <summary>
        /// ساخت آمار دسته‌بندی‌شده. برای نمودارهای دایره‌ای می‌توان تعداد دسته‌ها را محدود کرد.
        /// </summary>
        public static List<PostStat> BuildStats(DataTable data, PostChartSpec spec, int maxCategories = 0)
        {
            if (data == null || data.Rows.Count == 0 || spec == null)
                return new List<PostStat>();

            List<PostStat> stats;

            switch (spec.Kind)
            {
                case PostChartKind.OperationYear:
                    stats = data.AsEnumerable()
                        .Select(GetOperationYear)
                        .GroupBy(item => item.Label)
                        .Select(group => new PostStat(group.Key, group.Count(), group.First().SortOrder))
                        .OrderBy(stat => stat.SortOrder)
                        .ToList();
                    break;

                case PostChartKind.Capacity:
                    stats = data.AsEnumerable()
                        .Select(row => GetCapacity(row, spec.CapacityScale))
                        .GroupBy(item => item.Label)
                        .Select(group => new PostStat(group.Key, group.Count(), group.First().SortOrder))
                        .OrderBy(stat => stat.SortOrder)
                        .ToList();
                    break;

                default:
                    stats = data.AsEnumerable()
                        .GroupBy(row => GetLabel(spec, row))
                        .Select(group => new PostStat(group.Key, group.Count()))
                        .OrderBy(stat => stat.Label == "نامشخص" ? 1 : 0)
                        .ThenByDescending(stat => stat.Count)
                        .ThenBy(stat => stat.Label, StringComparer.CurrentCulture)
                        .ToList();
                    break;
            }

            if (maxCategories > 0 && stats.Count > maxCategories)
            {
                int keep = Math.Max(1, maxCategories - 1);
                List<PostStat> top = stats.Take(keep).ToList();
                int otherCount = stats.Skip(keep).Sum(stat => stat.Count);
                top.Add(new PostStat("سایر", otherCount, int.MaxValue, isOther: true));
                stats = top;
            }

            return stats;
        }

        public static void Render(Chart chart, DataTable data, PostChartSpec spec, SeriesChartType type, bool showTitle = true, bool enable3D = false)
        {
            if (chart == null || spec == null)
                return;

            chart.SuspendLayout();
            try
            {
                chart.Series.Clear();
                chart.Titles.Clear();
                chart.Legends.Clear();
                chart.ChartAreas.Clear();
                chart.BackColor = Color.White;
                chart.BorderlineColor = GridColor;
                chart.BorderlineWidth = 1;
                chart.AntiAliasing = AntiAliasingStyles.All;
                chart.TextAntiAliasingQuality = TextAntiAliasingQuality.High;
                chart.Palette = ChartColorPalette.None;

                Font labelFont = FontSettings.ChartLabelFont ?? new Font("Tahoma", 9F);
                Font titleFont = FontSettings.HeaderFont ?? new Font("Tahoma", 12F, FontStyle.Bold);

                var area = new ChartArea("MainArea")
                {
                    BackColor = Color.White,
                    BorderColor = Color.Transparent
                };

                bool pie = IsPie(type);
                bool horizontal = IsHorizontalBar(type);

                area.AxisX.MajorGrid.Enabled = false;
                area.AxisY.MajorGrid.Enabled = false;
                area.AxisX.LineColor = GridColor;
                area.AxisY.LineColor = GridColor;
                area.AxisX.LabelStyle.Font = labelFont;
                area.AxisY.LabelStyle.Font = labelFont;
                area.AxisX.LabelStyle.ForeColor = MutedText;
                area.AxisY.LabelStyle.ForeColor = MutedText;
                area.AxisX.IsMarginVisible = true;
                area.AxisY.IsStartedFromZero = true;

                if (pie)
                {
                    area.Area3DStyle.Enable3D = enable3D;
                    area.Area3DStyle.Inclination = enable3D ? 15 : 0;
                    area.Area3DStyle.Rotation = enable3D ? 45 : 0;
                    area.AxisX.Enabled = AxisEnabled.False;
                    area.AxisY.Enabled = AxisEnabled.False;
                }
                else
                {
                    area.AxisX.Enabled = AxisEnabled.True;
                    area.AxisY.Enabled = AxisEnabled.True;
                    area.AxisY.MajorGrid.Enabled = true;
                    area.AxisY.MajorGrid.LineColor = GridColor;
                    area.AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

                    if (horizontal)
                    {
                        area.AxisY.Interval = 1;
                        area.AxisY.LabelStyle.Angle = 0;
                        area.AxisX.MajorGrid.Enabled = true;
                        area.AxisY.MajorGrid.Enabled = false;
                    }
                    else
                    {
                        area.AxisX.Interval = 1;
                        area.AxisX.LabelStyle.Angle = -45;
                        area.AxisX.LabelStyle.IsStaggered = false;
                    }
                }

                chart.ChartAreas.Add(area);

                var legend = new Legend("Legend1")
                {
                    Docking = Docking.Bottom,
                    Alignment = StringAlignment.Center,
                    Font = labelFont,
                    ForeColor = TextColor,
                    BackColor = Color.Transparent,
                    IsTextAutoFit = true,
                    Enabled = pie
                };
                chart.Legends.Add(legend);

                int maxCategories = pie ? 12 : 0;
                List<PostStat> stats = BuildStats(data, spec, maxCategories);
                int total = stats.Sum(stat => stat.Count);

                var series = new Series("تعداد")
                {
                    ChartType = type,
                    XValueType = ChartValueType.String,
                    YValueType = ChartValueType.Int32,
                    IsXValueIndexed = true,
                    Font = labelFont,
                    IsValueShownAsLabel = !pie && stats.Count <= 20,
                    Legend = pie ? "Legend1" : string.Empty,
                    SmartLabelStyle =
                    {
                        Enabled = true,
                        CalloutStyle = LabelCalloutStyle.Underlined,
                        AllowOutsidePlotArea = LabelOutsidePlotAreaStyle.Yes
                    }
                };

                if (pie)
                {
                    series["PieLabelStyle"] = "Outside";
                    series["PieLineColor"] = "LightGray";
                    series["PieStartAngle"] = "270";
                }

                for (int i = 0; i < stats.Count; i++)
                {
                    PostStat stat = stats[i];
                    double pct = total > 0 ? stat.Count * 100.0 / total : 0;
                    int index = series.Points.AddXY(stat.Label, stat.Count);
                    DataPoint point = series.Points[index];

                    point.Color = stat.IsOther ? OtherColor : Palette[i % Palette.Length];
                    point.AxisLabel = stat.Label;
                    point.ToolTip = $"{stat.Label}: {stat.Count} پست ({pct:F1}%)";
                    point.Font = labelFont;
                    point.LabelForeColor = TextColor;

                    if (pie)
                    {
                        point.Label = stats.Count <= 8
                            ? $"{stat.Label}\n{stat.Count} ({pct:F1}%)"
                            : $"{pct:F1}%";
                    }
                    else
                    {
                        point.Label = stats.Count <= 20 ? stat.Count.ToString() : string.Empty;
                    }

                    if (!stat.IsOther)
                        point.Color = Palette[i % Palette.Length];
                }

                chart.Series.Add(series);

                if (showTitle)
                {
                    string titleText = stats.Count == 0
                        ? $"{spec.Title}  |  داده‌ای موجود نیست"
                        : $"{spec.Title}  |  {total} پست";

                    var title = new Title(titleText, Docking.Top, titleFont, TextColor)
                    {
                        Alignment = ContentAlignment.MiddleCenter,
                        Font = titleFont,
                        ForeColor = TextColor
                    };
                    chart.Titles.Add(title);
                }
            }
            finally
            {
                chart.ResumeLayout();
            }
        }

        private static (string Label, int SortOrder) GetOperationYear(DataRow row)
        {
            if (!row.Table.Columns.Contains("OperationYear"))
                return ("نامشخص", int.MaxValue);

            if (!int.TryParse(row["OperationYear"]?.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int year) || year <= 0)
                return ("نامشخص", int.MaxValue);

            int decade = (year / 10) * 10;
            return ($"{decade}–{decade + 9}", decade);
        }

        private static (string Label, int SortOrder) GetCapacity(DataRow row, PostCapacityScale scale)
        {
            string column = scale switch
            {
                PostCapacityScale.HighVoltage => "CapacityHV",
                PostCapacityScale.MediumVoltage => "CapacityMV",
                _ => "DistributedCapacity"
            };

            if (!row.Table.Columns.Contains(column))
                return ("بدون مقدار", int.MaxValue);

            if (!double.TryParse(row[column]?.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double value) || value <= 0)
                return ("بدون مقدار", int.MaxValue);

            return scale switch
            {
                PostCapacityScale.HighVoltage => value switch
                {
                    <= 50 => ("0 تا 50", 1),
                    <= 100 => ("50 تا 100", 2),
                    <= 250 => ("100 تا 250", 3),
                    <= 500 => ("250 تا 500", 4),
                    <= 1000 => ("500 تا 1000", 5),
                    _ => ("بیش از 1000", 6)
                },
                PostCapacityScale.MediumVoltage => value switch
                {
                    <= 10 => ("0 تا 10", 1),
                    <= 30 => ("10 تا 30", 2),
                    <= 60 => ("30 تا 60", 3),
                    <= 100 => ("60 تا 100", 4),
                    _ => ("بیش از 100", 5)
                },
                _ => value switch
                {
                    <= 1 => ("0 تا 1", 1),
                    <= 5 => ("1 تا 5", 2),
                    <= 10 => ("5 تا 10", 3),
                    <= 25 => ("10 تا 25", 4),
                    <= 50 => ("25 تا 50", 5),
                    _ => ("بیش از 50", 6)
                }
            };
        }

        public static SeriesChartType GetChartType(Chart? chart, SeriesChartType fallback = SeriesChartType.Pie)
        {
            if (chart?.Series.Count > 0)
                return chart.Series[0].ChartType;
            return fallback;
        }
    }

    /// <summary>
    /// ذخیره تصویر و چاپ نمودار پست‌ها.
    /// </summary>
    public static class PostChartImageExporter
    {
        private const TextFormatFlags RtlFlags =
            TextFormatFlags.RightToLeft |
            TextFormatFlags.WordBreak |
            TextFormatFlags.NoPrefix |
            TextFormatFlags.TextBoxControl;

        public static bool Save(Chart chart, DataTable data, PostChartSpec spec, IWin32Window? owner = null)
        {
            if (chart == null || spec == null)
                return false;

            using var dialog = new SaveFileDialog
            {
                Filter = "PNG Image|*.png|JPEG Image|*.jpg|BMP Image|*.bmp",
                Title = "ذخیره نمودار",
                FileName = $"نمودار_{spec.Key}_{DateTime.Now:yyyyMMdd_HHmmss}.png",
                AddExtension = true
            };

            if (dialog.ShowDialog(owner) != DialogResult.OK)
                return false;

            try
            {
                SaveToFile(chart, data, spec, dialog.FileName);
                MessageBox.Show(owner, "✅ نمودار با موفقیت ذخیره شد.", "ذخیره تصویر",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            catch (Exception ex)
            {
                ErrorLogger.Log("PostChartImageExporter.Save", ex);
                MessageBox.Show(owner, $"خطا در ذخیره تصویر:\n{ex.Message}", "خطا",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public static void SaveToFile(Chart chart, DataTable data, PostChartSpec spec, string filePath)
        {
            const int width = 1800;
            const int height = 1120;

            using var bitmap = new Bitmap(width, height);
            using Graphics graphics = Graphics.FromImage(bitmap);

            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            graphics.Clear(Color.White);

            var titleFont = new Font(FontSettings.FontFamilyName, 24, FontStyle.Bold);
            var subtitleFont = new Font(FontSettings.FontFamilyName, 11, FontStyle.Regular);
            var statsFont = new Font(FontSettings.FontFamilyName, 11, FontStyle.Regular);
            var statsTitleFont = new Font(FontSettings.FontFamilyName, 14, FontStyle.Bold);

            DrawHeader(graphics, width, spec, titleFont, subtitleFont);

            var chartRect = new Rectangle(45, 125, 1130, 920);
            var statsRect = new Rectangle(1215, 135, 535, 880);

            using (var separatorPen = new Pen(Color.FromArgb(226, 232, 240), 1))
            {
                graphics.DrawLine(separatorPen, 1195, 130, 1195, 1030);
            }

            chart.Printing.PrintPaint(graphics, chartRect);

            var stats = PostChartRenderer.BuildStats(data, spec, 22);
            int total = stats.Sum(stat => stat.Count);

            TextRenderer.DrawText(graphics, "آمار نمودار", statsTitleFont, new Rectangle(statsRect.X, statsRect.Y, statsRect.Width, 32), Color.FromArgb(30, 41, 59), RtlFlags);
            TextRenderer.DrawText(graphics, $"تعداد کل: {total} پست", statsFont,
                new Rectangle(statsRect.X, statsRect.Y + 40, statsRect.Width, 28), Color.FromArgb(0, 122, 204), RtlFlags);

            int y = statsRect.Y + 80;
            foreach (PostStat stat in stats)
            {
                if (y + 26 > statsRect.Bottom)
                    break;

                double pct = total > 0 ? stat.Count * 100.0 / total : 0;
                string line = $"• {stat.Label}: {stat.Count} پست ({pct:F1}%)";
                TextRenderer.DrawText(graphics, line, statsFont,
                    new Rectangle(statsRect.X, y, statsRect.Width, 25), Color.FromArgb(51, 65, 85), RtlFlags);
                y += 28;
            }

            TextRenderer.DrawText(graphics, $"تاریخ تولید: {DateTime.Now:yyyy/MM/dd - HH:mm}", subtitleFont,
                new Rectangle(45, height - 55, width - 90, 25), Color.FromArgb(100, 116, 139), RtlFlags);

            ImageFormat format = ImageFormat.Png;
            string extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension == ".jpg" || extension == ".jpeg")
                format = ImageFormat.Jpeg;
            else if (extension == ".bmp")
                format = ImageFormat.Bmp;

            bitmap.Save(filePath, format);
        }

        public static void PrintPreview(Chart chart, DataTable data, PostChartSpec spec, IWin32Window? owner = null)
        {
            if (chart == null || spec == null)
                return;

            using var document = new PrintDocument
            {
                DocumentName = spec.Title
            };
            document.DefaultPageSettings.Landscape = true;

            document.PrintPage += (_, e) =>
            {
                if (e.Graphics == null)
                    return;

                Rectangle bounds = e.MarginBounds;
                var titleFont = new Font(FontSettings.FontFamilyName, 18, FontStyle.Bold);
                var bodyFont = new Font(FontSettings.FontFamilyName, 9, FontStyle.Regular);

                TextRenderer.DrawText(e.Graphics, spec.Title, titleFont,
                    new Rectangle(bounds.X, bounds.Y, bounds.Width, 34), Color.FromArgb(30, 41, 59),
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPrefix);

                var chartRect = new Rectangle(
                    bounds.X,
                    bounds.Y + 45,
                    (int)(bounds.Width * 0.66),
                    bounds.Height - 90);

                var statsRect = new Rectangle(
                    bounds.X + (int)(bounds.Width * 0.69),
                    bounds.Y + 45,
                    (int)(bounds.Width * 0.31),
                    bounds.Height - 90);

                chart.Printing.PrintPaint(e.Graphics, chartRect);

                var stats = PostChartRenderer.BuildStats(data, spec, 25);
                int total = stats.Sum(stat => stat.Count);

                TextRenderer.DrawText(e.Graphics, $"تعداد کل: {total} پست", bodyFont,
                    new Rectangle(statsRect.X, statsRect.Y, statsRect.Width, 24),
                    Color.FromArgb(0, 122, 204), RtlFlags);

                int y = statsRect.Y + 32;
                foreach (PostStat stat in stats)
                {
                    if (y + 22 > statsRect.Bottom)
                        break;

                    double pct = total > 0 ? stat.Count * 100.0 / total : 0;
                    TextRenderer.DrawText(e.Graphics, $"• {stat.Label}: {stat.Count} ({pct:F1}%)", bodyFont,
                        new Rectangle(statsRect.X, y, statsRect.Width, 22), Color.FromArgb(51, 65, 85), RtlFlags);
                    y += 23;
                }

                e.HasMorePages = false;
            };

            using var preview = new PrintPreviewDialog
            {
                Document = document,
                Width = 1100,
                Height = 800,
                UseAntiAlias = true,
                Text = "پیش‌نمایش چاپ نمودار"
            };

            preview.ShowDialog(owner);
        }

        private static void DrawHeader(Graphics graphics, int width, PostChartSpec spec, Font titleFont, Font subtitleFont)
        {
            using var accentBrush = new SolidBrush(Color.FromArgb(0, 122, 204));
            graphics.FillRectangle(accentBrush, 45, 35, 8, 55);

            TextRenderer.DrawText(graphics, "تحلیل داده‌های پست‌ها", titleFont,
                new Rectangle(70, 28, width - 110, 42), Color.FromArgb(15, 23, 42),
                TextFormatFlags.RightToLeft | TextFormatFlags.NoPrefix);

            TextRenderer.DrawText(graphics, spec.Title, subtitleFont,
                new Rectangle(70, 72, width - 110, 28), Color.FromArgb(71, 85, 105),
                TextFormatFlags.RightToLeft | TextFormatFlags.NoPrefix);
        }
    }
}
