using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;

namespace PersonnelManagementApp
{
    public enum PostChartKind
    {
        Category,
        OperationYear,
        Capacity
    }

    public enum PostCapacityScale
    {
        None,
        HighVoltage,
        MediumVoltage,
        Distributed
    }

    /// <summary>
    /// تعریف یک نمودار تحلیلی پست‌ها.
    /// </summary>
    public sealed class PostChartSpec
    {
        public PostChartSpec(string key, string title, string column, PostChartKind kind,
            PostCapacityScale capacityScale = PostCapacityScale.None)
        {
            Key = key;
            Title = title;
            Column = column;
            Kind = kind;
            CapacityScale = capacityScale;
        }

        public string Key { get; }
        public string Title { get; }
        public string Column { get; }
        public PostChartKind Kind { get; }
        public PostCapacityScale CapacityScale { get; }
    }

    /// <summary>
    /// یک دسته در نمودار (نام + تعداد).
    /// </summary>
    public sealed class PostStat
    {
        public PostStat(string label, int count, int sortOrder = 0, bool isOther = false)
        {
            Label = label;
            Count = count;
            SortOrder = sortOrder;
            IsOther = isOther;
        }

        public string Label { get; }
        public int Count { get; }
        public int SortOrder { get; }
        public bool IsOther { get; }
    }

    /// <summary>
    /// فهرست نمودارهای پست‌ها.
    /// </summary>
    public static class PostChartCatalog
    {
        public static IReadOnlyList<PostChartSpec> All { get; } = new List<PostChartSpec>
        {
            new("province", "توزیع پست‌ها بر اساس استان", "ProvinceName", PostChartKind.Category),
            new("city", "توزیع پست‌ها بر اساس شهر", "CityName", PostChartKind.Category),
            new("affair", "توزیع پست‌ها بر اساس امور", "AffairName", PostChartKind.Category),
            new("department", "توزیع پست‌ها بر اساس اداره", "DeptName", PostChartKind.Category),
            new("district", "توزیع پست‌ها بر اساس ناحیه", "DistrictName", PostChartKind.Category),
            new("postname", "توزیع پست‌ها بر اساس نام پست", "PostName", PostChartKind.Category),
            new("voltage", "توزیع بر اساس سطح ولتاژ", "VoltageName", PostChartKind.Category),
            new("standard", "توزیع بر اساس استاندارد پست", "StandardName", PostChartKind.Category),
            new("type", "توزیع بر اساس نوع پست", "TypeName", PostChartKind.Category),
            new("connection", "توزیع بر اساس اتصال توزیع", "ConnName", PostChartKind.Category),
            new("insulation", "توزیع بر اساس نوع عایق", "InsName", PostChartKind.Category),
            new("posttype2", "توزیع بر اساس نوع پست ۲", "PT2Name", PostChartKind.Category),
            new("fixedmobile", "توزیع بر اساس ثابت / سیار", "FMName", PostChartKind.Category),
            new("circuit", "توزیع بر اساس وضعیت مدار", "CircuitName", PostChartKind.Category),
            new("diesel", "توزیع بر اساس دیزل ژنراتور", "DieselName", PostChartKind.Category),
            new("feed", "توزیع بر اساس نوع توزیع", "FeedName", PostChartKind.Category),
            new("water", "توزیع بر اساس وضعیت آب", "WaterName", PostChartKind.Category),
            new("guest", "توزیع بر اساس میهمان‌سرا", "GuestName", PostChartKind.Category),
            new("operationyear", "توزیع بر اساس دهه بهره‌برداری", "OperationYear", PostChartKind.OperationYear),
            new("capacityhv", "توزیع ظرفیت HV", "CapacityHV", PostChartKind.Capacity, PostCapacityScale.HighVoltage),
            new("capacitymv", "توزیع ظرفیت MV", "CapacityMV", PostChartKind.Capacity, PostCapacityScale.MediumVoltage),
            new("distributedcapacity", "توزیع ظرفیت توزیع‌شده", "DistributedCapacity", PostChartKind.Capacity, PostCapacityScale.Distributed)
        };

        public static PostChartSpec? Find(string? key)
        {
            if (string.IsNullOrWhiteSpace(key))
                return null;

            return All.FirstOrDefault(spec => string.Equals(spec.Key, key, StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>
    /// بارگذاری داده‌های پست‌ها به‌صورت کامل و مقاوم در برابر کلیدهای خارجی خالی.
    /// </summary>
    public sealed class PostsDataService
    {
        private readonly DbHelper _db;

        private static readonly LookupMap[] Lookups =
        {
            new("ProvinceID", "Provinces", "ProvinceID", "ProvinceName", "ProvinceName"),
            new("CityID", "Cities", "CityID", "CityName", "CityName"),
            new("AffairID", "TransferAffairs", "AffairID", "AffairName", "AffairName"),
            new("DeptID", "OperationDepartments", "DeptID", "DeptName", "DeptName"),
            new("DistrictID", "Districts", "DistrictID", "DistrictName", "DistrictName"),
            new("PostNameID", "PostsNames", "PostNameID", "PostName", "PostName"),
            new("VoltageID", "VoltageLevels", "VoltageID", "VoltageName", "VoltageName"),
            new("StandardID", "PostStandards", "StandardID", "StandardName", "StandardName"),
            new("TypeID", "PostTypes", "TypeID", "TypeName", "TypeName"),
            new("ConnID", "DistributedConnections", "ConnID", "ConnName", "ConnName"),
            new("InsID", "InsulationTypes", "InsID", "InsName", "InsName"),
            new("PT2ID", "PostTypeTwos", "PT2ID", "PT2Name", "PT2Name"),
            new("FMID", "FixedMobiles", "FMID", "FMName", "FMName"),
            new("CircuitID", "CircuitStatuses", "CircuitID", "CircuitName", "CircuitName"),
            new("DieselID", "DieselGenerators", "DieselID", "DieselName", "DieselName"),
            new("FeedID", "DistributionFeeds", "FeedID", "FeedName", "FeedName"),
            new("WaterID", "WaterStatuses", "WaterID", "WaterName", "WaterName"),
            new("GuestID", "GuestHouses", "GuestID", "GuestName", "GuestName")
        };

        public PostsDataService(DbHelper db)
        {
            _db = db;
        }

        public PostsDataService() : this(new DbHelper())
        {
        }

        /// <summary>
        /// خواندن همه پست‌ها و افزودن نام‌های نمایشی جدول‌های مرجع.
        /// </summary>
        public DataTable LoadAllPosts(bool showError = true)
        {
            DataTable? posts = _db.ExecuteQuery("SELECT * FROM Posts", null, showError);
            if (posts == null)
                return CreateEmptyPostsTable();

            foreach (LookupMap map in Lookups)
            {
                if (!posts.Columns.Contains(map.DisplayColumn))
                {
                    posts.Columns.Add(new DataColumn(map.DisplayColumn, typeof(string))
                    {
                        DefaultValue = "نامشخص"
                    });
                }
            }

            foreach (LookupMap map in Lookups)
            {
                Dictionary<int, string> values = LoadLookupValues(map);
                foreach (DataRow row in posts.Rows)
                {
                    int id = ToInt(row[map.ForeignKey]);
                    string name = id > 0 && values.TryGetValue(id, out string? value) && !string.IsNullOrWhiteSpace(value)
                        ? value
                        : "نامشخص";

                    row[map.DisplayColumn] = name;
                }
            }

            return posts;
        }

        private Dictionary<int, string> LoadLookupValues(LookupMap map)
        {
            var result = new Dictionary<int, string>();

            string query = $"SELECT [{map.IdColumn}], [{map.NameColumn}] FROM [{map.Table}]";
            DataTable? table = _db.ExecuteQuery(query, null, showError: false);
            if (table == null)
                return result;

            foreach (DataRow row in table.Rows)
            {
                int id = ToInt(row[map.IdColumn]);
                string name = row[map.NameColumn]?.ToString()?.Trim() ?? string.Empty;
                if (id > 0 && !string.IsNullOrWhiteSpace(name))
                    result[id] = name;
            }

            return result;
        }

        private static int ToInt(object? value)
        {
            if (value == null || value == DBNull.Value)
                return 0;

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private static DataTable CreateEmptyPostsTable()
        {
            var table = new DataTable("Posts");

            foreach (string column in new[]
            {
                "PostID", "ProvinceID", "CityID", "AffairID", "DeptID", "DistrictID",
                "PostNameID", "VoltageID", "StandardID", "TypeID", "ConnID", "InsID",
                "PT2ID", "FMID", "CircuitID", "DieselID", "FeedID", "WaterID", "GuestID"
            })
            {
                table.Columns.Add(column, typeof(int));
            }

            foreach (string column in new[]
            {
                "OperationYear", "ParkingBays"
            })
            {
                table.Columns.Add(column, typeof(int));
            }

            foreach (string column in new[]
            {
                "DistributedCapacity", "CapacityHV", "CapacityMV", "TotalArea",
                "BuildingsArea", "SiteArea", "DistanceToCapital"
            })
            {
                table.Columns.Add(column, typeof(double));
            }

            foreach (string column in new[]
            {
                "Longitude", "Latitude", "PostalCode", "PostalAddress", "Description"
            })
            {
                table.Columns.Add(column, typeof(string));
            }

            foreach (LookupMap map in Lookups)
                table.Columns.Add(map.DisplayColumn, typeof(string));

            return table;
        }

        private sealed class LookupMap
        {
            public LookupMap(string foreignKey, string table, string idColumn, string nameColumn, string displayColumn)
            {
                ForeignKey = foreignKey;
                Table = table;
                IdColumn = idColumn;
                NameColumn = nameColumn;
                DisplayColumn = displayColumn;
            }

            public string ForeignKey { get; }
            public string Table { get; }
            public string IdColumn { get; }
            public string NameColumn { get; }
            public string DisplayColumn { get; }
        }
    }
}
