using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;

namespace PersonnelManagementApp
{
    /// <summary>
    /// داده‌های قابل ویرایش یک پست.
    /// </summary>
    public sealed class PostEditData
    {
        public int ProvinceId { get; set; }
        public int CityId { get; set; }
        public int AffairId { get; set; }
        public int DeptId { get; set; }
        public int DistrictId { get; set; }
        public int PostNameId { get; set; }
        public int VoltageId { get; set; }
        public int StandardId { get; set; }
        public int TypeId { get; set; }
        public int ConnId { get; set; }
        public int InsId { get; set; }
        public int Pt2Id { get; set; }
        public int FmId { get; set; }
        public int CircuitId { get; set; }
        public int DieselId { get; set; }
        public int FeedId { get; set; }
        public int WaterId { get; set; }
        public int GuestId { get; set; }

        public int OperationYear { get; set; }
        public double CapacityHv { get; set; }
        public double CapacityMv { get; set; }
        public double DistributedCapacity { get; set; }
        public double TotalArea { get; set; }
        public double BuildingsArea { get; set; }
        public double SiteArea { get; set; }
        public int ParkingBays { get; set; }
        public double DistanceToCapital { get; set; }

        public string PostalCode { get; set; } = string.Empty;
        public string PostalAddress { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string Longitude { get; set; } = string.Empty;
        public string Latitude { get; set; } = string.Empty;
    }

    /// <summary>
    /// عملیات افزودن / ویرایش / حذف پست‌ها.
    /// </summary>
    public sealed class PostsCrudService
    {
        private readonly DbHelper _db;

        public PostsCrudService(DbHelper db)
        {
            _db = db;
        }

        public PostsCrudService() : this(new DbHelper())
        {
        }

        /// <summary>
        /// خواندن مقادیر یک جدول مرجع به صورت (شناسه، نام).
        /// </summary>
        public List<KeyValuePair<int, string>> GetLookupValues(string table, string idColumn, string nameColumn)
        {
            var result = new List<KeyValuePair<int, string>>();
            DataTable? data = _db.ExecuteQuery(
                $"SELECT [{idColumn}], [{nameColumn}] FROM [{table}] ORDER BY [{nameColumn}]",
                null,
                showError: false);

            if (data == null)
                return result;

            foreach (DataRow row in data.Rows)
            {
                int id = ToInt(row[idColumn]);
                string name = row[nameColumn]?.ToString()?.Trim() ?? string.Empty;
                if (id > 0 && !string.IsNullOrWhiteSpace(name))
                    result.Add(new KeyValuePair<int, string>(id, name));
            }

            return result;
        }

        /// <summary>
        /// افزودن یک پست جدید (PostID خودکار است). شناسه جدید را برمی‌گرداند.
        /// </summary>
        public int InsertPost(PostEditData data)
        {
            const string query =
                "INSERT INTO Posts (" +
                "ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, " +
                "VoltageID, StandardID, TypeID, ConnID, InsID, PT2ID, FMID, CircuitID, " +
                "DieselID, FeedID, WaterID, GuestID, " +
                "OperationYear, CapacityHV, CapacityMV, DistributedCapacity, " +
                "TotalArea, BuildingsArea, SiteArea, ParkingBays, DistanceToCapital, " +
                "PostalCode, PostalAddress, Description, Longitude, Latitude) " +
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

            int inserted = _db.ExecuteNonQuery(query, BuildParameters(data), showError: false);
            if (inserted <= 0)
                return 0;

            object? id = _db.ExecuteScalar("SELECT @@IDENTITY", null, showError: false);
            return ToInt(id);
        }

        /// <summary>
        /// ویرایش یک پست موجود.
        /// </summary>
        public bool UpdatePost(int postId, PostEditData data)
        {
            const string query =
                "UPDATE Posts SET " +
                "ProvinceID = ?, CityID = ?, AffairID = ?, DeptID = ?, DistrictID = ?, PostNameID = ?, " +
                "VoltageID = ?, StandardID = ?, TypeID = ?, ConnID = ?, InsID = ?, PT2ID = ?, FMID = ?, CircuitID = ?, " +
                "DieselID = ?, FeedID = ?, WaterID = ?, GuestID = ?, " +
                "OperationYear = ?, CapacityHV = ?, CapacityMV = ?, DistributedCapacity = ?, " +
                "TotalArea = ?, BuildingsArea = ?, SiteArea = ?, ParkingBays = ?, DistanceToCapital = ?, " +
                "PostalCode = ?, PostalAddress = ?, Description = ?, Longitude = ?, Latitude = ? " +
                "WHERE PostID = ?";

            OleDbParameter[] parameters = BuildParameters(data);
            Array.Resize(ref parameters, parameters.Length + 1);
            parameters[parameters.Length - 1] = new OleDbParameter("?", postId);

            return _db.ExecuteNonQuery(query, parameters, showError: false) > 0;
        }

        /// <summary>
        /// حذف یک پست.
        /// </summary>
        public bool DeletePost(int postId)
        {
            const string query = "DELETE FROM Posts WHERE PostID = ?";
            return _db.ExecuteNonQuery(query,
                new[] { new OleDbParameter("?", postId) },
                showError: false) > 0;
        }

        private static OleDbParameter[] BuildParameters(PostEditData d)
        {
            return new OleDbParameter[]
            {
                new("?", d.ProvinceId),
                new("?", d.CityId),
                new("?", d.AffairId),
                new("?", d.DeptId),
                new("?", d.DistrictId),
                new("?", d.PostNameId),
                new("?", d.VoltageId),
                new("?", d.StandardId),
                new("?", d.TypeId),
                new("?", d.ConnId),
                new("?", d.InsId),
                new("?", d.Pt2Id),
                new("?", d.FmId),
                new("?", d.CircuitId),
                new("?", d.DieselId),
                new("?", d.FeedId),
                new("?", d.WaterId),
                new("?", d.GuestId),
                new("?", d.OperationYear),
                new("?", d.CapacityHv),
                new("?", d.CapacityMv),
                new("?", d.DistributedCapacity),
                new("?", d.TotalArea),
                new("?", d.BuildingsArea),
                new("?", d.SiteArea),
                new("?", d.ParkingBays),
                new("?", d.DistanceToCapital),
                new("?", string.IsNullOrWhiteSpace(d.PostalCode) ? "" : d.PostalCode.Trim()),
                new("?", string.IsNullOrWhiteSpace(d.PostalAddress) ? "" : d.PostalAddress.Trim()),
                new("?", string.IsNullOrWhiteSpace(d.Description) ? "" : d.Description.Trim()),
                new("?", string.IsNullOrWhiteSpace(d.Longitude) ? "" : d.Longitude.Trim()),
                new("?", string.IsNullOrWhiteSpace(d.Latitude) ? "" : d.Latitude.Trim())
            };
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
    }
}
