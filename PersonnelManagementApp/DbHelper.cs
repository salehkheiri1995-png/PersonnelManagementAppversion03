using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    public class DbHelper
    {
        private readonly string connectionString;

        public DbHelper()
        {
            connectionString = GetConnectionString();
        }

        private string GetConnectionString()
        {
            try
            {
                // ⭐ استفاده از AppSettings برای مسیر دیتابیس
                string dbPath = AppSettings.DatabasePath;

                if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
                {
                    dbPath = SelectDatabasePath();
                    if (string.IsNullOrEmpty(dbPath))
                    {
                        throw new FileNotFoundException("پایگاه داده انتخاب نشد.");
                    }
                }

                return $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath};Persist Security Info=False;";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در یافتن مسیر پایگاه داده: {ex.Message}",
                                "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private string SelectDatabasePath()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "انتخاب فایل پایگاه داده";
                openFileDialog.Filter = "Access Database (*.accdb)|*.accdb|Access Database (*.mdb)|*.mdb|All Files (*.*)|*.*";
                openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = openFileDialog.FileName;
                    // ⭐ ذخیره در AppSettings
                    AppSettings.DatabasePath = selectedPath;
                    MessageBox.Show($"پایگاه داده انتخاب شد:\n{selectedPath}\n\nاین مسیر برای استفاده‌های بعدی ذخیره شد.",
                                    "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return selectedPath;
                }
            }

            return null;
        }

        public string GetConnectionString_Public()
        {
            return connectionString;
        }

        public DataTable? ExecuteQuery(string query, OleDbParameter[]? parameters = null)
        {
            DataTable dt = new DataTable();
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
                        {
                            da.Fill(dt);
                        }
                    }
                }
                return dt.Rows.Count > 0 ? dt : null;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در اتصال به پایگاه داده: {ex.Message}\nInner Exception: {ex.InnerException?.Message}\nConnection String: {connectionString}\nQuery: {query}",
                                "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        public int ExecuteNonQuery(string query, OleDbParameter[]? parameters = null)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    using (OleDbCommand command = new OleDbCommand(query, connection))
                    {
                        if (parameters != null)
                        {
                            command.Parameters.AddRange(parameters);
                        }
                        connection.Open();
                        return command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در اجرای عملیات: {ex.Message}\nInner Exception: {ex.InnerException?.Message}\nConnection String: {connectionString}\nQuery: {query}",
                                "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }
        }

        public bool TestConnection()
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"تست اتصال ناموفق بود: {ex.Message}\nInner Exception: {ex.InnerException?.Message}\nConnection String: {connectionString}",
                                "خطای اتصال", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public DataTable? SearchByPostName(string searchTerm)
        {
            string query = "SELECT Posts.PostID, Posts.ProvinceID, Posts.CityID, Posts.AffairID, Posts.DeptID, Posts.DistrictID, Posts.PostNameID, Posts.VoltageID, Posts.StandardID, Posts.TypeID, Posts.ConnID, Posts.DistributedCapacity, Posts.InsID, Posts.PT2ID, Posts.OperationYear, Posts.FMID, Posts.CircuitID, Posts.BayTrans400, Posts.BayLine400, Posts.BayTrans230, Posts.BayLine230, Posts.BayTrans132, Posts.BayLine132, Posts.BayTrans63, Posts.BayLine63, Posts.IncomingIF, Posts.Feeder20DG, Posts.FeederOut20, Posts.FeederCap20, Posts.CapacityHV, Posts.CapacityMV, Posts.DieselID, Posts.FeedID, Posts.WaterID, Posts.TotalArea, Posts.BuildingsArea, Posts.SiteArea, Posts.ParkingBays, Posts.DistanceToCapital, Posts.GuestID, Posts.Longitude, Posts.Latitude, Posts.PostalCode, Posts.PostalAddress, Posts.Description, " +
                          "Provinces.ProvinceName, Cities.CityName, TransferAffairs.AffairName, OperationDepartments.DeptName, Districts.DistrictName, PostsNames.PostName, VoltageLevels.VoltageName, PostStandards.StandardName, PostTypes.TypeName, DistributedConnections.ConnName, InsulationTypes.InsName, PostTypeTwos.PT2Name, FixedMobiles.FMName, CircuitStatuses.CircuitName, DieselGenerators.DieselName, DistributionFeeds.FeedName, WaterStatuses.WaterName, GuestHouses.GuestName " +
                          "FROM Posts " +
                          "INNER JOIN Provinces ON Posts.ProvinceID = Provinces.ProvinceID " +
                          "INNER JOIN Cities ON Posts.CityID = Cities.CityID " +
                          "INNER JOIN TransferAffairs ON Posts.AffairID = TransferAffairs.AffairID " +
                          "INNER JOIN OperationDepartments ON Posts.DeptID = OperationDepartments.DeptID " +
                          "INNER JOIN Districts ON Posts.DistrictID = Districts.DistrictID " +
                          "INNER JOIN PostsNames ON Posts.PostNameID = PostsNames.PostNameID " +
                          "INNER JOIN VoltageLevels ON Posts.VoltageID = VoltageLevels.VoltageID " +
                          "INNER JOIN PostStandards ON Posts.StandardID = PostStandards.StandardID " +
                          "INNER JOIN PostTypes ON Posts.TypeID = PostTypes.TypeID " +
                          "INNER JOIN DistributedConnections ON Posts.ConnID = DistributedConnections.ConnID " +
                          "INNER JOIN InsulationTypes ON Posts.InsID = InsulationTypes.InsID " +
                          "INNER JOIN PostTypeTwos ON Posts.PT2ID = PostTypeTwos.PT2ID " +
                          "INNER JOIN FixedMobiles ON Posts.FMID = FixedMobiles.FMID " +
                          "INNER JOIN CircuitStatuses ON Posts.CircuitID = CircuitStatuses.CircuitID " +
                          "INNER JOIN DieselGenerators ON Posts.DieselID = DieselGenerators.DieselID " +
                          "INNER JOIN DistributionFeeds ON Posts.FeedID = DistributionFeeds.FeedID " +
                          "INNER JOIN WaterStatuses ON Posts.WaterID = WaterStatuses.WaterID " +
                          "INNER JOIN GuestHouses ON Posts.GuestID = GuestHouses.GuestID " +
                          "WHERE PostsNames.PostName LIKE ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", "%" + searchTerm + "%") };
            return ExecuteQuery(query, parameters);
        }

        public DataTable? SearchByPersonnel(string searchTerm)
        {
            string query = "SELECT Personnel.*, Provinces.ProvinceName, Cities.CityName, TransferAffairs.AffairName, " +
                          "OperationDepartments.DeptName, Districts.DistrictName, PostsNames.PostName, " +
                          "VoltageLevels.VoltageName, WorkShift.WorkShiftName, Gender.GenderName, " +
                          "ContractType.ContractTypeName, JobLevel.JobLevelName, Company.CompanyName, " +
                          "Degree.DegreeName, DegreeField.DegreeFieldName, ChartAffairs1.ChartName AS MainJobTitle, " +
                          "ChartAffairs2.ChartName AS CurrentActivity, StatusPresence.StatusName " +
                          "FROM Personnel " +
                          "INNER JOIN Provinces ON Personnel.ProvinceID = Provinces.ProvinceID " +
                          "INNER JOIN Cities ON Personnel.CityID = Cities.CityID " +
                          "INNER JOIN TransferAffairs ON Personnel.AffairID = TransferAffairs.AffairID " +
                          "INNER JOIN OperationDepartments ON Personnel.DeptID = OperationDepartments.DeptID " +
                          "INNER JOIN Districts ON Personnel.DistrictID = Districts.DistrictID " +
                          "INNER JOIN PostsNames ON Personnel.PostNameID = PostsNames.PostNameID " +
                          "INNER JOIN VoltageLevels ON Personnel.VoltageID = VoltageLevels.VoltageID " +
                          "INNER JOIN WorkShift ON Personnel.WorkShiftID = WorkShift.WorkShiftID " +
                          "INNER JOIN Gender ON Personnel.GenderID = Gender.GenderID " +
                          "INNER JOIN ContractType ON Personnel.ContractTypeID = ContractType.ContractTypeID " +
                          "INNER JOIN JobLevel ON Personnel.JobLevelID = JobLevel.JobLevelID " +
                          "INNER JOIN Company ON Personnel.CompanyID = Company.CompanyID " +
                          "INNER JOIN Degree ON Personnel.DegreeID = Degree.DegreeID " +
                          "INNER JOIN DegreeField ON Personnel.DegreeFieldID = DegreeField.DegreeFieldID " +
                          "INNER JOIN ChartAffairs AS ChartAffairs1 ON Personnel.MainJobTitle = ChartAffairs1.ChartID " +
                          "INNER JOIN ChartAffairs AS ChartAffairs2 ON Personnel.CurrentActivity = ChartAffairs2.ChartID " +
                          "INNER JOIN StatusPresence ON Personnel.StatusID = StatusPresence.StatusID " +
                          "WHERE Personnel.FirstName LIKE ? OR Personnel.LastName LIKE ? OR Personnel.PersonnelNumber LIKE ? OR Personnel.NationalID LIKE ?";
            OleDbParameter[] parameters = new OleDbParameter[]
            {
                new OleDbParameter("?", "%" + searchTerm + "%"),
                new OleDbParameter("?", "%" + searchTerm + "%"),
                new OleDbParameter("?", "%" + searchTerm + "%"),
                new OleDbParameter("?", "%" + searchTerm + "%")
            };
            return ExecuteQuery(query, parameters);
        }

        public DataTable? GetPostsByProvince()
        {
            string query = "SELECT Provinces.ProvinceName, COUNT(Posts.PostID) AS PostCount " +
                          "FROM Posts INNER JOIN Provinces ON Posts.ProvinceID = Provinces.ProvinceID " +
                          "GROUP BY Provinces.ProvinceName";
            return ExecuteQuery(query);
        }

        public void ExportToCsv(DataTable? dt, string filePath)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("داده‌ای برای صدور وجود ندارد.", "هشدار", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (StreamWriter sw = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        sw.Write($"\"{dt.Columns[i].ColumnName}\"");
                        if (i < dt.Columns.Count - 1)
                            sw.Write(",");
                    }
                    sw.WriteLine();

                    foreach (DataRow row in dt.Rows)
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            string value = row[i]?.ToString()?.Replace("\"", "\"\"") ?? "";
                            sw.Write($"\"{value}\"");
                            if (i < dt.Columns.Count - 1)
                                sw.Write(",");
                        }
                        sw.WriteLine();
                    }
                }
                MessageBox.Show($"با موفقیت به {filePath} صادر شد!", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در صدور: {ex.Message}\nInner Exception: {ex.InnerException?.Message}",
                                "خطا", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public DataTable? GetProvinces()
        {
            return ExecuteQuery("SELECT ProvinceID, ProvinceName FROM Provinces ORDER BY ProvinceName");
        }

        public DataTable? GetCitiesByProvince(int provinceID)
        {
            string query = provinceID == 0
                ? "SELECT CityID, CityName FROM Cities ORDER BY CityName"
                : "SELECT CityID, CityName FROM Cities WHERE ProvinceID = ? ORDER BY CityName";
            OleDbParameter[]? parameters = provinceID == 0 ? null : new OleDbParameter[] { new OleDbParameter("?", provinceID) };
            return ExecuteQuery(query, parameters);
        }

        public DataTable? GetAffairsByProvince(int provinceID)
        {
            string query = provinceID == 0
                ? "SELECT AffairID, AffairName FROM TransferAffairs ORDER BY AffairName"
                : "SELECT AffairID, AffairName FROM TransferAffairs WHERE ProvinceID = ? ORDER BY AffairName";
            OleDbParameter[]? parameters = provinceID == 0 ? null : new OleDbParameter[] { new OleDbParameter("?", provinceID) };
            return ExecuteQuery(query, parameters);
        }

        public DataTable? GetDeptsByAffair(int affairID)
        {
            string query = affairID == 0
                ? "SELECT DeptID, DeptName FROM OperationDepartments ORDER BY DeptName"
                : "SELECT DeptID, DeptName FROM OperationDepartments WHERE AffairID = ? ORDER BY DeptName";
            OleDbParameter[]? parameters = affairID == 0 ? null : new OleDbParameter[] { new OleDbParameter("?", affairID) };
            return ExecuteQuery(query, parameters);
        }

        public DataTable? GetDistrictsByDept(int deptID)
        {
            string query = deptID == 0
                ? "SELECT DistrictID, DistrictName FROM Districts ORDER BY DistrictName"
                : "SELECT DistrictID, DistrictName FROM Districts WHERE DeptID = ? ORDER BY DistrictName";
            OleDbParameter[]? parameters = deptID == 0 ? null : new OleDbParameter[] { new OleDbParameter("?", deptID) };
            return ExecuteQuery(query, parameters);
        }

        public DataTable? GetPostNamesByDistrict(int districtID)
        {
            string query = districtID == 0
                ? "SELECT PostNameID, PostName FROM PostsNames ORDER BY PostName"
                : "SELECT PostNameID, PostName FROM PostsNames WHERE DistrictID = ? ORDER BY PostName";
            OleDbParameter[]? parameters = districtID == 0 ? null : new OleDbParameter[] { new OleDbParameter("?", districtID) };
            return ExecuteQuery(query, parameters);
        }

        public DataTable? GetVoltageLevels()
        {
            return ExecuteQuery("SELECT VoltageID, VoltageName FROM VoltageLevels ORDER BY VoltageName");
        }

        public DataTable? GetPostStandards()
        {
            return ExecuteQuery("SELECT StandardID, StandardName FROM PostStandards ORDER BY StandardName");
        }

        public DataTable? GetPostTypes()
        {
            return ExecuteQuery("SELECT TypeID, TypeName FROM PostTypes ORDER BY TypeName");
        }

        public DataTable? GetDistributedConnections()
        {
            return ExecuteQuery("SELECT ConnID, ConnName FROM DistributedConnections ORDER BY ConnName");
        }

        public DataTable? GetInsulationTypes()
        {
            return ExecuteQuery("SELECT InsID, InsName FROM InsulationTypes ORDER BY InsName");
        }

        public DataTable? GetPostTypeTwos()
        {
            return ExecuteQuery("SELECT PT2ID, PT2Name FROM PostTypeTwos ORDER BY PT2Name");
        }

        public DataTable? GetFixedMobiles()
        {
            return ExecuteQuery("SELECT FMID, FMName FROM FixedMobiles ORDER BY FMName");
        }

        public DataTable? GetCircuitStatuses()
        {
            return ExecuteQuery("SELECT CircuitID, CircuitName FROM CircuitStatuses ORDER BY CircuitName");
        }

        public DataTable? GetDieselGenerators()
        {
            return ExecuteQuery("SELECT DieselID, DieselName FROM DieselGenerators ORDER BY DieselName");
        }

        public DataTable? GetDistributionFeeds()
        {
            return ExecuteQuery("SELECT FeedID, FeedName FROM DistributionFeeds ORDER BY FeedName");
        }

        public DataTable? GetWaterStatuses()
        {
            return ExecuteQuery("SELECT WaterID, WaterName FROM WaterStatuses ORDER BY WaterName");
        }

        public DataTable? GetGuestHouses()
        {
            return ExecuteQuery("SELECT GuestID, GuestName FROM GuestHouses ORDER BY GuestName");
        }

        public DataTable? GetWorkShifts()
        {
            return ExecuteQuery("SELECT WorkShiftID, WorkShiftName FROM WorkShift ORDER BY WorkShiftName");
        }

        public DataTable? GetGenders()
        {
            return ExecuteQuery("SELECT GenderID, GenderName FROM Gender ORDER BY GenderName");
        }

        public DataTable? GetContractTypes()
        {
            return ExecuteQuery("SELECT ContractTypeID, ContractTypeName FROM ContractType ORDER BY ContractTypeName");
        }

        public DataTable? GetJobLevels()
        {
            return ExecuteQuery("SELECT JobLevelID, JobLevelName FROM JobLevel ORDER BY JobLevelName");
        }

        public DataTable? GetCompanies()
        {
            return ExecuteQuery("SELECT CompanyID, CompanyName FROM Company ORDER BY CompanyName");
        }

        public DataTable? GetDegrees()
        {
            return ExecuteQuery("SELECT DegreeID, DegreeName FROM Degree ORDER BY DegreeName");
        }

        public DataTable? GetDegreeFields()
        {
            return ExecuteQuery("SELECT DegreeFieldID, DegreeFieldName FROM DegreeField ORDER BY DegreeFieldName");
        }

        public DataTable? GetStatusPresence()
        {
            return ExecuteQuery("SELECT StatusID, StatusName FROM StatusPresence ORDER BY StatusName");
        }

        public DataTable? GetChartAffairs()
        {
            return ExecuteQuery("SELECT ChartID, AffairID, ChartName FROM ChartAffairs ORDER BY ChartName");
        }

        public DataTable? GetChartAffairs1()
        {
            return ExecuteQuery("SELECT ChartID, AffairID, ChartName FROM ChartAffairs ORDER BY ChartName");
        }

        public object? GetProvinceIDByName(string provinceName)
        {
            string query = "SELECT ProvinceID FROM Provinces WHERE ProvinceName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", provinceName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["ProvinceID"] : null;
        }

        public object? GetCityIDByName(string cityName)
        {
            string query = "SELECT CityID FROM Cities WHERE CityName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", cityName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["CityID"] : null;
        }

        public object? GetAffairIDByName(string affairName)
        {
            string query = "SELECT AffairID FROM TransferAffairs WHERE AffairName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", affairName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["AffairID"] : null;
        }

        public object? GetDeptIDByName(string deptName)
        {
            string query = "SELECT DeptID FROM OperationDepartments WHERE DeptName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", deptName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["DeptID"] : null;
        }

        public object? GetDistrictIDByName(string districtName)
        {
            string query = "SELECT DistrictID FROM Districts WHERE DistrictName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", districtName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["DistrictID"] : null;
        }

        public object? GetPostNameIDByName(string postName)
        {
            string query = "SELECT PostNameID FROM PostsNames WHERE PostName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", postName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["PostNameID"] : null;
        }

        public object? GetVoltageIDByName(string voltageName)
        {
            string query = "SELECT VoltageID FROM VoltageLevels WHERE VoltageName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", voltageName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["VoltageID"] : null;
        }

        public object? GetStandardIDByName(string standardName)
        {
            string query = "SELECT StandardID FROM PostStandards WHERE StandardName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", standardName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["StandardID"] : null;
        }

        public object? GetTypeIDByName(string typeName)
        {
            string query = "SELECT TypeID FROM PostTypes WHERE TypeName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", typeName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["TypeID"] : null;
        }

        public object? GetConnIDByName(string connName)
        {
            string query = "SELECT ConnID FROM DistributedConnections WHERE ConnName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", connName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["ConnID"] : null;
        }

        public object? GetInsIDByName(string insName)
        {
            string query = "SELECT InsID FROM InsulationTypes WHERE InsName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", insName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["InsID"] : null;
        }

        public object? GetPT2IDByName(string pt2Name)
        {
            string query = "SELECT PT2ID FROM PostTypeTwos WHERE PT2Name = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", pt2Name) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["PT2ID"] : null;
        }

        public object? GetFMIDByName(string fmName)
        {
            string query = "SELECT FMID FROM FixedMobiles WHERE FMName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", fmName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["FMID"] : null;
        }

        public object? GetCircuitIDByName(string circuitName)
        {
            string query = "SELECT CircuitID FROM CircuitStatuses WHERE CircuitName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", circuitName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["CircuitID"] : null;
        }

        public object? GetDieselIDByName(string dieselName)
        {
            string query = "SELECT DieselID FROM DieselGenerators WHERE DieselName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", dieselName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["DieselID"] : null;
        }

        public object? GetFeedIDByName(string feedName)
        {
            string query = "SELECT FeedID FROM DistributionFeeds WHERE FeedName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", feedName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["FeedID"] : null;
        }

        public object? GetWaterIDByName(string waterName)
        {
            string query = "SELECT WaterID FROM WaterStatuses WHERE WaterName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", waterName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["WaterID"] : null;
        }

        public object? GetGuestIDByName(string guestName)
        {
            string query = "SELECT GuestID FROM GuestHouses WHERE GuestName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", guestName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["GuestID"] : null;
        }

        public object? GetWorkShiftIDByName(string workShiftName)
        {
            string query = "SELECT WorkShiftID FROM WorkShift WHERE WorkShiftName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", workShiftName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["WorkShiftID"] : null;
        }

        public object? GetGenderIDByName(string genderName)
        {
            string query = "SELECT GenderID FROM Gender WHERE GenderName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", genderName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["GenderID"] : null;
        }

        public object? GetContractTypeIDByName(string contractTypeName)
        {
            string query = "SELECT ContractTypeID FROM ContractType WHERE ContractTypeName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", contractTypeName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["ContractTypeID"] : null;
        }

        public object? GetJobLevelIDByName(string jobLevelName)
        {
            string query = "SELECT JobLevelID FROM JobLevel WHERE JobLevelName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", jobLevelName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["JobLevelID"] : null;
        }

        public object? GetCompanyIDByName(string companyName)
        {
            string query = "SELECT CompanyID FROM Company WHERE CompanyName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", companyName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["CompanyID"] : null;
        }

        public object? GetDegreeIDByName(string degreeName)
        {
            string query = "SELECT DegreeID FROM Degree WHERE DegreeName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", degreeName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["DegreeID"] : null;
        }

        public object? GetDegreeFieldIDByName(string degreeFieldName)
        {
            string query = "SELECT DegreeFieldID FROM DegreeField WHERE DegreeFieldName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", degreeFieldName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["DegreeFieldID"] : null;
        }

        public object? GetChartIDByName(string chartName)
        {
            string query = "SELECT ChartID FROM ChartAffairs WHERE ChartName = ?";
            OleDbParameter[] parameters = new OleDbParameter[] { new OleDbParameter("?", chartName) };
            DataTable? dt = ExecuteQuery(query, parameters);
            return dt?.Rows.Count > 0 ? dt.Rows[0]["ChartID"] : null;
        }
    }
}