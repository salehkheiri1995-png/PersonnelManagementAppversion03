using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace PersonnelManagementApp
{
    // ========== AnalyticsDataModel ==========
    public class AnalyticsDataModel
    {
        private List<PersonnelRecord> personnelList = new List<PersonnelRecord>();
        private List<string> filteredProvinces = new List<string>();
        private List<string> filteredCities = new List<string>();
        private List<string> filteredAffairs = new List<string>();
        private List<string> filteredDepartments = new List<string>();
        private List<string> filteredDistricts = new List<string>();
        private List<string> filteredPositions = new List<string>();
        private List<string> filteredGenders = new List<string>();
        private List<string> filteredEducations = new List<string>();
        private List<string> filteredJobLevels = new List<string>();
        private List<string> filteredContractTypes = new List<string>();
        private List<string> filteredCompanies = new List<string>();
        private List<string> filteredWorkShifts = new List<string>();
        private DateTime? hireDateFrom = null;
        private DateTime? hireDateTo = null;
        private int? minAge = null;
        private int? maxAge = null;
        private int? minExperience = null;
        private int? maxExperience = null;

        private readonly Dictionary<int, string> provinceCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> cityCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> affairCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> departmentCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> districtCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> positionCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> genderCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> degreeCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> jobLevelCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> contractTypeCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> companyCache = new Dictionary<int, string>();
        private readonly Dictionary<int, string> workShiftCache = new Dictionary<int, string>();

        public int TotalPersonnel { get; private set; }
        public int DepartmentCount { get; private set; }
        public int PositionCount { get; private set; }
        public int ProvinceCount { get; private set; }
        public int CompanyCount { get; private set; }
        public int JobLevelCount { get; private set; }
        public int ContractTypeCount { get; private set; }
        public int EducationCount { get; private set; }
        public int WorkShiftCount { get; private set; }
        public int MaleCount { get; private set; }
        public int FemaleCount { get; private set; }

        public bool LoadData(DbHelper dbHelper)
        {
            try
            {
                LoadAllCaches(dbHelper);

                DataTable dt = dbHelper.ExecuteQuery(@"SELECT PersonnelID, ProvinceID, CityID, AffairID, DeptID, DistrictID, PostNameID, 
                    VoltageID, WorkShiftID, GenderID, FirstName, LastName, FatherName, PersonnelNumber, NationalID, MobileNumber, 
                    BirthDate, HireDate, StartDateOperation, ContractTypeID, JobLevelID, CompanyID, DegreeID, DegreeFieldID, 
                    MainJobTitle, CurrentActivity, StatusID FROM Personnel");

                if (dt?.Rows.Count == 0) return false;

                personnelList.Clear();
                foreach (DataRow row in dt.Rows)
                {
                    personnelList.Add(new PersonnelRecord
                    {
                        PersonnelID = Convert.ToInt32(row["PersonnelID"]),
                        ProvinceID = GetIntValue(row["ProvinceID"]),
                        CityID = GetIntValue(row["CityID"]),
                        AffairID = GetIntValue(row["AffairID"]),
                        DeptID = GetIntValue(row["DeptID"]),
                        DistrictID = GetIntValue(row["DistrictID"]),
                        PostNameID = GetIntValue(row["PostNameID"]),
                        VoltageID = GetIntValue(row["VoltageID"]),
                        WorkShiftID = GetIntValue(row["WorkShiftID"]),
                        GenderID = GetIntValue(row["GenderID"]),
                        FirstName = row["FirstName"]?.ToString() ?? "",
                        LastName = row["LastName"]?.ToString() ?? "",
                        FatherName = row["FatherName"]?.ToString() ?? "",
                        PersonnelNumber = row["PersonnelNumber"]?.ToString() ?? "",
                        NationalID = row["NationalID"]?.ToString() ?? "",
                        MobileNumber = row["MobileNumber"]?.ToString() ?? "",
                        BirthDate = GetDateValue(row["BirthDate"]),
                        HireDate = GetDateValue(row["HireDate"]),
                        StartDateOperation = GetDateValue(row["StartDateOperation"]),
                        ContractTypeID = GetIntValue(row["ContractTypeID"]),
                        JobLevelID = GetIntValue(row["JobLevelID"]),
                        CompanyID = GetIntValue(row["CompanyID"]),
                        DegreeID = GetIntValue(row["DegreeID"]),
                        DegreeFieldID = GetIntValue(row["DegreeFieldID"]),
                        MainJobTitle = GetIntValue(row["MainJobTitle"]),
                        CurrentActivity = GetIntValue(row["CurrentActivity"]),
                        StatusID = GetIntValue(row["StatusID"])
                    });
                }

                CalculateStatistics();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ خطا در بارگذاری دادهها: {ex.Message}");
                return false;
            }
        }

        private int GetIntValue(object value) => value != DBNull.Value && value != null ? Convert.ToInt32(value) : 0;
        private DateTime? GetDateValue(object value) => value != DBNull.Value && value != null ? Convert.ToDateTime(value) : (DateTime?)null;

        private void LoadAllCaches(DbHelper dbHelper)
        {
            LoadCache(dbHelper, "SELECT ProvinceID, ProvinceName FROM Provinces", provinceCache);
            LoadCache(dbHelper, "SELECT CityID, CityName FROM Cities", cityCache);
            LoadCache(dbHelper, "SELECT AffairID, AffairName FROM TransferAffairs", affairCache);
            LoadCache(dbHelper, "SELECT DeptID, DeptName FROM OperationDepartments", departmentCache);
            LoadCache(dbHelper, "SELECT DistrictID, DistrictName FROM Districts", districtCache);
            LoadCache(dbHelper, "SELECT PostNameID, PostName FROM PostsNames", positionCache);
            LoadCache(dbHelper, "SELECT GenderID, GenderName FROM Gender", genderCache);
            LoadCache(dbHelper, "SELECT DegreeID, DegreeName FROM Degree", degreeCache);
            LoadCache(dbHelper, "SELECT JobLevelID, JobLevelName FROM JobLevel", jobLevelCache);
            LoadCache(dbHelper, "SELECT ContractTypeID, ContractTypeName FROM ContractType", contractTypeCache);
            LoadCache(dbHelper, "SELECT CompanyID, CompanyName FROM Company", companyCache);
            LoadCache(dbHelper, "SELECT WorkShiftID, WorkShiftName FROM WorkShift", workShiftCache);
        }

        private void LoadCache(DbHelper dbHelper, string query, Dictionary<int, string> cache)
        {
            try
            {
                DataTable dt = dbHelper.ExecuteQuery(query);
                if (dt == null) return;

                string keyColumn = dt.Columns[0].ColumnName;
                string valueColumn = dt.Columns[1].ColumnName;

                foreach (DataRow row in dt.Rows)
                {
                    int key = Convert.ToInt32(row[keyColumn]);
                    string value = row[valueColumn]?.ToString() ?? "";
                    cache[key] = value;
                }
            }
            catch { }
        }

        private void CalculateStatistics()
        {
            TotalPersonnel = personnelList.Count;
            DepartmentCount = personnelList.Select(p => p.DeptID).Distinct().Count(x => x > 0);
            PositionCount = personnelList.Select(p => p.PostNameID).Distinct().Count(x => x > 0);
            ProvinceCount = personnelList.Select(p => p.ProvinceID).Distinct().Count(x => x > 0);
            CompanyCount = personnelList.Select(p => p.CompanyID).Distinct().Count(x => x > 0);
            JobLevelCount = personnelList.Select(p => p.JobLevelID).Distinct().Count(x => x > 0);
            ContractTypeCount = personnelList.Select(p => p.ContractTypeID).Distinct().Count(x => x > 0);
            EducationCount = personnelList.Select(p => p.DegreeID).Distinct().Count(x => x > 0);
            WorkShiftCount = personnelList.Select(p => p.WorkShiftID).Distinct().Count(x => x > 0);
            MaleCount = personnelList.Count(p => p.GenderID == 1);
            FemaleCount = personnelList.Count(p => p.GenderID == 2);
        }

        public void SetFilters(List<string> provinces, List<string> cities, List<string> affairs, List<string> depts,
            List<string> districts, List<string> positions, List<string> genders, List<string> educations,
            List<string> jobLevels, List<string> contractTypes, List<string> companies, List<string> workShifts,
            DateTime? hireFrom, DateTime? hireTo, int? ageMin = null, int? ageMax = null, 
            int? expMin = null, int? expMax = null)
        {
            filteredProvinces = provinces ?? new List<string>();
            filteredCities = cities ?? new List<string>();
            filteredAffairs = affairs ?? new List<string>();
            filteredDepartments = depts ?? new List<string>();
            filteredDistricts = districts ?? new List<string>();
            filteredPositions = positions ?? new List<string>();
            filteredGenders = genders ?? new List<string>();
            filteredEducations = educations ?? new List<string>();
            filteredJobLevels = jobLevels ?? new List<string>();
            filteredContractTypes = contractTypes ?? new List<string>();
            filteredCompanies = companies ?? new List<string>();
            filteredWorkShifts = workShifts ?? new List<string>();
            hireDateFrom = hireFrom;
            hireDateTo = hireTo;
            minAge = ageMin;
            maxAge = ageMax;
            minExperience = expMin;
            maxExperience = expMax;
        }

        public void ClearFilters()
        {
            filteredProvinces.Clear();
            filteredCities.Clear();
            filteredAffairs.Clear();
            filteredDepartments.Clear();
            filteredDistricts.Clear();
            filteredPositions.Clear();
            filteredGenders.Clear();
            filteredEducations.Clear();
            filteredJobLevels.Clear();
            filteredContractTypes.Clear();
            filteredCompanies.Clear();
            filteredWorkShifts.Clear();
            hireDateFrom = null;
            hireDateTo = null;
            minAge = null;
            maxAge = null;
            minExperience = null;
            maxExperience = null;
        }

        private int CalculateAge(DateTime? birthDate)
        {
            if (!birthDate.HasValue) return 0;
            var today = DateTime.Today;
            int age = today.Year - birthDate.Value.Year;
            if (birthDate.Value.Date > today.AddYears(-age)) age--;
            return age;
        }

        private int CalculateWorkExperience(DateTime? hireDate)
        {
            if (!hireDate.HasValue) return 0;
            var today = DateTime.Today;
            int years = today.Year - hireDate.Value.Year;
            if (hireDate.Value.Date > today.AddYears(-years)) years--;
            return years < 0 ? 0 : years;
        }

        private List<PersonnelRecord> GetFiltered()
        {
            var result = personnelList.AsEnumerable();

            if (filteredProvinces.Count > 0)
                result = result.Where(p => filteredProvinces.Contains(provinceCache.ContainsKey(p.ProvinceID) ? provinceCache[p.ProvinceID] : ""));

            if (filteredCities.Count > 0)
                result = result.Where(p => filteredCities.Contains(cityCache.ContainsKey(p.CityID) ? cityCache[p.CityID] : ""));

            if (filteredAffairs.Count > 0)
                result = result.Where(p => filteredAffairs.Contains(affairCache.ContainsKey(p.AffairID) ? affairCache[p.AffairID] : ""));

            if (filteredDepartments.Count > 0)
                result = result.Where(p => filteredDepartments.Contains(departmentCache.ContainsKey(p.DeptID) ? departmentCache[p.DeptID] : ""));

            if (filteredDistricts.Count > 0)
                result = result.Where(p => filteredDistricts.Contains(districtCache.ContainsKey(p.DistrictID) ? districtCache[p.DistrictID] : ""));

            if (filteredPositions.Count > 0)
                result = result.Where(p => filteredPositions.Contains(positionCache.ContainsKey(p.PostNameID) ? positionCache[p.PostNameID] : ""));

            if (filteredGenders.Count > 0)
                result = result.Where(p => filteredGenders.Contains(genderCache.ContainsKey(p.GenderID) ? genderCache[p.GenderID] : ""));

            if (filteredEducations.Count > 0)
                result = result.Where(p => filteredEducations.Contains(degreeCache.ContainsKey(p.DegreeID) ? degreeCache[p.DegreeID] : ""));

            if (filteredJobLevels.Count > 0)
                result = result.Where(p => filteredJobLevels.Contains(jobLevelCache.ContainsKey(p.JobLevelID) ? jobLevelCache[p.JobLevelID] : ""));

            if (filteredContractTypes.Count > 0)
                result = result.Where(p => filteredContractTypes.Contains(contractTypeCache.ContainsKey(p.ContractTypeID) ? contractTypeCache[p.ContractTypeID] : ""));

            if (filteredCompanies.Count > 0)
                result = result.Where(p => filteredCompanies.Contains(companyCache.ContainsKey(p.CompanyID) ? companyCache[p.CompanyID] : ""));

            if (filteredWorkShifts.Count > 0)
                result = result.Where(p => filteredWorkShifts.Contains(workShiftCache.ContainsKey(p.WorkShiftID) ? workShiftCache[p.WorkShiftID] : ""));

            if (hireDateFrom.HasValue && hireDateTo.HasValue)
                result = result.Where(p => p.HireDate.HasValue && p.HireDate >= hireDateFrom && p.HireDate <= hireDateTo);

            if (minAge.HasValue || maxAge.HasValue)
            {
                result = result.Where(p =>
                {
                    int age = CalculateAge(p.BirthDate);
                    if (age == 0) return false;
                    if (minAge.HasValue && age < minAge.Value) return false;
                    if (maxAge.HasValue && age > maxAge.Value) return false;
                    return true;
                });
            }

            if (minExperience.HasValue || maxExperience.HasValue)
            {
                result = result.Where(p =>
                {
                    int exp = CalculateWorkExperience(p.HireDate);
                    if (minExperience.HasValue && exp < minExperience.Value) return false;
                    if (maxExperience.HasValue && exp > maxExperience.Value) return false;
                    return true;
                });
            }

            return result.ToList();
        }

        public int GetFilteredTotal() => GetFiltered().Count;
        public int GetFilteredDepartmentCount() => GetFiltered().Select(p => p.DeptID).Distinct().Count(x => x > 0);
        public int GetFilteredPositionCount() => GetFiltered().Select(p => p.PostNameID).Distinct().Count(x => x > 0);
        public int GetFilteredMaleCount() => GetFiltered().Count(p => p.GenderID == 1);
        public int GetFilteredFemaleCount() => GetFiltered().Count(p => p.GenderID == 2);

        public List<string> GetAllProvinces() => provinceCache.Values.Distinct().OrderBy(x => x).ToList();
        public List<string> GetAllGenders() => genderCache.Values.Distinct().OrderBy(x => x).ToList();
        public List<string> GetAllEducations() => degreeCache.Values.Distinct().OrderBy(x => x).ToList();
        public List<string> GetAllJobLevels() => jobLevelCache.Values.Distinct().OrderBy(x => x).ToList();
        public List<string> GetAllContractTypes() => contractTypeCache.Values.Distinct().OrderBy(x => x).ToList();
        public List<string> GetAllCompanies() => companyCache.Values.Distinct().OrderBy(x => x).ToList();
        public List<string> GetAllWorkShifts() => workShiftCache.Values.Distinct().OrderBy(x => x).ToList();

        public List<string> GetCitiesByProvinces(List<string> provinces)
        {
            var provinceIds = provinceCache.Where(p => provinces.Contains(p.Value)).Select(p => p.Key).ToList();
            return personnelList.Where(p => provinceIds.Contains(p.ProvinceID) && p.CityID > 0)
                .Select(p => cityCache.ContainsKey(p.CityID) ? cityCache[p.CityID] : "")
                .Where(x => !string.IsNullOrEmpty(x)).Distinct().OrderBy(x => x).ToList();
        }

        public List<string> GetAffairsByProvinces(List<string> provinces)
        {
            var provinceIds = provinceCache.Where(p => provinces.Contains(p.Value)).Select(p => p.Key).ToList();
            return personnelList.Where(p => provinceIds.Contains(p.ProvinceID) && p.AffairID > 0)
                .Select(p => affairCache.ContainsKey(p.AffairID) ? affairCache[p.AffairID] : "")
                .Where(x => !string.IsNullOrEmpty(x)).Distinct().OrderBy(x => x).ToList();
        }

        public List<string> GetDepartmentsByFilters(List<string> provinces, List<string> cities, List<string> affairs)
        {
            var provinceIds = provinceCache.Where(p => provinces.Contains(p.Value)).Select(p => p.Key).ToList();
            var cityIds = cityCache.Where(p => cities.Contains(p.Value)).Select(p => p.Key).ToList();
            var affairIds = affairCache.Where(p => affairs.Contains(p.Value)).Select(p => p.Key).ToList();

            var filtered = personnelList.AsEnumerable();
            if (provinceIds.Count > 0) filtered = filtered.Where(p => provinceIds.Contains(p.ProvinceID));
            if (cityIds.Count > 0) filtered = filtered.Where(p => cityIds.Contains(p.CityID));
            if (affairIds.Count > 0) filtered = filtered.Where(p => affairIds.Contains(p.AffairID));

            return filtered.Where(p => p.DeptID > 0)
                .Select(p => departmentCache.ContainsKey(p.DeptID) ? departmentCache[p.DeptID] : "")
                .Where(x => !string.IsNullOrEmpty(x)).Distinct().OrderBy(x => x).ToList();
        }

        public List<string> GetDistrictsByDepartments(List<string> departments)
        {
            var deptIds = departmentCache.Where(p => departments.Contains(p.Value)).Select(p => p.Key).ToList();
            return personnelList.Where(p => deptIds.Contains(p.DeptID) && p.DistrictID > 0)
                .Select(p => districtCache.ContainsKey(p.DistrictID) ? districtCache[p.DistrictID] : "")
                .Where(x => !string.IsNullOrEmpty(x)).Distinct().OrderBy(x => x).ToList();
        }

        public List<string> GetPositionsByDistricts(List<string> districts)
        {
            var districtIds = districtCache.Where(p => districts.Contains(p.Value)).Select(p => p.Key).ToList();
            return personnelList.Where(p => districtIds.Contains(p.DistrictID) && p.PostNameID > 0)
                .Select(p => positionCache.ContainsKey(p.PostNameID) ? positionCache[p.PostNameID] : "")
                .Where(x => !string.IsNullOrEmpty(x)).Distinct().OrderBy(x => x).ToList();
        }

        public List<StatisticItem> GetFilteredDepartmentStatistics()
        {
            var filtered = GetFiltered();
            return filtered.Where(p => p.DeptID > 0).GroupBy(p => p.DeptID)
                .Select(g => new StatisticItem
                {
                    Name = departmentCache.ContainsKey(g.Key) ? departmentCache[g.Key] : $"اداره {g.Key}",
                    Count = g.Count()
                }).OrderByDescending(x => x.Count).ToList();
        }

        public List<StatisticItem> GetFilteredPositionStatistics()
        {
            var filtered = GetFiltered();
            return filtered.Where(p => p.PostNameID > 0).GroupBy(p => p.PostNameID)
                .Select(g => new StatisticItem
                {
                    Name = positionCache.ContainsKey(g.Key) ? positionCache[g.Key] : $"پست {g.Key}",
                    Count = g.Count()
                }).OrderByDescending(x => x.Count).ToList();
        }

        public List<StatisticItem> GetFilteredGenderStatistics()
        {
            var filtered = GetFiltered();
            return filtered.Where(p => p.GenderID > 0).GroupBy(p => p.GenderID)
                .Select(g => new StatisticItem
                {
                    Name = genderCache.ContainsKey(g.Key) ? genderCache[g.Key] : $"جنسیت {g.Key}",
                    Count = g.Count()
                }).OrderByDescending(x => x.Count).ToList();
        }

        public List<StatisticItem> GetFilteredJobLevelStatistics()
        {
            var filtered = GetFiltered();
            return filtered.Where(p => p.JobLevelID > 0).GroupBy(p => p.JobLevelID)
                .Select(g => new StatisticItem
                {
                    Name = jobLevelCache.ContainsKey(g.Key) ? jobLevelCache[g.Key] : $"سطح {g.Key}",
                    Count = g.Count()
                }).OrderByDescending(x => x.Count).ToList();
        }

        public List<StatisticItem> GetFilteredContractTypeStatistics()
        {
            var filtered = GetFiltered();
            return filtered.Where(p => p.ContractTypeID > 0).GroupBy(p => p.ContractTypeID)
                .Select(g => new StatisticItem
                {
                    Name = contractTypeCache.ContainsKey(g.Key) ? contractTypeCache[g.Key] : $"قرارداد {g.Key}",
                    Count = g.Count()
                }).OrderByDescending(x => x.Count).ToList();
        }

        public List<StatisticItem> GetFilteredProvinceStatistics()
        {
            var filtered = GetFiltered();
            return filtered.Where(p => p.ProvinceID > 0).GroupBy(p => p.ProvinceID)
                .Select(g => new StatisticItem
                {
                    Name = provinceCache.ContainsKey(g.Key) ? provinceCache[g.Key] : $"استان {g.Key}",
                    Count = g.Count()
                }).OrderByDescending(x => x.Count).ToList();
        }

        public List<StatisticItem> GetFilteredEducationStatistics()
        {
            var filtered = GetFiltered();
            return filtered.Where(p => p.DegreeID > 0).GroupBy(p => p.DegreeID)
                .Select(g => new StatisticItem
                {
                    Name = degreeCache.ContainsKey(g.Key) ? degreeCache[g.Key] : $"مدرک {g.Key}",
                    Count = g.Count()
                }).OrderByDescending(x => x.Count).ToList();
        }

        public List<StatisticItem> GetFilteredCompanyStatistics()
        {
            var filtered = GetFiltered();
            return filtered.Where(p => p.CompanyID > 0).GroupBy(p => p.CompanyID)
                .Select(g => new StatisticItem
                {
                    Name = companyCache.ContainsKey(g.Key) ? companyCache[g.Key] : $"شرکت {g.Key}",
                    Count = g.Count()
                }).OrderByDescending(x => x.Count).ToList();
        }

        public List<StatisticItem> GetFilteredWorkShiftStatistics()
        {
            var filtered = GetFiltered();
            return filtered.Where(p => p.WorkShiftID > 0).GroupBy(p => p.WorkShiftID)
                .Select(g => new StatisticItem
                {
                    Name = workShiftCache.ContainsKey(g.Key) ? workShiftCache[g.Key] : $"شیفت {g.Key}",
                    Count = g.Count()
                }).OrderByDescending(x => x.Count).ToList();
        }

        public List<StatisticItem> GetFilteredAgeStatistics()
        {
            var filtered = GetFiltered().Where(p => p.BirthDate.HasValue).ToList();
            var ageGroups = new Dictionary<string, int>
            {
                {"10-20 سال", 0},
                {"21-30 سال", 0},
                {"31-40 سال", 0},
                {"41-50 سال", 0},
                {"51-60 سال", 0},
                {"61-70 سال", 0},
                {"71-80 سال", 0},
                {"81-90 سال", 0},
                {"91-100 سال", 0}
            };

            foreach (var person in filtered)
            {
                int age = CalculateAge(person.BirthDate);
                if (age >= 10 && age <= 20) ageGroups["10-20 سال"]++;
                else if (age >= 21 && age <= 30) ageGroups["21-30 سال"]++;
                else if (age >= 31 && age <= 40) ageGroups["31-40 سال"]++;
                else if (age >= 41 && age <= 50) ageGroups["41-50 سال"]++;
                else if (age >= 51 && age <= 60) ageGroups["51-60 سال"]++;
                else if (age >= 61 && age <= 70) ageGroups["61-70 سال"]++;
                else if (age >= 71 && age <= 80) ageGroups["71-80 سال"]++;
                else if (age >= 81 && age <= 90) ageGroups["81-90 سال"]++;
                else if (age >= 91 && age <= 100) ageGroups["91-100 سال"]++;
            }

            return ageGroups.Where(x => x.Value > 0).Select(x => new StatisticItem { Name = x.Key, Count = x.Value }).ToList();
        }

        public List<StatisticItem> GetFilteredWorkExperienceStatistics()
        {
            var filtered = GetFiltered().Where(p => p.HireDate.HasValue).ToList();
            var expGroups = new Dictionary<string, int>
            {
                {"0-5 سال", 0},
                {"6-10 سال", 0},
                {"11-15 سال", 0},
                {"16-20 سال", 0},
                {"21-25 سال", 0},
                {"26-30 سال", 0},
                {"31-35 سال", 0},
                {"36-40 سال", 0},
                {"بیش از 40 سال", 0}
            };

            foreach (var person in filtered)
            {
                int exp = CalculateWorkExperience(person.HireDate);
                if (exp >= 0 && exp <= 5) expGroups["0-5 سال"]++;
                else if (exp >= 6 && exp <= 10) expGroups["6-10 سال"]++;
                else if (exp >= 11 && exp <= 15) expGroups["11-15 سال"]++;
                else if (exp >= 16 && exp <= 20) expGroups["16-20 سال"]++;
                else if (exp >= 21 && exp <= 25) expGroups["21-25 سال"]++;
                else if (exp >= 26 && exp <= 30) expGroups["26-30 سال"]++;
                else if (exp >= 31 && exp <= 35) expGroups["31-35 سال"]++;
                else if (exp >= 36 && exp <= 40) expGroups["36-40 سال"]++;
                else if (exp > 40) expGroups["بیش از 40 سال"]++;
            }

            return expGroups.Where(x => x.Value > 0).Select(x => new StatisticItem { Name = x.Key, Count = x.Value }).ToList();
        }

        public List<PersonnelDetail> GetPersonnelByFilter(string filterValue, Chart chart)
        {
            var filtered = GetFiltered();

            string title = chart.Titles.Count > 0 ? chart.Titles[0].Text : "";

            if (title.Contains("اداره"))
                return filtered.Where(p => p.DeptID > 0 && departmentCache.ContainsKey(p.DeptID) && departmentCache[p.DeptID] == filterValue)
                    .Select(ToDetail).ToList();

            if (title.Contains("پست"))
                return filtered.Where(p => p.PostNameID > 0 && positionCache.ContainsKey(p.PostNameID) && positionCache[p.PostNameID] == filterValue)
                    .Select(ToDetail).ToList();

            if (title.Contains("جنسیت"))
                return filtered.Where(p => p.GenderID > 0 && genderCache.ContainsKey(p.GenderID) && genderCache[p.GenderID] == filterValue)
                    .Select(ToDetail).ToList();

            if (title.Contains("سطح"))
                return filtered.Where(p => p.JobLevelID > 0 && jobLevelCache.ContainsKey(p.JobLevelID) && jobLevelCache[p.JobLevelID] == filterValue)
                    .Select(ToDetail).ToList();

            if (title.Contains("قرارداد"))
                return filtered.Where(p => p.ContractTypeID > 0 && contractTypeCache.ContainsKey(p.ContractTypeID) && contractTypeCache[p.ContractTypeID] == filterValue)
                    .Select(ToDetail).ToList();

            if (title.Contains("استان"))
                return filtered.Where(p => p.ProvinceID > 0 && provinceCache.ContainsKey(p.ProvinceID) && provinceCache[p.ProvinceID] == filterValue)
                    .Select(ToDetail).ToList();

            if (title.Contains("مدارک") || title.Contains("تحصیلات"))
                return filtered.Where(p => p.DegreeID > 0 && degreeCache.ContainsKey(p.DegreeID) && degreeCache[p.DegreeID] == filterValue)
                    .Select(ToDetail).ToList();

            if (title.Contains("شرکت"))
                return filtered.Where(p => p.CompanyID > 0 && companyCache.ContainsKey(p.CompanyID) && companyCache[p.CompanyID] == filterValue)
                    .Select(ToDetail).ToList();

            if (title.Contains("شیفت"))
                return filtered.Where(p => p.WorkShiftID > 0 && workShiftCache.ContainsKey(p.WorkShiftID) && workShiftCache[p.WorkShiftID] == filterValue)
                    .Select(ToDetail).ToList();

            if (title.Contains("سن"))
            {
                return filtered.Where(p => p.BirthDate.HasValue).Where(p =>
                {
                    int age = CalculateAge(p.BirthDate);
                    if (filterValue == "10-20 سال") return age >= 10 && age <= 20;
                    if (filterValue == "21-30 سال") return age >= 21 && age <= 30;
                    if (filterValue == "31-40 سال") return age >= 31 && age <= 40;
                    if (filterValue == "41-50 سال") return age >= 41 && age <= 50;
                    if (filterValue == "51-60 سال") return age >= 51 && age <= 60;
                    if (filterValue == "61-70 سال") return age >= 61 && age <= 70;
                    if (filterValue == "71-80 سال") return age >= 71 && age <= 80;
                    if (filterValue == "81-90 سال") return age >= 81 && age <= 90;
                    if (filterValue == "91-100 سال") return age >= 91 && age <= 100;
                    return false;
                }).Select(ToDetail).ToList();
            }

            if (title.Contains("سابقه"))
            {
                return filtered.Where(p => p.HireDate.HasValue).Where(p =>
                {
                    int exp = CalculateWorkExperience(p.HireDate);
                    if (filterValue == "0-5 سال") return exp >= 0 && exp <= 5;
                    if (filterValue == "6-10 سال") return exp >= 6 && exp <= 10;
                    if (filterValue == "11-15 سال") return exp >= 11 && exp <= 15;
                    if (filterValue == "16-20 سال") return exp >= 16 && exp <= 20;
                    if (filterValue == "21-25 سال") return exp >= 21 && exp <= 25;
                    if (filterValue == "26-30 سال") return exp >= 26 && exp <= 30;
                    if (filterValue == "31-35 سال") return exp >= 31 && exp <= 35;
                    if (filterValue == "36-40 سال") return exp >= 36 && exp <= 40;
                    if (filterValue == "بیش از 40 سال") return exp > 40;
                    return false;
                }).Select(ToDetail).ToList();
            }

            return new List<PersonnelDetail>();
        }

        private PersonnelDetail ToDetail(PersonnelRecord p) => new PersonnelDetail
        {
            PersonnelID = p.PersonnelID,
            FirstName = p.FirstName,
            LastName = p.LastName,
            PersonnelNumber = p.PersonnelNumber,
            NationalID = p.NationalID,
            PostName = positionCache.ContainsKey(p.PostNameID) ? positionCache[p.PostNameID] : "",
            DeptName = departmentCache.ContainsKey(p.DeptID) ? departmentCache[p.DeptID] : "",
            Province = provinceCache.ContainsKey(p.ProvinceID) ? provinceCache[p.ProvinceID] : "",
            City = cityCache.ContainsKey(p.CityID) ? cityCache[p.CityID] : "",
            Affair = affairCache.ContainsKey(p.AffairID) ? affairCache[p.AffairID] : "",
            District = districtCache.ContainsKey(p.DistrictID) ? districtCache[p.DistrictID] : "",
            ContractType = contractTypeCache.ContainsKey(p.ContractTypeID) ? contractTypeCache[p.ContractTypeID] : "",
            Gender = genderCache.ContainsKey(p.GenderID) ? genderCache[p.GenderID] : "",
            Education = degreeCache.ContainsKey(p.DegreeID) ? degreeCache[p.DegreeID] : "",
            JobLevel = jobLevelCache.ContainsKey(p.JobLevelID) ? jobLevelCache[p.JobLevelID] : "",
            Company = companyCache.ContainsKey(p.CompanyID) ? companyCache[p.CompanyID] : "",
            WorkShift = workShiftCache.ContainsKey(p.WorkShiftID) ? workShiftCache[p.WorkShiftID] : "",
            HireDate = p.HireDate,
            BirthDate = p.BirthDate,
            MobileNumber = p.MobileNumber
        };
    }

    // ========== PersonnelRecord ==========
    public class PersonnelRecord
    {
        public int PersonnelID { get; set; }
        public int ProvinceID { get; set; }
        public int CityID { get; set; }
        public int AffairID { get; set; }
        public int DeptID { get; set; }
        public int DistrictID { get; set; }
        public int PostNameID { get; set; }
        public int VoltageID { get; set; }
        public int WorkShiftID { get; set; }
        public int GenderID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string FatherName { get; set; }
        public string PersonnelNumber { get; set; }
        public string NationalID { get; set; }
        public string MobileNumber { get; set; }
        public DateTime? BirthDate { get; set; }
        public DateTime? HireDate { get; set; }
        public DateTime? StartDateOperation { get; set; }
        public int ContractTypeID { get; set; }
        public int JobLevelID { get; set; }
        public int CompanyID { get; set; }
        public int DegreeID { get; set; }
        public int DegreeFieldID { get; set; }
        public int MainJobTitle { get; set; }
        public int CurrentActivity { get; set; }
        public int StatusID { get; set; }
    }

    // ========== StatisticItem ==========
    public class StatisticItem
    {
        public string Name { get; set; }
        public int Count { get; set; }
    }

    // ========== PersonnelDetail ==========
    public class PersonnelDetail
    {
        public int PersonnelID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string PersonnelNumber { get; set; }
        public string NationalID { get; set; }
        public string PostName { get; set; }
        public string DeptName { get; set; }
        public string Province { get; set; }
        public string City { get; set; }
        public string Affair { get; set; }
        public string District { get; set; }
        public string ContractType { get; set; }
        public string Gender { get; set; }
        public string Education { get; set; }
        public string JobLevel { get; set; }
        public string Company { get; set; }
        public string WorkShift { get; set; }
        public DateTime? HireDate { get; set; }
        public DateTime? BirthDate { get; set; }
        public string MobileNumber { get; set; }
        public decimal Salary { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
    }
}