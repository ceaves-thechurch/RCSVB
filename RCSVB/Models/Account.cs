using System.Globalization;

namespace RCSVB.Models
{
    public class Account
    {
        public string Name { get; set; }
        public Department Department { get; set; }

        public float Actual { get; set; }
        public float Budget { get; set; }
        public float Variance { get; set; }

        private static NumberStyles _style = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;
        private static CultureInfo _culture = CultureInfo.CreateSpecificCulture("en-US");

        public Account(RealmsRecord record, Department department)
        {
            Department = department;
            Department.Accounts.Add(this);

            Name = record.Account.Trim();
            Actual = float.TryParse(record.Actual, _style, _culture, out float actual) ? actual : 0f;
            Budget = float.TryParse(record.Budget, _style, _culture, out float budget) ? budget : 0f;
            Variance = float.TryParse(record.Variance, _style, _culture, out float variance) ? variance : 0f;
        }

        public string DepartmentOwnerName ()
        {
            Department department = Department;
            while (department.Depth > Department.RootDepth)
            {
                department = department.ParentDepartment;
            }
            return department.Name;
        }

        public string DepartmentName ()
        {
            return Department.Name;
        }
    }
}
