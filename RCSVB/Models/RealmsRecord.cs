using System.Text.RegularExpressions;

namespace RCSVB.Models
{
    public class RealmsRecord
    {
        public string Account { get; set; }
        public string Actual { get; set; }
        public string Budget { get; set; }
        public string Variance { get; set; }

        // String starting with 0 or more spaces followed by exactly 5 digits
        private static Regex _accountRecordValidation()
        {
            return new Regex(@"^\s*[0-9]{5}");
        }

        // String starting with 0 or more spaces followed by "Total <department name>"
        private static Regex _totalRecordValidation(string departmentName)
        {
            return new Regex(@"^\s*Total\s" + departmentName);
        }

        public bool IsAccountRecord ()
        {
            if (!string.IsNullOrEmpty(Account) && 
                !string.IsNullOrEmpty(Actual) && 
                !string.IsNullOrEmpty(Budget) && 
                !string.IsNullOrEmpty(Variance) &&
                _accountRecordValidation().IsMatch(Account))
            {
                return true;
            }
            return false;
        }

        public bool IsDepartmentHeading ()
        {
            if (!string.IsNullOrEmpty(Account) &&
                string.IsNullOrEmpty(Actual) &&
                string.IsNullOrEmpty(Budget) &&
                string.IsNullOrEmpty(Variance))
            {
                return true;
            }
            return false;
        }

        public bool IsDepartmentTotalRow(Department department)
        {
            if (!string.IsNullOrEmpty(Account) &&
                !string.IsNullOrEmpty(Actual) &&
                !string.IsNullOrEmpty(Budget) &&
                !string.IsNullOrEmpty(Variance) &&
                _totalRecordValidation(department.Name).IsMatch(Account))
            {
                return true;
            }
            return false;
        }
    }
}
