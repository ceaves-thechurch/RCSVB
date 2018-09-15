using System.Globalization;

namespace RCSVB.Models
{
    public class Account
    {
        public string Name { get; set; }
        public Department Department { get; set; }

        public double Actual { get; set; }
        public double Budget { get; set; }
        public double Variance { get; set; }

        private static readonly NumberStyles _style = NumberStyles.Number | 
                                             NumberStyles.AllowCurrencySymbol |
                                             NumberStyles.AllowDecimalPoint |
                                             NumberStyles.AllowParentheses;
        private static readonly CultureInfo _culture = CultureInfo.CreateSpecificCulture("en-US");

        // Ctor
        public Account(RealmsRecord record, Department department)
        {
            Department = department;
            Department.Accounts.Add(this);

            Name = record.Account.Trim();
            Actual = double.TryParse(record.Actual, _style, _culture, out double actual) ? actual : 0f;
            Budget = double.TryParse(record.Budget, _style, _culture, out double budget) ? budget : 0f;
            Variance = double.TryParse(record.Variance, _style, _culture, out double variance) ? variance : 0f;
        }
    }
}
