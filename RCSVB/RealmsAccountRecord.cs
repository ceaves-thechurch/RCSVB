using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace RCSVB
{
    public class RealmsAccountRecord
    {
        private static Regex _accountRecordValidation = new Regex(@"^[0-9]{5}");
        private static Regex _totalRecordValidation = new Regex(@"^Total");

        public string Owner { get; set; }
        public string Department { get; set; }
        public string Account { get; set; }
        public string Actual { get; set; }
        public string Budget { get; set; }
        public string Variance { get; set; }

        public bool IsValidAccountRecord
        {
            get
            {
                return _accountRecordValidation.IsMatch(Account);
            }
        }

        public bool IsValidTotalRecord
        {
            get
            {
                return _totalRecordValidation.IsMatch(Account);
            }
        }

        internal void TrimCSVFields()
        {
            Account = Account.Trim();
            Actual = Actual.Trim();
            Budget = Budget.Trim();
            Variance = Variance.Trim();
        }
    }
}
