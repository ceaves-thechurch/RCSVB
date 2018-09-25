using System;
using System.Collections.Generic;
using System.Linq;
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

        public List<double> Actuals;
        public List<double> Budgets;
        public List<double> Variances;

        private static readonly NumberStyles _style = NumberStyles.Number | 
                                             NumberStyles.AllowCurrencySymbol |
                                             NumberStyles.AllowDecimalPoint |
                                             NumberStyles.AllowParentheses;
        private static readonly CultureInfo _culture = CultureInfo.CreateSpecificCulture("en-US");

        // Ctor
        public Account(RealmsRecord record, Department department)
        {
            Name = record.Account.Trim();

            Department = department;
            Department.Accounts.Add(this);

            Actuals = new List<double> ();
            Budgets = new List<double> ();
            Variances = new List<double> ();
        }

        public void SetActual (string actual)
        {
            Actual = double.TryParse(actual, _style, _culture, out double a) ? a : 0f;
        }

        public void SetActual(string actual, int campus)
        {
            while (Actuals.Count < campus)
            {
                Actuals.Add(0);
            }
            Actuals.Add(double.TryParse(actual, _style, _culture, out double a) ? a : 0f);
        }

        public void SetBudget (string budget)
        {
            Budget = double.TryParse(budget, _style, _culture, out double b) ? b : 0f;
        }

        public void SetBudget(string budget, int campus)
        {
            while (Budgets.Count < campus)
            {
                Budgets.Add(0);
            }
            Budgets.Add(double.TryParse(budget, _style, _culture, out double b) ? b : 0f);
        }

        public void SetVariance (string variance)
        {
            Variance = double.TryParse(variance, _style, _culture, out double v) ? v : 0f;
        }

        public void SetVariance(string variance, int campus)
        {
            while (Variances.Count < campus)
            {
                Variances.Add(0);
            }
            Variances.Add(double.TryParse(variance, _style, _culture, out double v) ? v : 0f);
        }

        // var actualsTotal = myAccount.Total (account => account.Actuals);
        public double Total (Func<Account, List<double>> method)
        {
            return method(this).Sum ();
        }
    }
}
