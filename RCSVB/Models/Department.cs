using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace RCSVB.Models
{
    public class Department    
    {
        public static readonly int OwnerDepth = 2;

        public string Name { get; set; }
        public int Depth { get; set; }
        public Department ParentDepartment { get; set; }

        public List<Account> Accounts { get; set; }
        public List<Department> Departments { get; set; }

        public Department(string name, Department parent)
        {
            Name = name.Trim();
            ParentDepartment = parent;
            if (parent != null)
            {
                Depth = parent.Depth + 1;
                parent.Departments.Add(this);
            }
            else
            {
                Depth = 0;
            }

            Accounts = new List<Account>();
            Departments = new List<Department>();
        }

        // Usage:
        // float actualTotal = dept.Total(account => account.Actuals);
        // float budgetTotal = dept.Total(account => account.Budget);
        public double Total (Func<Account, List<double>> method)
        {
            double total = 0;

            foreach (Account account in Accounts) {
                total += account.Total (method);
            }

            foreach (Department department in Departments) {
                total += department.Total (method);
            }

            return total;
        }

        public double CampusTotal(Func<Account, List<double>> method, int campus)
        {
            double total = 0;

            foreach (Account account in Accounts)
            {
                List<double> campusValues = method(account);
                if (campus < campusValues.Count)
                {
                    total += campusValues[campus];
                }
            }

            foreach (Department department in Departments)
            {
                total += department.CampusTotal(method, campus);
            }

            return total;
        }

        public void PrintExcelRows(Worksheet worksheet, ref int row, Func<Account, List<double>> method, string section)
        {
            int groupStartRow = row;

            Range departmentOwnerNameRange = worksheet.Range[worksheet.Cells[row, 1], worksheet.Cells[row + Accounts.Count - 1, 1]];
            Range departmentNameRange = worksheet.Range[worksheet.Cells[row, 2], worksheet.Cells[row + Accounts.Count - 1, 2]];

            if (Accounts.Count > 0)
            {
                departmentOwnerNameRange.Value = DepartmentOwnerName ();
                departmentNameRange.Value = Name;
            }

            foreach (Account account in Accounts) {
                worksheet.Cells[row, 3] = account.Name;

                Range range = worksheet.Range[worksheet.Cells[row, 4], worksheet.Cells[row, 11]];
                range.Value = 0;
                range.NumberFormat = "$#,###.00";

                double accountTotal = 0;
                for (int index = 0; index < account.Actuals.Count; ++index)
                {
                    worksheet.Cells[row, 4 + index] = method(account)[index];
                    accountTotal += method(account)[index];
                }

                worksheet.Cells[row, 11].Value = accountTotal;
                ++row;
            }

            foreach(Department department in Departments)
            {
                department.PrintExcelRows(worksheet, ref row, method, section);
            }

            if (Depth == OwnerDepth) {
                worksheet.Cells[row, 1] = Name + " Total";

                // Default and format cells
                Range range = worksheet.Range[worksheet.Cells[row, 4], worksheet.Cells[row, 11]];
                range.Value = 0;
                range.NumberFormat = "$#,###.00";

                // Populate per-campus totals
                for (int i = 0; i < 7; ++i)
                {
                    worksheet.Cells[row, 4 + i] = CampusTotal(method, i);
                }

                // Populate all campuses total
                worksheet.Cells[row, 11] = Total(method);

                ((Range) worksheet.Rows[row]).Font.Bold = true;
                GroupRows (worksheet, groupStartRow, row - 1);
                ++row;
            }
            else if (Depth == OwnerDepth - 1)
            {
                worksheet.Cells[row, 1] = section + " Grand Total";

                // Default and format cells
                Range range = worksheet.Range[worksheet.Cells[row, 4], worksheet.Cells[row, 11]];
                range.Value = 0;
                range.NumberFormat = "$#,###.00";

                // Populate per-campus totals
                for (int i = 0; i < 7; ++i)
                {
                    worksheet.Cells[row, 4 + i] = CampusTotal(method, i);
                }

                // Populate all campuses total
                worksheet.Cells[row, 11] = Total(method);

                ((Range)worksheet.Rows[row]).Font.Bold = true;
                GroupRows(worksheet, groupStartRow, row - 1);
                ++row;
            }

            if (Accounts.Count > 0)
            {
                worksheet.Cells[row, 2] = Name + " Total";

                // Default and format cells
                Range range = worksheet.Range[worksheet.Cells[row, 4], worksheet.Cells[row, 11]];
                range.Value = 0;
                range.NumberFormat = "$#,###.00";

                // Populate per-campus totals
                for (int i = 0; i < 7; ++i)
                {
                    worksheet.Cells[row, 4 + i] = CampusTotal(method, i);
                }

                // Populate all campuses total
                worksheet.Cells[row, 11] = Total(method);

                ((Range)worksheet.Rows[row]).Font.Bold = true;
                GroupRows(worksheet, groupStartRow, row - 1);
                ++row;
            }
        }

        public Account GetOrCreateAccount(RealmsRecord record, Department currentDepartment)
        {
            var account = Accounts.SingleOrDefault(a => a.Name == record.Account.Trim());
            if (account == null)
            {
                return new Account(record, currentDepartment);
            }
            return account;
        }

        public string DepartmentOwnerName() 
        {
            Department department = this;
            while (department.Depth > OwnerDepth) {
                department = department.ParentDepartment;
            }
            return department.Name;
        }

        public Department GetOrCreateDepartment(string name, Department currentDepartment)
        {
            var department = Departments.SingleOrDefault(d => d.Name == name.Trim ());
            if (department == null)
            {
                return new Department(name, currentDepartment);
            }
            return department;
        }

        private static void GroupRows(_Worksheet worksheet, int start, int end) {
            worksheet.Rows[$"{start}:{end}"].Group ();
        }
    }
}
