using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace RCSVB.Models
{
    public class Department    
    {
        public static readonly int OwnerDepth = 1;

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
        // float actualTotal = dept.Total(account => account.Actual);
        // float budgetTotal = dept.Total(account => account.Budget);
        public double Total (Func<Account, double> method)
        {
            double total = 0;

            foreach (Account account in Accounts) {
                total += method (account);
            }

            foreach (Department department in Departments) {
                total += department.Total (method);
            }

            return total;
        }

        public void PrintExcelRows(Worksheet worksheet, ref int row, Func<Account, double> method, string section)
        {
            string departmentOwnerName = DepartmentOwnerName();
            string departmentName = Name;

            int groupStartRow = row;

            foreach (Account account in Accounts) {
                worksheet.Cells[row, 1] = departmentOwnerName;
                worksheet.Cells[row, 2] = departmentName;
                worksheet.Cells[row, 3] = account.Name;
                worksheet.Cells[row, 11] = method(account);
                ++row;
            }

            foreach(Department department in Departments)
            {
                department.PrintExcelRows(worksheet, ref row, method, section);
            }

            if (Depth == 1) {
                worksheet.Cells[row, 1] = departmentName + " Total";
                worksheet.Cells[row, 11] = Total (method);
                ((Range) worksheet.Rows[row]).Font.Bold = true;
                GroupRows (worksheet, groupStartRow, row - 1);
                ++row;
            }
            else if (Depth == 0)
            {
                worksheet.Cells[row, 1] = section + " Grand Total";
                worksheet.Cells[row, 11] = Total(method);
                ((Range)worksheet.Rows[row]).Font.Bold = true;
                GroupRows(worksheet, groupStartRow, row - 1);
                ++row;
            }

            if (Accounts.Count > 0)
            {
                worksheet.Cells[row, 2] = departmentName + " Total";
                worksheet.Cells[row, 11] = Total(method); // $"=SUM(K{groupStartRow}:K{row - 1})";
                ((Range)worksheet.Rows[row]).Font.Bold = true;
                GroupRows(worksheet, groupStartRow, row - 1);
                ++row;
            }
        }

        public string DepartmentOwnerName() 
        {
            Department department = this;
            while (department.Depth > OwnerDepth) {
                department = department.ParentDepartment;
            }
            return department.Name;
        }

        private static void GroupRows(_Worksheet worksheet, int start, int end) {
            worksheet.Rows[$"{start}:{end}"].Group ();
        }
    }
}
