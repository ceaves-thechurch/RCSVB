using System;
using System.Collections.Generic;

namespace RCSVB.Models
{
    public class Department    
    {
        public static readonly int RootDepth = 1;

        public string Name { get; set; }
        public int Depth { get; set; }
        public Department ParentDepartment { get; set; }

        public List<Account> Accounts { get; set; }
        public List<Department> Departments { get; set; }

        public Department(string name)
        {
            Name = name;
            Depth = 0;
            ParentDepartment = null;

            Accounts = new List<Account>();
            Departments = new List<Department>();
        }

        public Department(string name, Department parent)
        {
            Name = name;
            Depth = parent.Depth + 1;
            ParentDepartment = parent;
            parent.Departments.Add(this);

            Accounts = new List<Account>();
            Departments = new List<Department>();
        }

        // Usage:
        // float actualTotal = dept.Total(account => account.Actual);
        // float budgetTotal = dept.Total(account => account.Budget);
        public float Total (Func<Account, float> method)
        {
            return AccountsTotal(method) + DepartmentsTotal(method);
        }

        public float AccountsTotal(Func<Account, float> method)
        {
            float total = 0f;
            foreach (Account account in Accounts)
            {
                total += method(account);
            }
            return total;
        }

        public float DepartmentsTotal(Func<Account, float> method)
        {
            float total = 0f;
            foreach (Department department in Departments)
            {
                total += department.Total(method);
            }
            return total;
        }
    }
}
