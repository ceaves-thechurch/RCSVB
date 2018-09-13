namespace RCSVB.Models
{
    public class Account
    {
        public string Name { get; set; }
        public Department Department { get; set; }

        public float Actual { get; set; }
        public float Budget { get; set; }
        public float Variance { get; set; }

        public Account(Department department)
        {
            Department = department;
            Department.Accounts.Add(this);
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
