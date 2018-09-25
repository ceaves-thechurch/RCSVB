using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using CsvHelper;
using RCSVB.Models;
using System.Reflection;

namespace RCSVB
{
    class ExcelBuilder
    {
        public static int CreateFromRealmsCSV(string source, string destination)
        {
            var app = CreateExcelApp(destination);
            app.ScreenUpdating = false;

            Workbooks workbooks = app.Workbooks;
            Workbook workbook = workbooks.Add();
            Worksheet worksheet = (Worksheet)workbook.Worksheets.Item[1];

            CreateWorksheetTemplate (worksheet);

            var csv = ConfigureCSV(source);

            // Reading records the old way
            var departmentRoot = ReadRealmsRecords(csv);

            // Writing records the old way
            PopulateTemplate(worksheet, departmentRoot);

            workbook.SaveAs (destination);

            // Cleanup
            app.ScreenUpdating = true;
            Cleanup(app, workbooks, workbook, worksheet);
            return 0;
        }

        private static Application CreateExcelApp(string destination)
        {
            if (string.IsNullOrWhiteSpace(destination))
            {
                System.Windows.MessageBox.Show("Please select a valid output path.");
                return null;
            }

            return new Application();
        }

        private static CsvReader ConfigureCSV(string source)
        {
            var sr = new StreamReader(source);
            var csv = new CsvReader(sr);

            csv.Configuration.MissingFieldFound = null;
            csv.Configuration.RegisterClassMap<RealmsRecordMap>();

            return csv;
        }

        private static void Cleanup (Application application, Workbooks workbooks, Workbook workbook, Worksheet worksheet)
        {
            workbook.Close();
            workbooks.Close();
            application.Quit();

            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(workbooks);
            Marshal.FinalReleaseComObject(application);

            worksheet = null;
            workbook = null;
            workbooks = null;
            application = null;

            GC.Collect();
        }


        private static Department ReadRealmsRecords(CsvReader csv)
        {
            Campus currentCampus = null;
            Department currentDepartment = new Department ("root", null);

            // Read CSV headers
            csv.Read();
            csv.ReadHeader();

            // Read each record
            while (csv.Read())
            {
                // Populate record using RealmsAccountRecordMap
                var record = csv.GetRecord<RealmsRecord>();

                if (record.IsCampusHeading())
                {
                    currentCampus = new Campus()
                    {
                        ID = currentCampus == null ? 0 : currentCampus.ID + 1,
                        Name = record.Account.Remove(0, @"Fund: ".Length)
                    };
                }
                else if (record.IsAccountRecord())
                {
                    // Add Account to currentDepartment
                    var account = currentDepartment.GetOrCreateAccount(record, currentDepartment);
                    account.SetActual(record.Actual, currentCampus.ID);
                    account.SetBudget(record.Budget, currentCampus.ID);
                    account.SetVariance(record.Variance, currentCampus.ID);
                }
                else if (record.IsDepartmentHeading())
                {
                    // Set currentDepartment to new Department
                    var department = currentDepartment.GetOrCreateDepartment(record.Account, currentDepartment);
                    //var department = new Department(record.Account, currentDepartment);
                    currentDepartment = department;
                }
                else if (record.IsDepartmentTotalRow(currentDepartment))
                {
                    // Set currentDepartment to parent
                    if (currentDepartment.ParentDepartment != null)
                    {
                        currentDepartment = currentDepartment.ParentDepartment;
                    }
                }
            }

            while (currentDepartment.ParentDepartment != null)
            {
                currentDepartment = currentDepartment.ParentDepartment;
            }
            return currentDepartment;
        }

        private static void CreateWorksheetTemplate(Worksheet worksheet)
        {
            worksheet.Cells[1, 1] = "THE CHURCH AT";
            worksheet.Cells[2, 1] = "BUDGET VS ACTUALS";
            worksheet.Cells[3, 1] = "FY 2019";

            worksheet.Cells[6, 1] = "LEAD/OWNER";
            worksheet.Cells[6, 2] = "DEPT";
            worksheet.Cells[6, 3] = "ACCOUNT #";

            worksheet.Cells[6,  4] = "Battlecreek";
            worksheet.Cells[6,  5] = "Downtown";
            worksheet.Cells[6,  6] = "Jenks";
            worksheet.Cells[6,  7] = "Midtown";
            worksheet.Cells[6,  8] = "Mission House";
            worksheet.Cells[6,  9] = "Owasso";
            worksheet.Cells[6, 10] = "South Tulsa";
            worksheet.Cells[6, 11] = "All Campuses";

            worksheet.Range["A1:A6"].EntireRow.Font.Bold = true;

            worksheet.Range["D7:D7"].Select();
            worksheet.Application.ActiveWindow.FreezePanes = true;
        }

        private static void PopulateTemplate (Worksheet worksheet, Department departmentRoot)
        {
            int row = 8;

            // Populate Actuals
            PopulateTemplateSection(worksheet, departmentRoot, account => account.Actuals, "ACTUALS", ref row);

            ++row;

            PopulateTemplateSection(worksheet, departmentRoot, account => account.Budgets, "BUDGETS", ref row);

            ++row;

            PopulateTemplateSection(worksheet, departmentRoot, account => account.Variances, "VARIANCES", ref row);

            worksheet.Columns.AutoFit();
        }

        private static void PopulateTemplateSection(Worksheet worksheet, Department departmentRoot, Func<Account, List<double>> method, string sectionName, ref int row)
        {
            worksheet.Cells[row, 1] = sectionName;
            ((Range)worksheet.Cells[row, 1]).Font.Bold = true;
            ((Range)worksheet.Cells[row, 1]).Font.Underline = true;
            ++row;
            departmentRoot.PrintExcelRows(worksheet, ref row, account => account.Actuals, sectionName);
        }
    }
}
