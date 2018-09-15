using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using CsvHelper;
using RCSVB.Models;

namespace RCSVB
{
    class ExcelBuilder
    {
        public static int CreateFromRealmsCSV(string source, string destination)
        {
            var app = CreateExcelApp(destination);

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
            Department currentDepartment = null;

            // Read CSV headers
            csv.Read();
            csv.ReadHeader();

            // Read each record
            while (csv.Read())
            {
                // Populate record using RealmsAccountRecordMap
                var record = csv.GetRecord<RealmsRecord>();

                if (record.IsAccountRecord())
                {
                    // Add Account to currentDepartment
                    var account = new Account(record, currentDepartment);
                }
                else if (record.IsDepartmentHeading())
                {
                    // Set currentDepartment to new Department
                    var department = new Department(record.Account, currentDepartment);
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

            worksheet.Cells[6,  4] = "CC";
            worksheet.Cells[6,  5] = "BC";
            worksheet.Cells[6,  6] = "DT";
            worksheet.Cells[6,  7] = "JK";
            worksheet.Cells[6,  8] = "MT";
            worksheet.Cells[6,  9] = "OW";
            worksheet.Cells[6, 10] = "ST";
            worksheet.Cells[6, 11] = "TOTAL TC";

            worksheet.Range["A1:A6"].EntireRow.Font.Bold = true;

            worksheet.Range["D7:D7"].Select();
            worksheet.Application.ActiveWindow.FreezePanes = true;

            worksheet.Columns[11].NumberFormat = "$#,###.00";
        }

        private static void PopulateTemplate (Worksheet worksheet, Department departmentRoot)
        {
            int row = 8;

            // Populate Actuals
            worksheet.Cells[row, 1] = "ACTUALS";
            ((Range)worksheet.Cells[row, 1]).Font.Bold = true;
            ((Range)worksheet.Cells[row, 1]).Font.Underline = true;
            ++row;
            departmentRoot.PrintExcelRows (worksheet, ref row, account => account.Actual, "ACTUALS");

            ++row;

            worksheet.Cells[row, 1] = "BUDGET";
            ((Range)worksheet.Cells[row, 1]).Font.Bold = true;
            ((Range)worksheet.Cells[row, 1]).Font.Underline = true;
            ++row;
            departmentRoot.PrintExcelRows (worksheet, ref row, account => account.Budget, "BUDGET");

            ++row;

            worksheet.Cells[row, 1] = "VARIANCE";
            ((Range)worksheet.Cells[row, 1]).Font.Bold = true;
            ((Range)worksheet.Cells[row, 1]).Font.Underline = true;
            ++row;
            departmentRoot.PrintExcelRows (worksheet, ref row, account => account.Variance, "VARIANCE");

            // Autofit
            worksheet.Columns.AutoFit();
        }
    }
}
