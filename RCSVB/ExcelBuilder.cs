using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using RCSVB.Models;

namespace RCSVB
{
    class ExcelBuilder
    {
        public static void CreateFromRealmsCSV(string source, string destination)
        {
            var app = CreateExcelApp(destination);

            Workbooks workbooks = app.Workbooks;
            Workbook workbook = workbooks.Add();
            Worksheet worksheet = (Worksheet)workbook.Worksheets.Item[1];

            var csv = ConfigureCSV(source);

            // Reading records the old way
            var records = AccountRecordsFromCSV(csv);

            CreateWorksheetTemplate(worksheet);

            // Writing records the old way
            PopulateTemplate(worksheet, records);

            workbook.SaveAs (destination);

            // Cleanup
            Cleanup(app, workbooks, workbook, worksheet);
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
            csv.Configuration.RegisterClassMap<RealmsAccountRecordMap>();

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


        private static Department RealmsRecords(CsvReader csv)
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

                if (record.IsAccountRecord)
                {
                    // Add Account to currentDepartment
                    var account = new Account(currentDepartment)
                    {
                        Name = record.Account,
                        Actual = float.Parse(record.Actual),
                        Budget = float.Parse(record.Budget),
                        Variance = float.Parse(record.Budget)
                    };
                    currentDepartment.Accounts.Add(account);
                }
                else if (record.IsDepartmentHeading)
                {
                    // Set currentDepartment to new Department
                    var department = new Department(record.Account, currentDepartment);
                    currentDepartment = department;
                }
                else if (record.IsDepartmentTotalRow)
                {
                    // Set currentDepartment to parent
                    if (currentDepartment.ParentDepartment != null)
                    {
                        currentDepartment = currentDepartment.ParentDepartment;
                    }
                }
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

        }

        private static void PopulateTemplate (Worksheet worksheet, List<RealmsAccountRecord> records) {

            int row = 8;
            // Populate Actuals
            worksheet.Cells[row, 1] = "ACTUALS";
            ((Range)worksheet.Cells[row, 1]).Font.Bold = true;
            ((Range)worksheet.Cells[row, 1]).Font.Underline = true;

            int actualsStartRow = ++row;
            int ownerStartRow = actualsStartRow;

            for (int i = 0; i < records.Count; ++i, ++row)
            {
                worksheet.Cells[row, 1] = records[i].Owner;
                worksheet.Cells[row, 2] = records[i].Department;
                worksheet.Cells[row, 3] = records[i].Account;
                worksheet.Cells[row, 11] = records[i].Actual;

                if (i < records.Count - 1 && records[i].Owner != records[i + 1].Owner)
                {
                    ++row;
                    worksheet.Cells[row, 1] = records[i].Owner + " Total";
                    worksheet.Cells[row, 11] = string.Format("=SUM(K{0}:K{1})", ownerStartRow, row - 1);
                    ((Range)worksheet.Cells[row, 11]).NumberFormat = "$ #,###.00";
                    ((Range)worksheet.Rows[row]).Font.Bold = true;
                    GroupRows(worksheet, ownerStartRow, row - 1);
                    ownerStartRow = row + 1;
                }
                else if (i == records.Count)
                {
                    ++row;
                    worksheet.Cells[row, 1] = records[i].Owner + " Total";
                    worksheet.Cells[row, 11] = string.Format("=SUM(K{0}:K{1})", ownerStartRow, row - 1);
                    ((Range)worksheet.Cells[row, 11]).NumberFormat = "$ #,###.00";
                    ((Range)worksheet.Rows[row]).Font.Bold = true;
                    GroupRows(worksheet, ownerStartRow, row - 1);
                }

            }

            int actualsEndRow = row - 1;
            GroupRows(worksheet, actualsStartRow, actualsEndRow);
        

            // Populate Budget

            // Populate Variance

            // Autofit
            worksheet.Columns.AutoFit();

        }

        private static void GroupRows (Worksheet worksheet, int start, int end)
        {
            worksheet.Rows[string.Format("{0}:{1}", start, end)].Group();
        }
    }
}
