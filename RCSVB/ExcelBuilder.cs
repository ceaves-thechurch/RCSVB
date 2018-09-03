using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using MissingFieldException = CsvHelper.MissingFieldException;

namespace RCSVB
{
    class ExcelBuilder
    {

        // Create an empty workbook
        public static void CreateEmptyWorkbook(string destination)
        {
            if (string.IsNullOrWhiteSpace (destination))
            {
                System.Windows.MessageBox.Show("Please select a valid output path.");
                return;
            }

            var excelApp = new Application();

            if (excelApp == null)
            {
                System.Windows.MessageBox.Show("Excel is not installed on your system.");
                return;
            }

            var excelWorkbook = excelApp.Workbooks.Add();

            excelWorkbook.SaveAs(destination);
        }

        // Create excel template file from source Realms CSV
        public static void CreateFromRealmsCSV(string source, string destination)
        {

            if (string.IsNullOrWhiteSpace(destination))
            {
                System.Windows.MessageBox.Show ("Please select a valid output path.");
                return;
            }

            var app = new Application();

            if (app == null)
            {
                System.Windows.MessageBox.Show ("Excel is not installed on your system.");
                return;
            }

            Workbooks workbooks = app.Workbooks;
            Workbook workbook = workbooks.Add();
            Worksheet worksheet = (Worksheet)workbook.Worksheets.Item[1];


            // Get records from csv
            var records = AccountRecordsFromCSV(source);

            // Create Excel Template
            CreateWorksheetTemplate(worksheet);

            // Populate template with records

            workbook.SaveAs (destination);

            // Cleanup
            workbook.Close();
            workbooks.Close();
            app.Quit();

            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(workbooks);
            Marshal.FinalReleaseComObject(app);

            worksheet = null;
            workbook = null;
            workbooks = null;
            app = null;

            GC.Collect();
        }

        public static List<RealmsAccountRecord> AccountRecordsFromCSV(string source)
        {
            // Open CSV for reading
            var sr = new StreamReader(source);
            var csv = new CsvReader(sr);

            // CSV Reader configuration
            csv.Configuration.MissingFieldFound = null;
            csv.Configuration.RegisterClassMap<RealmsAccountRecordMap>();

            List<RealmsAccountRecord> realmsAccountRecords = new List<RealmsAccountRecord>();

            // Read CSV headers
            csv.Read ();
            csv.ReadHeader ();

            string currentOwner = "";
            string currentDepartment = "";

            // Read each record
            while (csv.Read ()) {
                // Populate record using RealmsAccountRecordMap
                var record = csv.GetRecord<RealmsAccountRecord>();

                // Check if record is owner or department heading
                if (string.IsNullOrEmpty (record.Actual) ||
                    string.IsNullOrEmpty (record.Budget) ||
                    string.IsNullOrEmpty (record.Variance))
                {
                    // Determine if owner or department
                    if (record.Account.StartsWith("         ")) // 9 spaces
                    {
                        currentDepartment = record.Account.Trim();
                    } else if (record.Account.StartsWith("      ")) // 6 spaces
                    {

                    } else if (record.Account.StartsWith("   ")) // 3 spaces
                    {
                        currentOwner = record.Account.Trim();
                    }

                    continue;
                }

                // Populate additional record properties
                record.Owner = currentOwner;
                record.Department = currentDepartment;
                record.Account = record.Account.Trim();

                realmsAccountRecords.Add(record);
            }

            return realmsAccountRecords;
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
            worksheet.Columns.AutoFit();

            worksheet.Range["D7:D7"].Select();
            worksheet.Application.ActiveWindow.FreezePanes = true;

        }
    }
}
