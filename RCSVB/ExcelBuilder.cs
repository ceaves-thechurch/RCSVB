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
            var app = CreateExcelApp(destination);

            Workbooks workbooks = app.Workbooks;
            Workbook workbook = workbooks.Add();
            Worksheet worksheet = (Worksheet)workbook.Worksheets.Item[1];

            // Open CSV for reading
            var sr = new StreamReader(source);
            var csv = new CsvReader(sr);

            // CSV Reader configuration
            csv.Configuration.MissingFieldFound = null;
            csv.Configuration.RegisterClassMap<RealmsAccountRecordMap>();

            // Get records from csv
            var records = AccountRecordsFromCSV(csv);

            // Create Excel Template
            CreateWorksheetTemplate(worksheet);

            // Populate template with records
            PopulateTemplate(worksheet, records);

            // Save
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

        private static Application CreateExcelApp (string destination)
        {
            if (string.IsNullOrWhiteSpace(destination))
            {
                System.Windows.MessageBox.Show("Please select a valid output path.");
                return null;
            }

            return new Application();
        }

        private static List<RealmsAccountRecord> AccountRecordsFromCSV(CsvReader csv)
        {
            List<RealmsAccountRecord> realmsAccountRecords = new List<RealmsAccountRecord>();

            // Read CSV headers
            csv.Read();
            csv.ReadHeader();

            string owner = "";
            string department = "";

            // Read each record
            while (csv.Read())
            {
                // Populate record using RealmsAccountRecordMap
                var record = csv.GetRecord<RealmsAccountRecord>();
                record.TrimCSVFields();

                // Determine if record is owner, department, or account
                if (record.IsValidAccountRecord)
                {
                    record.Owner = owner;
                    record.Department = department;
                    record.Account = record.Account;

                    realmsAccountRecords.Add(record);

                    continue;
                }

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

            worksheet.Range["D7:D7"].Select();
            worksheet.Application.ActiveWindow.FreezePanes = true;

        }

        private static void PopulateTemplate (Worksheet worksheet, List<RealmsAccountRecord> records) {

            int row = 8;
            // Populate Actuals
            worksheet.Cells[row, 1] = "ACTUALS";
            ((Range)worksheet.Cells[row, 1]).Font.Bold = true;
            ((Range)worksheet.Cells[row, 1]).Font.Underline = true;

            row++;

            foreach (RealmsAccountRecord record in records)
            {
                if (record.Account.Contains ("Total")) {
                    worksheet.Cells[row,  2] = record.Department + " Total";
                    ((Range)worksheet.Cells[row, 2]).Font.Bold = true;
                    worksheet.Cells[row, 11] = record.Actual;
                } 
                else
                {
                    worksheet.Cells[row, 1] = record.Owner;
                    worksheet.Cells[row, 2] = record.Department;
                    worksheet.Cells[row, 3] = record.Account;
                    worksheet.Cells[row, 11] = record.Actual;
                }
                

                ++row;
            }


            // Populate Budget

            // Populate Variance

            // Autofit
            worksheet.Columns.AutoFit();

        }
    }
}
