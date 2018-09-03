using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;

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

            var excelApp = new Application();

            if (excelApp == null)
            {
                System.Windows.MessageBox.Show ("Excel is not installed on your system.");
                return;
            }

            var excelWorkbook = excelApp.Workbooks.Add();
            var excelWorksheet = (Worksheet)excelWorkbook.Worksheets.Item[1];

            // Get records from csv
            var records = AccountRecordsFromCSV(source);

            // Create Excel Template

            // Populate template with records

            excelWorkbook.SaveAs (destination);

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
    }
}
