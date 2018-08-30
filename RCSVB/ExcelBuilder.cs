using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RCSVB
{
    class ExcelBuilder
    {

        public static void CreateWorkbook(string source, string destination)
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

            var excelWorksheet = (Worksheet)excelWorkbook.Worksheets.get_Item(1);

            excelWorkbook.SaveAs(destination);
        }
    }
}
