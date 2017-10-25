using OfficeOpenXml; // This is the top namespace for EPPlus, if your reference isn't found use the command -> Update-Package -reinstall in the Nuget Console
using System.IO;

namespace WaterAnalysisTool
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create new workbook, worksheet, and modify a cell
            using (var p = new ExcelPackage()) // Using will dispose of package at exit (it's like close(file) I think)
            {
                // Must add a worksheet to the book
                var ws = p.Workbook.Worksheets.Add("TestWorksheet");

                // Set values in worksheet using Cell indexer
                ws.Cells["A1"].Value = "This is cell A1";

                // SaveAs with specified file name
                p.SaveAs(new FileInfo(@"testworkbook.xlsx")); // It is gonna dump this in bin/Debug
            }

            // Re-open the previously created workbook, if it DNE then create a new one
            FileInfo fi = new FileInfo(@"testworkbook.xlsx");
            using (var p = new ExcelPackage(fi))
            {
                // Get the previously created worksheet
                var ws = p.Workbook.Worksheets["TestWorksheet"];

                // Set values in worksheet using row and col
                ws.Cells[2, 1].Value = "This is Cell B1. Its style is set to bold.";

                // Use style object to set formatting and styles
                ws.Cells[2, 1].Style.Font.Bold = true;

                p.Save();
            }
        }
    }
}
