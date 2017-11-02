using OfficeOpenXml; // This is the top namespace for EPPlus, if your reference isn't found use the command -> Update-Package -reinstall in the NuGet Console
using System.IO;
using WaterAnalysisTool.Loader;
using WaterAnalysisTool.Components;

namespace WaterAnalysisTool
{
    class Program
    {
        static void Main(string[] args)
        {
        
        }

        /* EPPlus Example. Find documentation at: http://www.nudoq.org/#!/Packages/EPPlus/EPPlus/OfficeOpenXml
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
        */
        /* DataLoader Load() Tester
        static void Main(string[] args)
        {
            FileInfo fi = new FileInfo(@"tester.xlsx");
            if (fi.Exists)
                fi.Delete();

            using (var p = new ExcelPackage(new FileInfo(@"tester.xlsx")))
            {
                p.Workbook.Properties.Title = "Title of Workbook";
                var ws = p.Workbook.Worksheets.Add("Data");
                p.Workbook.Worksheets.Add("Calibration Standards");


                DataLoader loader = new DataLoader(null, p);
                Sample generic = new Sample("Method Name", "Generic Sample", "10/28/20117 12:00:00", "Sample Type", 3);
                generic.AddElement(new Element("Element 1", "Units", 1.0, 0.0, 1.0));
                generic.AddElement(new Element("Element 2", "Units", 1.0, 0.0, 1.0));
                generic.AddElement(new Element("Element 3", "Units", 1.0, 0.0, 1.0));
                generic.AddElement(new Element("Element 4", "Units", 1.0, 0.0, 1.0));
                generic.AddElement(new Element("Element 5", "Units", 1.0, 0.0, 1.0));
                generic.AddElement(new Element("Element 6", "Units", 1.0, 0.0, 1.0));
                generic.AddElement(new Element("Element 7", "Units", 1.0, 0.0, 1.0));
                generic.AddElement(new Element("Element 8", "Units", 1.0, 0.0, 1.0));
                generic.AddElement(new Element("Element 9", "Units", 1.0, 0.0, 1.0));
                generic.AddElement(new Element("Element 10", "Units", 1.0, 0.0, 1.0));
                loader.AddSample(generic);

                Sample calib1 = new Sample("Method Name", "Instrument Blank", "10/28/20117 12:00:00", "Sample Type", 3);
                calib1.AddElement(new Element("Element 1", "Units", 1.0, 1.0, 1.0));
                calib1.AddElement(new Element("Element 2", "Units", 2.0, 2.0, 1.0));
                calib1.AddElement(new Element("Element 3", "Units", 3.0, 3.0, 1.0));
                calib1.AddElement(new Element("Element 4", "Units", 4.0, 4.0, 1.0));
                calib1.AddElement(new Element("Element 5", "Units", 5.0, 5.0, 1.0));
                calib1.AddElement(new Element("Element 6", "Units", 6.0, 6.0, 1.0));
                calib1.AddElement(new Element("Element 7", "Units", 7.0, 7.0, 1.0));
                calib1.AddElement(new Element("Element 8", "Units", 8.0, 8.0, 1.0));
                calib1.AddElement(new Element("Element 9", "Units", 9.0, 9.0, 1.0));
                calib1.AddElement(new Element("Element 10", "Units", 10.0, 10.0, 1.0));
                loader.AddCalibrationSample(calib1);

                Sample calib2 = new Sample("Method Name", "Instrument Blank", "10/28/20117 12:00:00", "Sample Type", 3);
                calib2.AddElement(new Element("Element 1", "Units", 1.0, 1.0, 1.0));
                calib2.AddElement(new Element("Element 2", "Units", 2.0, 2.0, 1.0));
                calib2.AddElement(new Element("Element 3", "Units", 3.0, 3.0, 1.0));
                calib2.AddElement(new Element("Element 4", "Units", 4.0, 4.0, 1.0));
                calib2.AddElement(new Element("Element 5", "Units", 5.0, 5.0, 1.0));
                calib2.AddElement(new Element("Element 6", "Units", 6.0, 6.0, 1.0));
                calib2.AddElement(new Element("Element 7", "Units", 7.0, 7.0, 1.0));
                calib2.AddElement(new Element("Element 8", "Units", 8.0, 8.0, 1.0));
                calib2.AddElement(new Element("Element 9", "Units", 9.0, 9.0, 1.0));
                calib2.AddElement(new Element("Element 10", "Units", 10.0, 10.0, 1.0));
                loader.AddCalibrationSample(calib2);

                loader.Load(); // Load calls Parse, don't need to in main

                //System.Console.WriteLine(ws.Cells[1, 5].Address);
                //System.Console.ReadLine();

                p.Save();
            }
        }
        */
    }
}
