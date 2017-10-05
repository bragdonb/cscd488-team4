using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace WaterAnalysisTool
{
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo newFile = new FileInfo(@"sample1.xlsx");
            if(newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(@"sample1.xlsx");
            }

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                
            }
        }
    }
}
