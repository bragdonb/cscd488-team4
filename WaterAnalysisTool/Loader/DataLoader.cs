using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using WaterAnalysisTool.Components;

namespace WaterAnalysisTool.Loader
{
    class DataLoader
    {
        /* Attributes */
        private List<Sample> CalibrationSamples;        // Quality Control Solutions (Insturment Blanks) -> Sample Type: QC
        private List<Sample> CalibrationsStandards;     // Calibration Standard -> Sample Type: Cal
        private List<Sample> QualityControlSamples;     // Stated Values (CCV) -> Sample Type: QC
        private List<Sample> CertifiedValueSamples;     // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC
        private List<Sample> Samples;

        private StreamReader Input;
        private ExcelPackage Output;

        /* Constructors */
        public DataLoader(StreamReader inf, ExcelPackage outf)
        {
            this.Input = inf;
            this.Output = outf;

            this.CalibrationSamples = new List<Sample>();
            this.QualityControlSamples = new List<Sample>();
            this.CertifiedValueSamples = new List<Sample>();
            this.Samples = new List<Sample>();
        }

        /* Public Methods */
        public void Load()
        {
            // TODO
            // Load performs the following functions:
            // 1. Write QC Sample, Certified Val Sample, and Sample data into the Data worksheet
            //  1.1 Method Header, Analysis Date, and Sample Name Descriptor in as first three rows
            //  1.2 Bolded Element Name headers (x2 one for mg/L and another for RSD)
            //  1.3 Bolded Units (x2 one for mg/L and another for RSD)
            //  1.4 Write QC 
            // 2. Write Calibration Sample data into the Calibration Standards Worksheet
            // Load expects the package to have all required worksheets

            // TODO error checking
            #region Error Checking
            if (this.Output.Workbook == null)
                throw new ArgumentNullException("Workbook is null.\n");

            if (this.Output.Workbook.Worksheets.Count < 2)
                throw new ArgumentOutOfRangeException("Invalid number of worksheets present in workbook.\n");
            #endregion

            //DataLoaderParser parser = new DataLoaderParser(this, Input);
            //parser.Parse();

            var dataws = Output.Workbook.Worksheets[1]; // The Data worksheet should be the first worksheet, indeces start at 1.

            // Write header info
            Sample headerSample = Samples[Samples.Count - 1];
            dataws.Cells["A1"].Value = headerSample.Method;
            dataws.Cells["A2"].Value = headerSample.RunTime.Split(' ')[0];
            dataws.Cells["A2"].Value = Output.Workbook.Properties.Title; // Assumes this was set to like the filename, change later to accept user input for title?

            // Write element header rows
            int col = 3; // Start at: row 5, column C
            foreach (Element e in headerSample.Elements)
            {
                // Concentration headers
                dataws.Cells[5, col].Value = e.Name;
                dataws.Cells[5, col].Style.Font.Bold = true;

                dataws.Cells[6, col].Value = e.Units;
                dataws.Cells[6, col].Style.Font.Bold = true;

                // RSD headers
                dataws.Cells[5, col + headerSample.Elements.Count + 2].Value = e.Name;
                dataws.Cells[5, col + headerSample.Elements.Count + 2].Style.Font.Bold = true;

                dataws.Cells[6, col + headerSample.Elements.Count + 2].Value = "RSD";
                dataws.Cells[6, col + headerSample.Elements.Count + 2].Style.Font.Bold = true;

                col++;
            }

            // Write QC/QA samples
            int row = 7; // Start at row 7, col 1
            int count = 0;
            int rowStart, rowEnd;
            if(CalibrationSamples.Count > 0)
            {
                dataws.Cells[row, 1].Value = "Quality Control Solutions";
                dataws.Cells[row, 1].Style.Font.Bold = true;
                row++;
                rowStart = row;

                foreach(Sample s in CalibrationSamples)
                {
                    col = 1;
                    count = 0;

                    dataws.Cells[row, col].Value = s.Name;
                    dataws.Cells[row, ++col].Value = s.RunTime.Split(' ')[1];

                    foreach(Element e in s.Elements)
                    {
                        count++;

                        // Write Analyte concentrations
                        dataws.Cells[row, col + 1].Value = e.Average;
                        // TODO apply QA/QC formatting

                        // Write RSD
                        dataws.Cells[row, col + 1 + s.Elements.Count + 2].Value = e.RSD;
                        // TODO apply QA/AC formatting

                        col++;
                    }

                    row++;
                }

                rowEnd = row - 1;

                dataws.Cells[row, 1].Value = "average";
                dataws.Cells[row, 1].Style.Font.Bold = true;

                for (col = 3; col <= count + 2; col++)
                    dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";

                row++;
                dataws.Cells[row, 1].Value = "LOD";
                dataws.Cells[row, 1].Style.Font.Bold = true;

                for(col = 3; col <= count + 2; col++)
                    dataws.Cells[row, col].Formula = "3*STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";

                row++;
                dataws.Cells[row, 1].Value = "LOQ";
                dataws.Cells[row, 1].Style.Font.Bold = true;

                for (col = 3; col <= count + 2; col++)
                    dataws.Cells[row, col].Formula = "10*STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";

                row+=2;
            }
        } // end Load

        // TODO
        // Error checking for Add<Sample> functions
        public void AddCalibrationSample(Sample sample)
        {
            this.CalibrationSamples.Add(sample);
        }

        public void AddQualityControlSample(Sample sample)
        {
            this.QualityControlSamples.Add(sample);
        }

        public void AddCertifiedValueSample(Sample sample)
        {
            this.CertifiedValueSamples.Add(sample);
        }

        public void AddSample(Sample sample)
        {
            this.Samples.Add(sample);
        }
    }
}
