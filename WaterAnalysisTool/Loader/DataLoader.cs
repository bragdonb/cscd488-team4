using System;
using System.Drawing.Color;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using WaterAnalysisTool.Components;

namespace WaterAnalysisTool.Loader
{
    class DataLoader
    {
        /* Attributes */
        private SampleGroup CalibrationSamples;             // Quality Control Solutions (Insturment Blanks) -> Sample Type: QC
        private SampleGroup CalibrationStandards;           // Calibration Standard -> Sample Type: Cal
        private SampleGroup QualityControlSamples;          // Stated Values (CCV) -> Sample Type: QC
        private List<SampleGroup> CertifiedValueSamples;    // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC
        private List<SampleGroup> Samples;                  // Generic Samples -> Sample Type: Unk

        private StreamReader Input;
        private ExcelPackage Output;

        /* Constructors */
        public DataLoader(StreamReader inf, ExcelPackage outf)
        {
            this.Input = inf;
            this.Output = outf;

            this.CertifiedValueSamples = new List<SampleGroup>();
            this.Samples = new List<SampleGroup>();
        }

        /* Public Methods */
        public void Load()
        {
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

            DataLoaderParser parser = new DataLoaderParser(this, Input);
            parser.Parse();

            var dataws = Output.Workbook.Worksheets[1]; // The Data worksheet should be the first worksheet, indeces start at 1.

            // Write header info
            Sample headerSample = Samples[Samples.Count - 1].Samples[Samples[Samples.Count - 1].Samples.Count - 1]; // good God
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

            // Write samples
            int row = 7; // Start at row 7, col 1

            if(CalibrationSamples.Samples.Count > 0)
                row = WriteSamples(dataws, CalibrationSamples, nameof(CalibrationSamples), row);

            if(QualityControlSamples.Samples.Count > 0)
                row = WriteSamples(dataws, QualityControlSamples, nameof(QualityControlSamples), row);

            foreach (SampleGroup g in CertifiedValueSamples)
            {
                if (g.Samples.Count > 0)
                    row = WriteSamples(dataws, g, nameof(CertifiedValueSamples), row);
            }

            foreach (SampleGroup g in Samples)
            {
                if (Samples.Count > 0)
                    row = WriteSamples(dataws, g, nameof(Samples), row);
            }

            // Write calibration standards
            var calibws = Output.Workbook.Worksheets[2]; // The calibration worksheet is the second worksheet
            WriteStandards(calibws, CalibrationStandards);
        } // end Load

        #region Add<Sample>
        public void AddCalibrationSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Calibration Sample) is null.\n");

            this.CalibrationSamples = sample;
        }

        public void AddCalibrationStandard(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Calibration Standard) is null.\n");

            this.CalibrationStandards = sample;
        }

        public void AddQualityControlSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Quality Control) is null.\n");

            this.QualityControlSamples = sample;
        }

        public void AddCertifiedValueSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Certified Value) is null.\n");

            this.CertifiedValueSamples.Add(sample);
        }

        public void AddSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Generic) is null.\n");

            this.Samples.Add(sample);
        }
        #endregion

        /* Private Methods */
        // TODO this needs to be rewritten to handle SampleGroups as an input; gives access to calculations already performed
        private int WriteSamples(ExcelWorksheet dataws, SampleGroup samples, String type, int row)
        {
            int count = 0;
            int rowStart, rowEnd, col;

            // Write header sample name
            switch (type)
            {
                case "CalibrationSamples":
                    dataws.Cells[row, 1].Value = "Quality Control Solutions";
                    break;

                case "QualityControlSamples":
                    dataws.Cells[row, 1].Value = "Stated Values";
                    // TODO write the first sample in this row (comes from second file)
                    break;

                case "CertifiedValuesSamples":
                    // TODO write the first sample in this row (comes from second file)
                    break;

                default:
                    dataws.Cells[row, 1].Value = samples.Name;
                    break;
            }

            dataws.Cells[row, 1].Style.Font.Bold = true;

            row++;
            rowStart = row;

            // Write sample data
            foreach (Sample s in samples.Samples)
            {
                col = 1;
                count = 0;

                dataws.Cells[row, col].Value = s.Name;
                dataws.Cells[row, ++col].Value = s.RunTime.Split(' ')[1];

                foreach (Element e in s.Elements)
                {
                    count++;

                    // Write Analyte concentrations
                    dataws.Cells[row, col + 1].Value = e.Average;

                    // Do QA/QC formatting to analyt concentrations
                    if(type == "Samples")
                    {
                        // REQ-S3R2
                        if (e.Average > this.CalibrationSamples.LOD[count])
                            dataws.Cells[row, col + 1].Style.Font.Color.SetColor(System.Drawing.Color.Firebrick);

                        // REQ-S3R3
                        else if (e.Average < this.CalibrationSamples.LOQ[count] && e.Average > this.CalibrationSamples.LOD[count])
                            dataws.Cells[row, col + 1].Style.Font.Color.SetColor(System.Drawing.Color.Orange);

                        // REQ-S3R4
                        foreach (SampleGroup g in this.CertifiedValueSamples)
                            if (g.Average[count] < e.Average + 0.5 && g.Average[count] > e.Average - 0.5)
                                if (g.Recovery[count] > 110 || g.Recovery[count] < 90)
                                    dataws.Cells[row, col + 1].Style.Font.Color.SetColor(System.Drawing.Color.DodgerBlue);
                    
                        // REQ-S3R6

                    }

                    else if(type == "CalibrationSamples")
                    {
                        // REQ-S3R5
                    }

                    // Write RSD
                    dataws.Cells[row, col + 1 + s.Elements.Count + 2].Value = e.RSD;

                    col++;
                }

                row++;
            }

            rowEnd = row - 1;

            #region Write Unique Rows
            switch (type)
            {
                case "CalibrationSamples":
                    row++;
                    dataws.Cells[row, 1].Value = "average";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                        dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";

                    row++;
                    dataws.Cells[row, 1].Value = "LOD";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                        dataws.Cells[row, col].Formula = "3*STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";

                    row++;
                    dataws.Cells[row, 1].Value = "LOQ";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                        dataws.Cells[row, col].Formula = "10*STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";

                    break;

                case "QualityControlSamples":
                    row++;
                    dataws.Cells[row, 1].Value = "average";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                        dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";

                    row++;
                    dataws.Cells[row, 1].Value = "% difference";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                        dataws.Cells[row, col].Formula = "(" + dataws.Cells[rowEnd + 1, col].Address + "-" + dataws.Cells[rowStart - 1, col].Address + ")/" + dataws.Cells[rowStart - 1, col].Address + "*100"; // TODO There are extra numbers in the same row as the title "Stated Values"... They are used in this calc

                    break;

                case "CertifiedValueSamples":
                    row++;
                    dataws.Cells[row, 1].Value = "average";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                        dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";

                    row++;
                    dataws.Cells[row, 1].Value = "rsd (%)";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                        dataws.Cells[row, col].Formula = "STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")/" + dataws.Cells[rowEnd + 1, col].Address + "*100";

                    row++;
                    dataws.Cells[row, 1].Value = "recovery (%)";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count; col++)
                        dataws.Cells[row, col].Formula = dataws.Cells[rowEnd + 1, col].Address + "/" + dataws.Cells[rowStart - 1, col].Address + "*100"; // TODO There are extra numbers in the same row as the title "Certified Values"... They are used in this calc

                    break;

                default:
                    break;

            }
            #endregion

            return  row + 2;
        }// end WriteSamples

        private void WriteStandards(ExcelWorksheet calibws, SampleGroup standards)
        {
            // Write element header rows
            Sample headerSample = standards.Samples[standards.Samples.Count - 1];

            int col = 3; //Start at row 2 col 3
            foreach (Element e in headerSample.Elements)
            {
                // Concentration headers
                calibws.Cells[2, col].Value = e.Name;
                calibws.Cells[2, col].Style.Font.Bold = true;

                calibws.Cells[3, col].Value = e.Units;
                calibws.Cells[3, col].Style.Font.Bold = true;

                col++;
            }

            // Write standards data
            int row = 4;
            col = 1;
            foreach (Sample s in standards.Samples)
            {
                calibws.Cells[row, col].Value = s.Name;
                calibws.Cells[row, ++col].Value = s.RunTime;

                foreach (Element e in s.Elements)
                {
                    calibws.Cells[row, ++col].Value = e.Average;
                }
            }

            // Create the calibration curve graph
            // 1. Open the ICP-OESstandards-master list Excel sheet or some config sheet where the stock solution concentrations can be found
            //  1.1 Have to worry about not every concentration in the standards list, what about using the ratio in their name and multiplying by the known mg/L
            // 2. Create a graph with the measured counts per second in the standards list over their respective stock solution concentration

        }// end WriteStandards
    }// end DataLoader class    
}// end WaterAnalysisTool.Loader namespace
