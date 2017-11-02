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
        // These need to change to SampleGroup classes
        // Certified values will need to be a list of SampleGroups because there can be different sample names in that group 
        private List<Sample> CalibrationSamples;        // Quality Control Solutions (Insturment Blanks) -> Sample Type: QC
        private List<Sample> CalibrationsStandards;     // Calibration Standard -> Sample Type: Cal
        private List<Sample> QualityControlSamples;     // Stated Values (CCV) -> Sample Type: QC
        private List<Sample> CertifiedValueSamples;     // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC
        private List<Sample> Samples;                   // Generic Samples -> Sample Type: Unk

        private StreamReader Input;
        private ExcelPackage Output;

        /* Constructors */
        public DataLoader(StreamReader inf, ExcelPackage outf)
        {
            this.Input = inf;
            this.Output = outf;

            this.CalibrationSamples = new List<Sample>();
            this.CalibrationsStandards = new List<Sample>();
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

            DataLoaderParser parser = new DataLoaderParser(this, Input);
            parser.Parse();

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

            // Write samples
            int row = 7; // Start at row 7, col 1

            if(CalibrationSamples.Count > 0)
                row = WriteSamples(dataws, CalibrationSamples, nameof(CalibrationSamples), row);

            if(QualityControlSamples.Count > 0)
                row = WriteSamples(dataws, QualityControlSamples, nameof(QualityControlSamples), row);

            if(CertifiedValueSamples.Count > 0)
                row = WriteSamples(dataws, CertifiedValueSamples, nameof(CertifiedValueSamples), row);

            if(Samples.Count > 0)
                row = WriteSamples(dataws, Samples, nameof(Samples), row);

            // Write calibration standards
            var calibws = Output.Workbook.Worksheets[2]; // The calibration worksheet is the second worksheet
            WriteStandards(calibws, CalibrationSamples);
        } // end Load

        #region Add<Sample>
        public void AddCalibrationSample(Sample sample)
        {
            if (sample == null)
                throw new ArgumentNullException("Sample is null.\n");

            this.CalibrationSamples.Add(sample);
        }

        public void AddQualityControlSample(Sample sample)
        {
            if (sample == null)
                throw new ArgumentNullException("Sample is null.\n");

            this.QualityControlSamples.Add(sample);
        }

        public void AddCertifiedValueSample(Sample sample)
        {
            if (sample == null)
                throw new ArgumentNullException("Sample is null.\n");

            this.CertifiedValueSamples.Add(sample);
        }

        public void AddSample(Sample sample)
        {
            if (sample == null)
                throw new ArgumentNullException("Sample is null.\n");

            this.Samples.Add(sample);
        }
        #endregion

        /* Private Methods */
        private int WriteSamples(ExcelWorksheet dataws, List<Sample> samples, String type, int row)
        {
            int count = 0;
            int rowStart, rowEnd, col;

            switch (type)
            {
                case "CalibrationSamples":
                    dataws.Cells[row, 1].Value = "Quality Control Solutions";
                    break;

                case "QualityControlSamples":
                    dataws.Cells[row, 1].Value = "Stated Values";
                    // TODO figure out what the heck the numbers in this row come from and write them
                    break;

                case "CertifiedValuesSamples":
                    //calculate average/rsd(%)/recovery(%)
                    break;

                default:
                    break;
            }

            dataws.Cells[row, 1].Style.Font.Bold = true;

            row++;
            rowStart = row;

            foreach (Sample s in samples)
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

                    // TODO apply QA/QC formatting if applicable to sample type
                    switch (type)
                    {
                        case "CalibrationSamples":
                            //calculate average/LOD/LOQ

                            break;

                        case "QualityControlSamples":
                            //calculate average and %difference
                            break;

                        case "CertifiedValuesSamples":
                            //calculate average/rsd(%)/recovery(%)
                            break;

                        default:
                            break;
                    }
                    // Write RSD
                    dataws.Cells[row, col + 1 + s.Elements.Count + 2].Value = e.RSD;

                    // TODO apply QA/AC formatting if applicable to sample type

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

                case "CertifiedValueSamples": // TODO figure out how the heck to split up certified values
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

        private void WriteStandards(ExcelWorksheet calibws, List<Sample> standards)
        {
            // Write element header rows
            Sample headerSample = standards[standards.Count - 1];

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
            foreach (Sample s in standards)
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
