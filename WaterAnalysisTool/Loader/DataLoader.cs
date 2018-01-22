using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using WaterAnalysisTool.Components;
using WaterAnalysisTool.Exceptions;
using System.Text.RegularExpressions;

namespace WaterAnalysisTool.Loader
{
    class DataLoader
    {
        #region Attributes
        private SampleGroup CalibrationSamples;             // Quality Control Solutions (Insturment Blanks) -> Sample Type: QC
        private SampleGroup CalibrationStandards;           // Calibration Standard -> Sample Type: Cal
        private SampleGroup QualityControlSamples;          // Stated Values (CCV) -> Sample Type: QC
        private List<SampleGroup> CertifiedValueSamples;    // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC
        private List<SampleGroup> Samples;                  // Generic Samples -> Sample Type: Unk

        private List<String> Messages;
        private StreamReader Input;
        private ExcelPackage Output;
        #endregion

        #region Constructors
        public DataLoader(StreamReader inf, ExcelPackage outf)
        {
            this.Input = inf;
            this.Output = outf;
            this.Output.Workbook.Worksheets.Add("Data");
            this.Output.Workbook.Worksheets.Add("Calibration Standards");

            this.CertifiedValueSamples = new List<SampleGroup>();
            this.Samples = new List<SampleGroup>();

            this.Messages = new List<String>();
        }
        #endregion

        #region Public Methods
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

            // TODO checking if there are actually samples and stuff, more error checking
            #region Error Checking
            if (this.Output.Workbook == null)
                throw new ArgumentNullException("Workbook is null.\n");

            if (this.Output.Workbook.Worksheets.Count < 2)
                throw new ArgumentOutOfRangeException("Invalid number of worksheets present in workbook.\n");
            #endregion

            DataLoaderParser parser = new DataLoaderParser(this, Input);
            //parser.Parse(); // TODO uncomment me when the parser is done

            var dataws = this.Output.Workbook.Worksheets[1]; // The Data worksheet should be the first worksheet, indeces start at 1.

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

            // Freeze top 6 rows and left 2 columns
            dataws.View.FreezePanes(7, 3); // row, col: represents the first row/col that is not frozen

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

            dataws.Cells[row, 1].Value = "Samples";
            dataws.Cells[row, 1].Style.Font.Bold = true;
            row++;
            foreach (SampleGroup g in Samples)
            {
                if (Samples.Count > 0)
                {
                    row = WriteSamples(dataws, g, nameof(Samples), row);
                    row--;
                }
            }

            this.Messages.Add("Samples written to excel sheet successfully.");

            // Write calibration standards
            var calibws = this.Output.Workbook.Worksheets[2]; // The calibration worksheet is the second worksheet
            WriteStandards(calibws, CalibrationStandards);

            this.Output.Save();

            this.Messages.Add("Formatted Excel sheet generated successfullly.");

            foreach (String msg in this.Messages)
                Console.WriteLine("\t" + msg);
        } // end Load

        #region Add<Sample>
        public void AddCalibrationSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Calibration Sample) is null.\n");

            this.CalibrationSamples = (SampleGroup) sample.Clone();
        }

        public void AddCalibrationStandard(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Calibration Standard) is null.\n");

            this.CalibrationStandards = (SampleGroup) sample.Clone();
        }

        public void AddQualityControlSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Quality Control) is null.\n");

            this.QualityControlSamples = (SampleGroup) sample.Clone();
        }

        public void AddCertifiedValueSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Certified Value) is null.\n");

            this.CertifiedValueSamples.Add((SampleGroup) sample.Clone());
        }

        public void AddSampleGroup(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("SampleGroup (Generic) is null.\n");

            this.Samples.Add((SampleGroup) sample.Clone());
        }
        #endregion
        #endregion

        #region Private Methods
        private int WriteSamples(ExcelWorksheet dataws, SampleGroup samples, String type, int row)
        {
            int count = 0;
            int rowStart, rowEnd, col;
            bool flag = false;
            Sample known;

            // Write sample name header
            switch (type)
            {
                case "CalibrationSamples":
                    dataws.Cells[row, 1].Value = "Quality Control Solutions";
                    break;

                case "QualityControlSamples":
                    dataws.Cells[row, 1].Value = "Stated Values";

                    known = samples.Samples[0];
                    col = 3;

                    foreach(Element e in known.Elements)
                    {
                        if (e.Average != -1) // assumes parser set average in elements with no data to -1
                        {
                            dataws.Cells[row, col].Value = e.Average;
                            dataws.Cells[row, col].Style.Font.Bold = true;
                        }

                        col++;
                    }

                    break;

                case "CertifiedValueSamples":
                    dataws.Cells[row, 1].Value = "Certified Values";

                    known = samples.Samples[0];
                    col = 3;

                    foreach (Element e in known.Elements)
                    {
                        if (e.Average != -1) // assumes parse set average in elements with no data to -1
                        {
                            dataws.Cells[row, col].Value = e.Average;
                            dataws.Cells[row, col].Style.Font.Bold = true;
                        }

                        col++;
                    }

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

                if (type == "QualityControlSamples" || type == "CertifiedValueSamples") // skip the first sample in these types because the first sample is known values and already written; seems like this could be cleaned up
                {
                    if(s != samples.Samples[0])
                    {
                        dataws.Cells[row, col].Value = s.Name;
                        dataws.Cells[row, ++col].Value = s.RunTime.Split(' ')[1];

                        foreach (Element e in s.Elements)
                        {
                            count++;

                            if (e.Average != -1) // won't bother with cells where data does not exist (assumes parser set average in elements with no data to -1)
                            {
                                // Write Analyte concentrations
                                dataws.Cells[row, col + 1].Value = e.Average;

                                // Write RSD
                                dataws.Cells[row, col + 1 + s.Elements.Count + 2].Value = e.RSD;

                                // Do QA/QC formatting to analyte concentrations
                                #region QA/AC Formatting
                                if (type == "Samples")
                                {

                                    // REQ-S3R7, lowest in heirarchy
                                    dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Green);

                                    // REQ-S3R2, 1st in heirarchy
                                    if (e.Average > this.CalibrationSamples.LOD[count])
                                    {
                                        dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Firebrick);
                                        flag = true;
                                    }

                                    // REQ-S3R3, 2nd in heirarchy
                                    else if (e.Average < this.CalibrationSamples.LOQ[count] && e.Average > this.CalibrationSamples.LOD[count])
                                    {
                                        dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Orange);
                                        flag = true;
                                    }

                                    // REQ-S3R4, 3rd in heirarchy
                                    else if (!flag)
                                    {
                                        foreach (SampleGroup g in this.CertifiedValueSamples)
                                            if (g.Average[count] < e.Average + 0.5 && g.Average[count] > e.Average - 0.5)
                                                if (g.Recovery[count] > 110 || g.Recovery[count] < 90)
                                                    dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.DodgerBlue);
                                    }

                                    // REQ-S3R5, 4th in heirarchy
                                    else if (this.CalibrationSamples.Average[count] > 0.05 * e.Average)
                                    {
                                        dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Black);
                                        dataws.Cells[row, col + 1].Style.Fill.BackgroundColor.SetColor(Color.Firebrick);
                                        flag = true;
                                    }

                                    // REQ-S3R6, 5th in heirarchy
                                    else if (!flag)
                                    {
                                        Double highest = 0.0;

                                        foreach (Sample std in this.CalibrationStandards.Samples)
                                        {
                                            if (std.Elements[count].Average > highest)
                                                highest = std.Elements[count].Average;
                                        }

                                        if (e.Average > highest)
                                            dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.BlueViolet);
                                    }
                                }
                                #endregion
                            }

                            col++;
                        }

                        row++;
                    }
                }

                else
                {
                    dataws.Cells[row, col].Value = s.Name;
                    dataws.Cells[row, ++col].Value = s.RunTime.Split(' ')[1];

                    foreach (Element e in s.Elements)
                    {
                        count++;

                        if (e.Average != -1) // won't bother with cells where data does not exist (assumes parser set average in elements with no data to -1)
                        {
                            // Write Analyte concentrations
                            dataws.Cells[row, col + 1].Value = e.Average;

                            // Write RSD
                            dataws.Cells[row, col + 1 + s.Elements.Count + 2].Value = e.RSD;

                            // Do QA/QC formatting to analyte concentrations
                            #region QA/AC Formatting
                            if (type == "Samples")
                            {

                                // REQ-S3R7, lowest in heirarchy
                                dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Green);

                                // REQ-S3R2, 1st in heirarchy
                                if (e.Average > this.CalibrationSamples.LOD[count])
                                {
                                    dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Firebrick);
                                    flag = true;
                                }

                                // REQ-S3R3, 2nd in heirarchy
                                else if (e.Average < this.CalibrationSamples.LOQ[count] && e.Average > this.CalibrationSamples.LOD[count])
                                {
                                    dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Orange);
                                    flag = true;
                                }

                                // REQ-S3R4, 3rd in heirarchy
                                else if (!flag)
                                {
                                    foreach (SampleGroup g in this.CertifiedValueSamples)
                                        if (g.Average[count] < e.Average + 0.5 && g.Average[count] > e.Average - 0.5)
                                            if (g.Recovery[count] > 110 || g.Recovery[count] < 90)
                                                dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.DodgerBlue);
                                }

                                // REQ-S3R5, 4th in heirarchy
                                else if (this.CalibrationSamples.Average[count] > 0.05 * e.Average)
                                {
                                    dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.Black);
                                    dataws.Cells[row, col + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    dataws.Cells[row, col + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Firebrick);
                                    flag = true;
                                }

                                // REQ-S3R6, 5th in heirarchy
                                else if (!flag)
                                {
                                    Double highest = 0.0;

                                    foreach (Sample std in this.CalibrationStandards.Samples)
                                    {
                                        if (std.Elements[count].Average > highest)
                                            highest = std.Elements[count].Average;
                                    }

                                    if (e.Average > highest)
                                        dataws.Cells[row, col + 1].Style.Font.Color.SetColor(Color.BlueViolet);
                                }
                            }
                            #endregion
                        }

                        col++;
                    }

                    row++;
                }
            }

            rowEnd = row - 1;

            #region Write Unique Rows
            // TODO determine if we want to write formulas for elements that weren't measure for (have no data, see examples where there are formula errors)
            switch (type)
            {
                case "CalibrationSamples":
                    dataws.Cells[row, 1].Value = "average";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "LOD";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "3*STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "LOQ";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "10*STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                    }

                    break;

                case "QualityControlSamples":
                    dataws.Cells[row, 1].Value = "average";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "% difference";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "(" + dataws.Cells[rowEnd + 1, col].Address + "-" + dataws.Cells[rowStart - 1, col].Address + ")/" + dataws.Cells[rowStart - 1, col].Address + "*100";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                    }

                    break;

                case "CertifiedValueSamples":
                    dataws.Cells[row, 1].Value = "average";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "AVERAGE(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")";
                        dataws.Cells[row, col].Style.Font.Bold = true;
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "rsd (%)";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = "STDEV(" + dataws.Cells[rowStart, col].Address + ":" + dataws.Cells[rowEnd, col].Address + ")/" + dataws.Cells[rowEnd + 1, col].Address + "*100";
                        dataws.Cells[row, col].Style.Font.Bold = true;

                        if (samples.RSD[count - 1] > 10)
                            dataws.Cells[row, col].Style.Font.Color.SetColor(Color.Firebrick);
                    }

                    row++;
                    dataws.Cells[row, 1].Value = "recovery (%)";
                    dataws.Cells[row, 1].Style.Font.Bold = true;

                    for (col = 3; col <= count + 2; col++)
                    {
                        dataws.Cells[row, col].Formula = dataws.Cells[rowEnd + 1, col].Address + "/" + dataws.Cells[rowStart - 1, col].Address + "*100";
                        dataws.Cells[row, col].Style.Font.Bold = true;

                        if (samples.Recovery[count - 1] < 90 || samples.Recovery[count - 1] > 110)
                            dataws.Cells[row, col].Style.Font.Color.SetColor(Color.Firebrick); 
                    }

                    break;

                default:
                    break;

            }
            #endregion

            this.Messages.Add(type + " written to excel sheet successfully");

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
                col = 1;

                calibws.Cells[row, col].Value = s.Name;
                calibws.Cells[row, ++col].Value = s.RunTime;

                foreach (Element e in s.Elements)
                {
                    calibws.Cells[row, ++col].Value = e.Average;
                }

                row++;
            }

            int endRow = row + 2;

            this.Messages.Add("Calibration standards written to excel sheet successfully");

            // Calibration Curve
            // 1. Open the CheckStandards.xlsx sheet where the stock solution concentrations can be found and read them in
            //  1.1 Have to worry about not every concentration in the standards list (these will have to be 0's in the .xlsx)
            // 2. Create a graph with the measured counts per second in the standards list over their respective stock solution concentration
            try
            {
                FileInfo fi = new FileInfo("CheckStandards.xlsx");
                if (!fi.Exists)
                    throw new FileNotFoundException("The CheckStandards.xlsx config file does not exist or could not be found and a calibration curve could not be generated.");

                using (var p = new ExcelPackage(fi))
                {
                    ExcelWorksheet standardsws = p.Workbook.Worksheets[2]; // TODO this index may change depending on if the CheckStandards.xlxs file changing
                
                    // Find Continuing Calibration Verification (CCV) seciton
                    row = 1;
                    int blankCount = 0;
                    while(blankCount < 5 && blankCount >= 0)
                    {
                        if(standardsws.Cells[row, 1].Value != null)
                        {
                            if(!standardsws.Cells[row, 1].Value.ToString().Equals("Continuing Calibration Verification (CCV)"))
                            {
                                 row++;
                                 blankCount = 0;
                            }

                            else
                                break;
                        }

                        else
                        {
                            blankCount++;
                            row++;
                        }
                    }

                    if(blankCount > 4)
                        throw new ConfigurationErrorException("Could not find \"Continuing Calibration Verification (CCV)\" section in CheckStandards.xlsx config file.");

                    row++;

                    // Find the row that corresponds to the CCV ratio (QualityControlStandards) !! Don't need this if we do in fact use CCV section in CheckStandards, should check if ratios match !!
                    // String[] QCSName = this.QualityControlSamples.Name.Split(null);
                    // TODO check if QCSName correct length
                    // String ratio = QCSName[1];

                    //while(!standardsws.Cells[row, 1].Value.ToString().Contains(ratio))
                    //{
                        //row++;

                        // Checking if this is an infinite loop
                        //if(standardws.Cells[row, 1].Value == null)
                            //throw new ConfigurationErrorException("Could not find a Check Standard in CheckStandards.xlsx that matches the CCV ratio of " + ratio + ".");
                    //}

                    // Write CCV avg to Calibration Standards worksheet for use as range? !! Don't need this if we don't use CCV section in CheckStandards, instead find on calibws !!
                    col = 2;
                    foreach(double avg in this.QualityControlSamples.Average)
                    {
                        calibws.Cells[endRow, col].Value = avg;
                        col++;
                    }

                    endRow++;

                    // Read in check standards data and write to Calibration Standards worksheet in Output package at endRow
                    col = 2;
                    while(standardsws.Cells[row, col + 1].Value != null)
                    {
                        calibws.Cells[endRow, col].Value = standardsws.Cells[row, col + 1].Value;
                        col++;
                    }

                    // Create the chart
                    ExcelChart calCurve = calibws.Drawings.AddChart("Calibration Curve", eChartType.XYScatter);
                    calCurve.Title.Text = "Calibration Curve";
                    calCurve.SetPosition(endRow + 2, 0, 1, 0);
                    calCurve.SetSize(600, 400);
                    calCurve.YAxis.MinValue = 0;
                    calCurve.XAxis.MinValue = 0;
                    calCurve.Legend.Remove();

                    var yrange = calibws.Cells[endRow - 1, 2, endRow - 1, col];
                    var xrange = calibws.Cells[endRow, 2, endRow, col];

                    var series1 = calCurve.Series.Add(yrange, xrange);
                    series1.TrendLines.Add(eTrendLine.Linear);
                    
                }


                this.Messages.Add("Calibration curve generated successfully");
            }

            catch (Exception e)
            {
                this.Messages.Add("Calibration curve could not be generated. Error: " + e.Message);
                //Console.WriteLine(e.Message);
            }
        }// end WriteStandards
        #endregion
    }   
}
