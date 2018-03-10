using System;
using System.IO;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WaterAnalysisTool.Components;
using WaterAnalysisTool.Exceptions;
using OfficeOpenXml;

namespace WaterAnalysisTool.Loader
{
    class DataLoaderParser
    {

        #region Attributes

        private DataLoader Loader;
        private StreamReader Input;

        private List<Sample> CalibrationStandardsList; // Calibration Standard -> Sample Type: Cal, --- These go in the Calibration Standards worksheet.  Calib Blank, CalibStd
        private List<Sample> CalibrationSamplesList;  // Quality Control Solutions -> Sample Type: QC --- These are usually named Instrument Blank
        private List<Sample> QualityControlSamplesList;  // Stated Values (CCV) -> Sample Type: QC --- These will have CCV in the name

        private List<List<Sample>> SampleList; // Samples -> Sample Type: Unk --- These will have very different names (Perry/DFW/etc.)
        private List<List<Sample>> CertifiedValueList;  // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC --- The analytes (elements) found under Check Standards in the xlsx file will not always match up with the analytes of the Certified Value samples

        #endregion

        #region Constructors

        public DataLoaderParser(DataLoader loader, StreamReader inf)
        {
            this.Loader = loader;
            this.Input = inf;

            this.CalibrationStandardsList = new List<Sample>(); // CalibStd
            this.CalibrationSamplesList = new List<Sample>(); // Instrument Blank
            this.QualityControlSamplesList = new List<Sample>(); // CCV

            this.SampleList = new List<List<Sample>>();
            this.CertifiedValueList = new List<List<Sample>>();
        }

        #endregion

        #region Public Methods

        public void Parse()
        {
            this.ParseCheckStandards();

            while (this.Input.Peek() >= 0)
            {
                Sample samp = this.ParseHeader();
                this.ParseResults(samp);
                this.ParseInternalStandards(samp);
                this.AddSampleToList(samp);
            }

            this.PassToDataLoader();
        }

        #endregion

        #region Private Methods

        private SampleGroup CreateSampleGroup(List<Sample> sampleList, string name, bool skipFirst)
        {
            if (sampleList == null || name == null)
                throw new ArgumentException("The SampleGroup you are trying to create will contain a null member variable\n");

            return new SampleGroup(sampleList, name, skipFirst);
        }


        private Sample CreateSample(string method, string name, string comment, string runTime, string sampleType, Int32 repeats)
        {
            if (method == null || name == null || comment == null || runTime == null || sampleType == null || repeats < 0)
                throw new ArgumentNullException("The Sample you are trying to create will contain a null member variable\n");

            return new Sample(method, name, comment, runTime, sampleType, repeats);
        }


        private Element CreateElement(string name, string units, Double avg, Double stddev, Double rsd)
        {
            if (name == null || units == null)
                throw new ArgumentNullException("The Element you are trying to instantiate will contain a null member variable\n");

            return new Element(name, units, avg, stddev, rsd);
        }


        private void AddElementToSample(Sample sample, Element element)
        {
            if (sample == null)
                throw new ArgumentNullException("The sample you are attempting to add an element to is null\n");
            else if (element == null)
                throw new ArgumentNullException("The element you are attempting to add to the sample is null\n");

            sample.Elements.Add(element);
        }


        private void ParseCheckStandards()
        {
            FileInfo fi = new FileInfo("CheckStandards.xlsx");
            if (!fi.Exists)
                throw new FileNotFoundException("The CheckStandards.xlsx config file does not exist or could not be found and a calibration curve could not be generated.");

            using (ExcelPackage ep = new ExcelPackage(fi))
            {
                ExcelWorksheet ews = ep.Workbook.Worksheets[2];
                int row = 3; // Start from row 3. Assuming the format of CheckStandards.xlsx will remain the same
                int col = 3;
                int analyteCounter = 0;
                String sampleName;
                double avg;
                List<string> elementNames = new List<string>();
                Element tmpElem;
                Sample tmpSample;

                while (ews.Cells[row, col].Value != null)
                {
                    elementNames.Add(ews.Cells[row, col].Value.ToString()); // Make a list of element names for reference & creation of List<Element> to add to the sample
                    analyteCounter++;
                    col++;
                }

                row = 1;
                col = 1;
                int blankCounter = 0;

                while (blankCounter < 4)
                {
                    if (row > 14)
                        throw new ConfigurationErrorException("Could not find Continuing Calibration Verification (CCV) section in CheckStandars.xlsx config file.");
                    else if (ews.Cells[row, col].Value == null)
                    {
                        blankCounter++;
                        row++;
                    }
                    else
                        row++;
                }

                if (ews.Cells[row, col].Value.ToString().ToLower().Equals("continuing calibration verification (ccv)")) // parsing CCV section
                {
                    row++;

                    while (ews.Cells[row, col].Value != null)
                    {
                        sampleName = ews.Cells[row, col].Value.ToString(); // Reads in the name of the sample
                        tmpSample = new Sample("", sampleName, "", "", "QC", 0);

                        for (col = 3; col < analyteCounter + 3; col++) // Increment by three. Otherwise we could potentially miss the last two columns of analytes due to the value of the column (col) starting at 3 in the for loop
                        {
                            if (ews.Cells[row, col].Value == null || !(double.TryParse(ews.Cells[row, col].Value.ToString(), out avg)))
                                avg = Double.NaN;

                            tmpElem = new Element(elementNames[col - 3], "mg/L", avg, Double.NaN, Double.NaN); // Stddev and RSD get Double.NaN, assuming all values in CheckStandards are avg. Subtract columns (col) by 3 so that the correct corresponding element name is retrieved
                            tmpSample.AddElement(tmpElem);
                        }

                        this.QualityControlSamplesList.Add(tmpSample);
                        row++;
                        col = 1;
                    }
                }
                else
                    throw new ConfigurationErrorException("Could not find Continuing Calibration Verification (CCV) section in CheckStandars.xlsx config file.");

                if (ews.Cells[row, col].Value == null)
                {
                    blankCounter++;
                    row++;
                }

                if (blankCounter == 5 && ews.Cells[row, col].Value.ToString().ToLower().Equals("check standards")) // parsing check standards section
                {
                    row++;

                    while (ews.Cells[row, col].Value != null)
                    {
                        sampleName = ews.Cells[row, col].Value.ToString();
                        tmpSample = new Sample("", sampleName, "", "", "QC", 0);

                        for (col = 3; col < analyteCounter + 3; col++)
                        {
                            if (ews.Cells[row, col].Value == null || !(double.TryParse(ews.Cells[row, col].Value.ToString(), out avg)))
                                avg = Double.NaN;

                            tmpElem = new Element(elementNames[col - 3], "mg/L", avg, Double.NaN, Double.NaN);
                            tmpSample.AddElement(tmpElem);
                        }

                        List<Sample> certifiedValueTmpList = new List<Sample>();
                        certifiedValueTmpList.Add(tmpSample);
                        this.CertifiedValueList.Add(certifiedValueTmpList);
                        row++;
                        col = 1;
                    }
                }
                else
                    throw new ConfigurationErrorException("Could not find Check Standards section in CheckStandards.xlsx config file.");
            }
        }


        private Sample ParseHeader()
        {
            Sample sample;
            List<string> stringList = new List<string>();
            string tmp = "";

            for (int x = 0; x < 2; x++)
            {
                if (this.Input.Peek() >= 0)
                    tmp = this.Input.ReadLine();
            }

            if (!(tmp.Equals("[Sample Header]")))
                throw new FormatException("The Sample Header section of the input file could not be found.\n");

            if (this.Input.Peek() >= 0) // Start reading in sample meta data
            {
                tmp = this.Input.ReadLine();

                while (!(string.IsNullOrEmpty(tmp)))
                {
                    stringList.Add(tmp);

                    if (this.Input.Peek() >= 0)
                        tmp = this.Input.ReadLine();
                    else
                        throw new FormatException("The Sample Header section of the input file is not formatted correctly.\n");
                }

                if (stringList.Count != 12)
                    throw new FormatException("The Sample Header section of the input file is not formatted correctly.\n");
            }
            else
                throw new FormatException("The Sample Header section of the input file is not formatted correctly.\n");

            stringList[1] = stringList[1].Replace("SampleName=", ""); // SampleName trimming
            stringList[3] = stringList[3].Replace("Comment=", ""); // Comment trimming
            stringList[7] = stringList[7].Substring(stringList[7].IndexOf(' ', 4)); // Getting the correct format for the time
            stringList[8] = stringList[8].Replace("Sample Type=", ""); // SampleType trimming
            stringList[11] = stringList[11].Replace("Repeats=", ""); // Repeats trimming

            sample = CreateSample(stringList[0], stringList[1], stringList[3], stringList[7], stringList[8], int.Parse(stringList[11]));

            return sample;
        }


        private void ParseResults (Sample sample)
        {
            this.CheckForNullSample(sample);
            List<string> stringList = new List<string>();
            string tmp = "";

            if (this.Input.Peek() >= 0) // Consumes line containing "[Results]", line containing labels for elements "Elem,Units,Avg,Stddev,RSD"
            {
                tmp = this.Input.ReadLine();

                if (!(tmp.Equals("[Results]")))
                    throw new FormatException("The Results section of the input file could not be found.\n");
            }

            if (this.Input.Peek() >= 0)
            {
                tmp = this.Input.ReadLine();

                if (!(tmp.Equals("Elem,Units,Avg,Stddev,RSD")))
                    throw new FormatException("The Results section of the input file is not formatted correctly.\n");
            }

            if (this.Input.Peek() >= 0) // Start reading in elements data
            {
                tmp = this.Input.ReadLine();

                while (!(string.IsNullOrEmpty(tmp)))
                {
                    stringList.Add(tmp);

                    if (this.Input.Peek() >= 0)
                        tmp = this.Input.ReadLine();
                    else
                        throw new FormatException("The Results section of the input file is not formatted correctly.\n");
                }
            }
            else
                throw new FormatException("The Results section of the input file could not be found.\n");

            foreach (string str in stringList) // Data scrubbing and Element creation
            {
                string[] strArray = str.Split(',');

                double avg;
                double stddev;
                double rsd;
                string pattern = @"([a-zA-Z]*)(\s+)";

                strArray[2] = Regex.Replace(strArray[2], pattern, ""); // Removes the whitespace and "F" in front of avgs

                if (!(double.TryParse(strArray[2], out avg)))
                    avg = Double.NaN;
                if (!(double.TryParse(strArray[3], out stddev)))
                    stddev = Double.NaN;
                if (!(double.TryParse(strArray[4], out rsd)))
                    rsd = Double.NaN;

                Element newElement = this.CreateElement(strArray[0], strArray[1], avg, stddev, rsd);
                this.AddElementToSample(sample, newElement);
            }
        }


        private void ParseInternalStandards(Sample sample)
        {
            this.CheckForNullSample(sample);

            string str = "";
            string[] strList;

            if (this.Input.Peek() >= 0)
                str = this.Input.ReadLine(); // Consumes line containing "[Internal Standards]"
            else
                throw new FormatException("The Internal Standards section of the input file could not be found.\n");

            if (!(str.Equals("[Internal Standards]")))
                throw new FormatException("The Internal Standards section of the input file is not formatted correctly.\n");

            if (this.Input.Peek() >= 0)
            {
                str = this.Input.ReadLine();
                strList = str.Split(',');
            }
            else
                throw new FormatException("The Internal Standards section of the input file is not formatted correctly.\n");
        }


        private void AddSampleToList(Sample samp)
        {
            this.CheckForNullSample(samp);
            bool certifiedValue = false;
            bool newSample = true;

            if (string.Compare(samp.SampleType, "Cal") == 0)
                this.CalibrationStandardsList.Add(samp); // CalibStd
            else if (samp.Name.StartsWith("CCV"))
                this.QualityControlSamplesList.Add(samp); // CCV
            else if (string.Compare(samp.Name, "Instrument Blank") == 0)
                this.CalibrationSamplesList.Add(samp); // Assuming all Calibration Samples will be named "Instrument Blank"
            else
            {
                for (int x = 0; x < this.CertifiedValueList.Count; x++)
                {
                    if (samp.Name.StartsWith(this.CertifiedValueList[x][0].Name.Substring(0, 3)))
                    {
                        int w = 1;

                        for (; w < this.CertifiedValueList.Count && !(certifiedValue); w++)
                        {
                            if (samp.Name.Equals(this.CertifiedValueList[w][0].Name))
                            {
                                this.CertifiedValueList[w].Add(samp);
                                certifiedValue = true;
                            }
                        }

                        if (w == this.CertifiedValueList.Count && !(certifiedValue))
                        {
                            this.CertifiedValueList[x].Add(samp);
                            certifiedValue = true;
                        }
                    }
                }

                if (!certifiedValue)
                {
                    for (int x = 0; x < this.SampleList.Count; x++)
                    {
                        int subStrLen = Utils.Utils.LongestCommonSubstring(samp.Name, this.SampleList[x][0].Name);

                        if (newSample && subStrLen > 3)
                        {
                            this.SampleList[x].Add(samp);
                            newSample = false;
                        }
                    }
                }

                if (!certifiedValue && newSample)
                    this.CreateNewSampleSubList(samp);
            }
        }


        private void PassToDataLoader()
        {
            foreach (List<Sample> sampleList in this.CertifiedValueList)
                this.Loader.AddCertifiedValueSampleGroup(new SampleGroup(sampleList, "Certified Values", true));

            foreach (List<Sample> sampleList in this.SampleList)
                this.Loader.AddSampleGroup(new SampleGroup(sampleList, sampleList[0].Name.Substring(0, sampleList[0].Name.LastIndexOf(' ')), false));

            this.Loader.AddCalibrationStandard(this.CreateSampleGroup(this.CalibrationStandardsList, "Calibration Standards", false)); // CalibStd
            this.Loader.AddCalibrationSampleGroup(this.CreateSampleGroup(this.CalibrationSamplesList, "Quality Control Solutions", false)); // Instrument Blank
            this.Loader.AddQualityControlSampleGroup(this.CreateSampleGroup(this.QualityControlSamplesList, "Stated Value", true)); // CCV
        }


        private void CreateNewSampleSubList(Sample samp)
        {
            List<Sample> tempList = new List<Sample>();
            tempList.Add(samp);
            this.SampleList.Add(tempList);
        }


        private void CheckForNullSample(Sample samp)
        {
            if (samp == null)
                throw new ArgumentNullException("The sample that was passed in is null\n");
        }

        #endregion
    }
}