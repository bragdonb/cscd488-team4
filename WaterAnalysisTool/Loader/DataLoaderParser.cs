using System;
using System.IO;
using System.Collections.Generic;
using WaterAnalysisTool.Components;
using WaterAnalysisTool.Exceptions;
using OfficeOpenXml;

namespace WaterAnalysisTool.Loader
{
    class DataLoaderParser
    {


        /* Attributes */


        private DataLoader Loader;
        private StreamReader Input;

        private List<Sample> CalibrationStandardsList; // Calibration Standard -> Sample Type: Cal, --- These go in the Calibration Standards worksheet.  Calib Blank, CalibStd
        private List<Sample> CalibrationSamplesList;  // Quality Control Solutions -> Sample Type: QC --- These are usually named Instrument Blank
        private List<Sample> QualityControlSamplesList;  // Stated Values (CCV) -> Sample Type: QC --- These will have CCV in the name

        private List<List<Sample>> SampleList; // Samples -> Sample Type: Unk --- These will have very different names (Perry/DFW/etc.)
        private List<List<Sample>> CertifiedValueList;  // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC --- The analytes (elements) found under Check Standards in the xlsx file will not always match up with the analytes of the Certified Value samples


        /* Constructors */


        public DataLoaderParser (DataLoader loader, StreamReader inf)
        {
            this.Loader = loader;
            this.Input = inf;

            this.CalibrationStandardsList = new List<Sample>(); // CalibStd
            this.CalibrationSamplesList = new List<Sample>(); // Instrument Blank
            this.QualityControlSamplesList = new List<Sample>(); // CCV

            this.SampleList = new List<List<Sample>>();
            this.CertifiedValueList = new List<List<Sample>>();
        }


        /* Public Methods */


        public void Parse ()
        {
            this.Input.ReadLine(); // Consumes empty line at the beginning of the file
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


        /* Private Methods */


        // SampleGroup, Sample & Element creation


        private SampleGroup CreateSampleGroup (List<Sample> sampleList, string name, bool skipFirst)
        {
            if (sampleList == null || name == null)
                throw new ArgumentException("The SampleGroup you are trying to create will contain a null member variable\n");

            return new SampleGroup(sampleList, name, skipFirst);
        }


        private Sample CreateSample (string method, string name, string comment, string runTime, string sampleType, Int32 repeats)
        {
            if (method == null || name == null || comment == null || runTime == null || sampleType == null || repeats < 0)
                throw new ArgumentNullException("The Sample you are trying to create will contain a null member variable\n");

            return new Sample(method, name, comment, runTime, sampleType, repeats); // TODO see sample constructors
        }

        
        private Element CreateElement (string name, string units, Double avg, Double stddev, Double rsd)
        {
            if (name == null || units == null)
                throw new ArgumentNullException("The Element you are trying to instantiate will contain a null member variable\n");

            return new Element(name, units, avg, stddev, rsd);
        }


        private void AddElementToSample (Sample sample, Element element)
        {
            if (sample == null)
                throw new ArgumentNullException("The sample you are attempting to add an element to is null\n");
            else if (element == null)
                throw new ArgumentNullException("The element you are attempting to add to the sample is null\n");

            sample.Elements.Add(element);
        }


        // Parse() helper methods


        private void ParseCheckStandards()
        {
            FileInfo fi = new FileInfo("CheckStandards.xlsx");
            if (!fi.Exists)
                throw new FileNotFoundException("The CheckStandards.xlsx config file does not exist or could not be found and a calibration curve could not be generated.");

            using (ExcelPackage ep = new ExcelPackage(fi))
            {
                ExcelWorksheet ews = ep.Workbook.Worksheets[2];
                int row = 3; // We start from row 3. Assuming the format of CheckStandards.xlsx will remain the same
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
                        sampleName = ews.Cells[row, col].Value.ToString(); // Reads in the name of the sample. Name may need to be "Stated Value" (check the example output xlsx)
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


        private Sample ParseHeader ()
        {
            Sample sample;
            List<string> stringList = new List<string>();
            string tmp;

            if (this.Input.Peek() >= 0)
                this.Input.ReadLine(); // Consumes "[Sample Header]"
            else
                throw new FormatException("The file used as input is not formatted correctly.\n");

            if (this.Input.Peek() >= 0) // Start reading in sample meta data
            {
                tmp = this.Input.ReadLine();

                while (!(string.IsNullOrEmpty(tmp)))
                {
                    stringList.Add(tmp);

                    if (this.Input.Peek() >= 0)
                        tmp = this.Input.ReadLine();
                    else
                        throw new FormatException("The file used as input is not formatted correctly.\n");
                }

                if (stringList.Count != 12)
                    throw new FormatException("The file used as input is not formatted correctly.\n");
            }
            else
                throw new FormatException("The file used as input is not formatted correctly. ParseHeader4\n");

            stringList[1] = stringList[1].Replace("SampleName=", ""); // SampleName trimming
            stringList[3] = stringList[3].Replace("Comment=", ""); // Comment trimming
            stringList[7] = stringList[7].Substring(stringList[7].IndexOf(' ', 4)); // Getting the correct format for the time
            stringList[8] = stringList[8].Replace("Sample Type=", ""); // SampleType trimming
            stringList[11] = stringList[11].Replace("Repeats=", ""); // Repeats trimming

            sample = CreateSample(stringList[0], stringList[1], stringList[3], stringList[7], stringList[8], int.Parse(stringList[11])); // TODO int.Parse() throws a FormatException (not a number)

            return sample;
        }


        private void ParseResults (Sample sample)
        {
            this.CheckForNullSample(sample);
            List<string> stringList = new List<string>();
            string tmp;

            for (int x = 0; x < 2; x++) // Consumes line containing "[Results]", line containing labels for elements "Elem,Units,Avg,Stddev,RSD"
            {
                if (this.Input.Peek() >= 0)
                    this.Input.ReadLine();
                else
                    throw new FormatException("The file used as input is not formatted correctly\n");
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
                        throw new FormatException("The file used as input is not formatted correctly.\n");
                }
            }

            foreach (string str in stringList) // Data scrubbing and Element creation
            {
                string[] strArray = str.Split(',');

                double avg;
                double stddev;
                double rsd;

                strArray[2] = strArray[2].Replace(" ", ""); // Removes the whitespace character in front avgs

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


        private void ParseInternalStandards (Sample sample)
        {
            this.CheckForNullSample(sample);

            string str;
            string[] strList;

            if (this.Input.Peek() >= 0)
                this.Input.ReadLine(); // Consumes line containing "[Internal Standards]"
            else
                throw new FormatException("The file used as input is not formatted correctly.\n");

            if (this.Input.Peek() >= 0)
            {
                str = this.Input.ReadLine();
                strList = str.Split(','); // Do something with this?
                this.Input.ReadLine(); // Consumes empty line after Internal Standards section
            }
        }


        private void AddSampleToList (Sample samp)
        {
            this.CheckForNullSample(samp);
            bool certifiedValue = false;
            bool newSample = false;

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
                    if (samp.Name.StartsWith(this.CertifiedValueList[x][0].Name.Substring(0, 4))) // SoilB in the input txt file is SoilB. In the CheckStandards.xlsx it is Soil B. Makes it problematic for this line
                    {                                                                              // Because of this condition there are duplicate TMDW lists, one list has a first sample named "TMDW", the other list has a first sample named "TMDW 1:10", Basically need a more robust condition here. The one sample named DFW 17-002 Dup is also in it's own list when it shouldn't be. This is a super long comment
                        this.CertifiedValueList[x].Add(samp);
                        certifiedValue = true;
                    }
                }

                if (!certifiedValue && this.SampleList.Count == 0)
                    this.CreateNewSampleSubList(samp);
                else if (!certifiedValue)
                {
                    for (int x = 0; x < this.SampleList.Count; x++)
                    {
                        if (samp.Name.StartsWith(this.SampleList[x][0].Name.Substring(0, 4)))
                        {
                            this.SampleList[x].Add(samp);
                            newSample = false;
                        }
                        else
                            newSample = true;
                    }
                }

                if (newSample)
                    this.CreateNewSampleSubList(samp);
            }
        }


        private void PassToDataLoader ()
        {
            foreach (List<Sample> sampleList in this.CertifiedValueList)
                this.Loader.AddCertifiedValueSampleGroup(new SampleGroup(sampleList, "Certified Values", true));

            foreach (List<Sample> sampleList in this.SampleList)
                this.Loader.AddSampleGroup(new SampleGroup(sampleList, "Samples", false));

            this.Loader.AddCalibrationStandard(this.CreateSampleGroup(this.CalibrationStandardsList, "Calibration Standards", false)); // CalibStd
            this.Loader.AddCalibrationSampleGroup(this.CreateSampleGroup(this.CalibrationSamplesList, "Quality Control Solutions", false)); // Instrument Blank
            this.Loader.AddQualityControlSampleGroup(this.CreateSampleGroup(this.QualityControlSamplesList, "Stated Value", true)); // CCV
        }


        /* private void CreateNewCertifiedValueSubList (Sample samp)
        {
            List<Sample> tempList = new List<Sample>();
            tempList.Add(samp);
            this.CertifiedValueList.Add(tempList); // Soil B, TMDW etc.
        } */


        private void CreateNewSampleSubList (Sample samp)
        {
            List<Sample> tempList = new List<Sample>();
            tempList.Add(samp);
            this.SampleList.Add(tempList);
        }


        private void CheckForNullSample (Sample samp)
        {
            if (samp == null)
                throw new ArgumentNullException("The sample that was passed in is null\n");
        }
    }
}
