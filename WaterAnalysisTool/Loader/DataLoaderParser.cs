using System;
using System.IO;
using System.Collections.Generic;
using WaterAnalysisTool.Components;

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

            while (this.Input.Peek() >= 0)
            {
                Sample samp = this.ParseHeader();
                this.ParseResults(samp);
                this.ParseInternalStandards(samp);
                this.AddSampleToList(samp);
            }

            this.CombSampleList();
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

            if (string.Compare(samp.SampleType, "Cal") == 0)
                this.CalibrationStandardsList.Add(samp); // CalibStd
            else if (samp.Name.StartsWith("CCV"))
                this.QualityControlSamplesList.Add(samp); // CCV
            else if (string.Compare(samp.Name, "Instrument Blank") == 0)
                this.CalibrationSamplesList.Add(samp); // Assuming all Calibration Samples will be named "Instrument Blank"
            else if (string.Compare(samp.SampleType, "QC") == 0)
            {
                if (this.CertifiedValueList.Count == 0)
                    this.CreateNewCertifiedValueSubList(samp);
                else
                {
                    foreach (List<Sample> sampleList in this.CertifiedValueList)
                    {
                        if (samp.Name.StartsWith(sampleList[0].Name.Substring(0, 4))) //
                            sampleList.Add(samp);
                        else
                            this.CreateNewCertifiedValueSubList(samp);
                    }
                }
            }
            else if (string.Compare(samp.SampleType, "Unk") == 0)
            {
                if (this.SampleList.Count == 0)
                    this.CreateNewSampleSubList(samp);
                else
                {
                    foreach (List<Sample> sampleList in this.SampleList)
                    {
                        if (samp.Name.StartsWith(sampleList[0].Name.Substring(0, 4))) //
                            sampleList.Add(samp);
                        else
                            this.CreateNewSampleSubList(samp);
                    }
                }
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


        private void CombSampleList () // This method will comb through List<List<Sample>> SampleList and pull out certified value samples that have sample type "Unk", placing them in List<List<Sample>> CertifiedValueList
        {
            foreach (List<Sample> sampleList in this.SampleList)
            {
                foreach (List<Sample> certifiedValueList in this.CertifiedValueList)
                {
                    if (sampleList[0].Name.StartsWith(certifiedValueList[0].Name.Substring(0, 4)))
                        // 
                }
            }
        }


        private void CreateNewCertifiedValueSubList (Sample samp)
        {
            List<Sample> tempList = new List<Sample>();
            tempList.Add(samp);
            this.CertifiedValueList.Add(tempList); // Soil B, TMDW etc.
        }


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
