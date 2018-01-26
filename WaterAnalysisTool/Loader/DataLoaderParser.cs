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
        private List<Sample> SampleList; // Samples -> Sample Type: Unk --- These will have very different names (Perry/DFW/etc.)
        private List<List<Sample>> CertifiedValueList;  // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC --- The analytes (elements) found under Check Standards in the xlsx file will not always match up with the analytes of the Certified Value samples
                                                        // The names of the various Certified Values are not guaranteed to be SoilB/TMDW/etc. --- These can have different names with each run


        /* Constructors */


        public DataLoaderParser (DataLoader loader, StreamReader inf)
        {
            this.Loader = loader;
            this.Input = inf;

            this.CalibrationStandardsList = new List<Sample>(); // CalibStd
            this.CalibrationSamplesList = new List<Sample>(); // Instrument Blank
            this.QualityControlSamplesList = new List<Sample>(); // CCV
            this.SampleList = new List<Sample>();
            this.CertifiedValueList = new List<List<Sample>>();
        }


        /* Public Methods */


        public void Parse ()
        {
            Sample samp = null;

            while (this.Input.Peek() >= 0)
            {
                samp = this.ParseHeader();
                this.ParseResults(samp);
                this.ParseInternalStandards(samp);
            }

            this.AddSampleToList(samp);
            this.PassToDataLoader();
        }


        /* Private Methods */


        private SampleGroup CreateSampleGroup (List<Sample> sampleList, string name, bool skipFirst)
        {
            if (sampleList == null || name == null)
                throw new ArgumentException("The SampleGroup you are trying to create will contain a null member variable\n");

            return new SampleGroup(sampleList, name, skipFirst);
        }


        private void PassToDataLoader ()
        {
            foreach (List<Sample> sampleList in this.CertifiedValueList)
            {
                this.Loader.AddCertifiedValueSampleGroup(new SampleGroup(sampleList, "Certified Values", true));
            }

            this.Loader.AddCalibrationStandard(this.CreateSampleGroup(this.CalibrationStandardsList, "Calibration Standards", false)); // CalibStd
            this.Loader.AddCalibrationSampleGroup(this.CreateSampleGroup(this.CalibrationSamplesList, "Quality Control Solutions", false)); // Instrument Blank
            this.Loader.AddQualityControlSampleGroup(this.CreateSampleGroup(this.QualityControlSamplesList, "Stated Value", true)); // CCV
            this.Loader.AddSampleGroup(this.CreateSampleGroup(this.SampleList, "Samples", false));
        }


        private Sample CreateSample (string method, string name, string comment, string runTime, string sampleType, Int32 repeats)
        {
            if (method == null || name == null || comment == null || runTime == null || sampleType == null || repeats < 0)
                throw new ArgumentNullException("The sample you are trying to create will contain a null member variable\n");

            return new Sample(method, name, comment, runTime, sampleType, repeats); // TODO see sample constructors
        }


        private void AddSampleToList (Sample samp)
        {
            if (samp != null)
            {
                if (string.Compare(samp.SampleType, "Cal") == 0)
                    this.CalibrationStandardsList.Add(samp); // CalibStd
                else if (samp.Name.StartsWith("CCV"))
                    this.QualityControlSamplesList.Add(samp); // CCV
                else if (string.Compare(samp.Name, "Instrument Blank") == 0)
                    this.CalibrationSamplesList.Add(samp); // Assuming all Calibration Samples will be named "Instrument Blank" at this point
                else if (string.Compare(samp.SampleType, "QC") == 0)
                {
                    foreach (List<Sample> sampleList in this.CertifiedValueList)
                    {
                        if (string.Compare(sampleList[0].Name, samp.Name) == 0) // This condition may not be correct. Need to ask Carmen
                            sampleList.Add(samp);
                        else
                        {
                            List<Sample> tempList = new List<Sample>();
                            tempList.Add(samp);
                            this.CertifiedValueList.Add(tempList); // Soil B, TMDW etc.
                        }
                    }
                }
                else if (string.Compare(samp.SampleType, "Unk") == 0) // Certified Values may have a sample type of "Unk"
                    this.SampleList.Add(samp);
            }
            else
                throw new ArgumentNullException("The sample that was to be added to a list was null\n");
        }


        private Element CreateElement (string name, string units, Double avg, Double stddev, Double rsd)
        {
            if (name == null || units == null)
                throw new ArgumentNullException("The element you are trying to instantiate will contain a null member variable\n");

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


        private Sample ParseHeader ()
        {
            Sample sample;
            List<string> stringList = new List<string>();
            string tmp;

            for (int x = 0; x < 2; x++) // Consumes empty line before "[Sample Header]" as well as line containing "[Sample Header]"
            {
                if (this.Input.Peek() >= 0)
                    this.Input.ReadLine();
                else
                    throw new FormatException("The file used as input is not formatted correctly.\n");
            }

            if (this.Input.Peek() >= 0) // Read in sample meta data
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
                throw new FormatException("The file used as input is not formatted correctly.\n");

            // String trimming and sample creation

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

            for (int x = 0; x < 3; x++) // Consumes empty line before "[Results]", line containing "[Results]", line containing labels for elements "Elem,Units,Avg,Stddev,RSD"
            {
                if (this.Input.Peek() >= 0)
                    this.Input.ReadLine();
                else
                    throw new FormatException("The file used as input is not formatted correctly.\n");
            }

            if (this.Input.Peek() >= 0) // Reading in elements
            {
                tmp = this.Input.ReadLine();

                while (string.IsNullOrEmpty(tmp))
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


        private void ParseInternalStandards (Sample sample) // Not sure if this method is needed. Ask Carmen
        {
            this.CheckForNullSample(sample);

            string str;
            string[] strList;

            for (int x = 0; x < 2; x++) // Consumes empty line before "[Internal Standards]" and line containing "[Internal Standards]"
            {
                if (this.Input.Peek() >= 0)
                    this.Input.ReadLine();
                else
                    throw new FormatException("The file used as input is not formatted correctly.\n");
            }

            if (this.Input.Peek() >= 0)
            {
                str = this.Input.ReadLine();
                strList = str.Split(','); // Do something with this?
            }
        }


        private void CheckForNullSample (Sample samp)
        {
            if (samp == null)
                throw new ArgumentNullException("The sample that was passed in is null\n");
        }
    }
}
