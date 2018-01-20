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
                else if (string.Compare(samp.SampleType, "Unk") == 0)
                    this.SampleList.Add(samp);
            }


            // Hand off to DataLoader


            foreach (List<Sample> sampleList in this.CertifiedValueList)
            {
                this.Loader.AddCertifiedValueSampleGroup(new SampleGroup(sampleList, "Certified Values", true));
            }

            this.Loader.AddCalibrationStandard(this.CreateSampleGroup(this.CalibrationStandardsList, "Calibration Standards", false)); // CalibStd
            this.Loader.AddCalibrationSampleGroup(this.CreateSampleGroup(this.CalibrationSamplesList, "Quality Control Solutions", false)); // Instrument Blank
            this.Loader.AddQualityControlSampleGroup(this.CreateSampleGroup(this.QualityControlSamplesList, "Stated Value", true)); // CCV
            this.Loader.AddSampleGroup(this.CreateSampleGroup(this.SampleList, "Samples", false));
        }


        /* Private Methods */


        private SampleGroup CreateSampleGroup (List<Sample> sampleList, string name, bool skipFirst)
        {
            if (sampleList == null || name == null)
                throw new ArgumentException("The SampleGroup you are trying to create will contain a null member variable\n");

            return new SampleGroup(sampleList, name, skipFirst);
        }


        private Sample CreateSample(string name, string comment, string runTime, string sampleType, Int32 repeats)
        {
            if (name == null || comment == null || runTime == null || sampleType == null || repeats > -1)
                throw new ArgumentNullException("The sample you are trying to create will contain a null member variable\n");

            return new Sample(name, comment, runTime, sampleType, repeats);
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
            this.Input.ReadLine(); // Consumes the empty line before "[Sample Header]"

            if (this.Input.Peek() >= 0)
            {
                string line = this.Input.ReadLine();

                if (string.Compare(line, "[Sample Header]") == 0)
                {
                    List<string> stringList = new List<string>();

                    while (this.Input.Peek() >= 0)
                    {
                        stringList.Add(this.Input.ReadLine());
                    }

                    // String Trimming
                    stringList[1].Replace("SampleName=", "");
                    stringList[7] = stringList[7].Substring(stringList[7].IndexOf(' ', 4)); // Getting the correct format for the time
                    stringList[8].Replace("Sample Type=", "");
                    stringList[11].Replace("Repeats=", "");

                    sample = this.CreateSample(stringList[1], stringList[3], stringList[7], stringList[8], int.Parse(stringList[11]));

                    return sample;
                }
            }

            return null;
        }


        private void ParseResults (Sample sample)
        {
            this.CheckForNullSample(sample);
            this.Input.ReadLine(); // Consumes empty line before "[Results]"

            if (this.Input.Peek() >= 0)
            {
                string line = this.Input.ReadLine();

                if (string.Compare(line, "[Results]") == 0)
                {
                    line = this.Input.ReadLine(); // Consumes the line containing the labels of the Results section
                    List<string> stringList = new List<string>();

                    while (this.Input.Peek() >= 0)
                    {
                        stringList.Add(this.Input.ReadLine());
                    }

                    foreach (string str in stringList)
                    {
                        string[] strArray = str.Split(',');

                        double avg;
                        double stddev;
                        double rsd;

                        if (!(double.TryParse(strArray[2], out avg)))
                            avg = -1;
                        if (!(double.TryParse(strArray[3], out stddev)))
                            stddev = -1;
                        if (!(double.TryParse(strArray[4], out rsd)))
                            rsd = -1;

                        Element newElement = this.CreateElement(strArray[0], strArray[1], avg, stddev, rsd);
                        this.AddElementToSample(sample, newElement);
                    }
                }
            }
        }


        private void ParseInternalStandards (Sample sample) // Not sure if this method is needed. Ask Carmen
        {
            this.CheckForNullSample(sample);
            this.Input.ReadLine(); // Consumes empty line before "[Internal Standards]"
            string str = this.Input.ReadLine();
            string[] strList = str.Split(',');
        }


        private void CheckForNullSample (Sample samp)
        {
            if (samp == null)
                throw new ArgumentNullException("The sample that was passed in is null\n");
        }
    }
}
