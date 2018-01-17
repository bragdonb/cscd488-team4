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

        private List<Sample> CalibrationStandardsList;
        private List<Sample> CalibrationSamplesList;
        private List<Sample> QualityControlSamplesList;

        private List<List<Sample>> CertifiedValueList;
        private List<List<Sample>> SampleList;

        private SampleGroup CalibrationStandards;     // Calibration Standard -> Sample Type: Cal, --- These go in the Calibration Standards worksheet.  Calib Blank, CalibStd
        private SampleGroup CalibrationSamples;        // Quality Control Solutions -> Sample Type: QC --- These are usually named Instrument Blank
        private SampleGroup QualityControlSamples;     // Stated Values (CCV) -> Sample Type: QC --- These will have CCV in the name

        // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC --- The analytes (elements) found under Check Standards in the xlsx file will not always match up with the analytes of the Certified Value samples
        private List<SampleGroup> CertifiedValueSampleGroups; // The names of the various Certified Values are not guaranteed to be SoilB/TMDW/etc. --- These can have different names with each run
        private List<SampleGroup> SampleSampleGroups; // Samples -> Sample Type: Unk --- These will have very different names (Perry/DFW/etc.)

        private List<Sample> CertifiedValuesCheckStandards; // The Check Standards samples are found in the .xlsx



        /* Constructors */



        public DataLoaderParser (DataLoader loader, StreamReader inf)
        {
            this.Loader = loader;
            this.Input = inf;

            this.CalibrationStandardsList = new List<Sample>();
            this.CalibrationSamplesList = new List<Sample>();
            this.QualityControlSamplesList = new List<Sample>();

            this.CertifiedValueList = new List<List<Sample>>();
            this.SampleList = new List<List<Sample>>();

            this.CertifiedValueSampleGroups = new List<SampleGroup>();
            this.SampleSampleGroups = new List<SampleGroup>();
        }



        /* Public Methods */



        public void Parse ()
        {
            Sample samp = null;

            this.CertifiedValuesCheckStandards = this.ParseCheckStandards(); // Still need to parse CheckStandards file

            this.Input.ReadLine(); // Consumes the first line of the file that is always empty

            while (this.Input.Peek() >= 0)
            {
                samp = this.ParseHeader();
                this.ParseResults(samp);
                this.ParseInternalStandards(samp);
            }

            if (samp != null)
            {
                if (String.Compare(samp.SampleType, "Cal") == 0)
                    this.CalibrationStandardsList.Add(samp);
                else if (samp.Name.StartsWith("CCV"))
                    this.QualityControlSamplesList.Add(samp);
                else if (String.Compare(samp.Name, "Instrument Blank") == 0)
                    this.CalibrationSamplesList.Add(samp); // Assuming all Calibration Samples will be name "Instrument Blank" at this point
                else if (String.Compare(samp.SampleType, "QC") == 0)
                {
                    foreach (List<Sample> sampleList in this.CertifiedValueList)
                    {
                        foreach (Sample sample in sampleList)
                        {
                            if (String.Compare(sample.Name, samp.Name) == 0) // This condition is not entirely correct
                            {
                                sampleList.Add(samp);
                            }
                            else
                            {
                                List<Sample> tempList = new List<Sample>();
                                tempList.Add(samp);
                                this.CertifiedValueList.Add(tempList);
                            }
                        }
                    }
                }
                else if (String.Compare(samp.SampleType, "Unk") == 0)
                {
                    foreach (List<Sample> sampleList in this.SampleList)
                    {
                        foreach (Sample sample in sampleList)
                        {
                            if (String.Compare(sample.Name, samp.Name) == 0) // This condition is not entirely correct
                            {
                                sampleList.Add(samp);
                            }
                            else
                            {
                                List<Sample> tempList = new List<Sample>();
                                tempList.Add(samp);
                                this.SampleList.Add(tempList);
                            }
                        }
                    }
                }
            }

            // Create SampleGroups here && hand them off to DataLoader

        }



        /* Private Methods */



        // Methods for creating SampleGroups



        private SampleGroup CreateSampleGroup (List<Sample> sampleList, String name, bool skipFirst)
        {
            // TODO More error checking?

            if (sampleList == null || name == null)
                throw new ArgumentException("The SampleGroup you are trying to create will contain a null member variable\n");

            return new SampleGroup(sampleList, name, skipFirst);
        }



        // Methods for creating Elements and adding them to a sample



        private Element CreateElement (String name, String units, Double avg, Double stddev, Double rsd)
        {
            // TODO More error checking?

            if (name == null || units == null)
                throw new ArgumentNullException("The element you are trying to instantiate will contain a null member variable\n");

            return new Element(name, units, avg, stddev, rsd);
        }



        private void AddElementToSample (Sample sample, Element element)
        {
            if (sample == null)
                throw new ArgumentException("The sample you are attempting to add an element to is null\n");
            else if (element == null)
                throw new ArgumentNullException("The element you are attempting to add to the sample is null\n");

            sample.Elements.Add(element);
        }



        // Methods for creating a sample and adding it to the correct SampleGroup



        private Sample CreateSample(String name, String comment, String runTime, String sampleType, Int32 repeats)
        {
            // TODO More error checking?

            if (name == null || comment == null || runTime == null || sampleType == null || repeats > -1)
                throw new ArgumentNullException("The sample you are trying to create will contain a null member variable\n");

            return new Sample(name, comment, runTime, sampleType, repeats);
        }



        // Parse() helper methods



        private Sample ParseHeader ()
        {
            Sample sample;

            if (this.Input.Peek() >= 0)
            {
                String line = this.Input.ReadLine();

                if (String.Compare(line, "[Sample Header]") == 0)
                {
                    List<String> stringList = new List<String>();

                    while (this.Input.Peek() >= 0)
                    {
                        stringList.Add(this.Input.ReadLine());
                    }

                    // String Trimming
                    stringList[1].Replace("SampleName=", "");
                    stringList[7] = stringList[7].Substring(stringList[7].IndexOf(" ", 4)); // Getting the correct format for the time
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

        }



        private void ParseInternalStandards (Sample sample)
        {

        }



        private List<Sample> ParseCheckStandards()
        {
            return new List<Sample>();
        }
    }
}
