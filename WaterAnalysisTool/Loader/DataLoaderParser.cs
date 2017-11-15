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

        private List<Sample> CalibrationsStandards;     // Calibration Standard -> Sample Type: Cal, These go in the Calibration Standards worksheet Calib Blank, CalibStd
        private List<Sample> CalibrationSamples;        // Insturment Blanks -> Sample Type: QC
        private List<Sample> QualityControlSamples;     // Stated Values (CCV) -> Sample Type: QC
        
        // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC - The analytes found under Check Standards in the xlsx file will not always match up with the analytes of the Certified Value samples
        private List<SampleGroup> CertifiedValueSampleGroup; // The names of the various Certified Values are not guaranteed to be SoilB/TMDW/etc.
        private List<SampleGroup> Samples;



        /* Constructors */

        public DataLoaderParser (DataLoader loader, StreamReader inf)
        {
            this.Loader = loader;
            this.Input = inf;

            this.CalibrationsStandards = new List<Sample>();
            this.CalibrationSamples = new List<Sample>();
            this.QualityControlSamples = new List<Sample>();

            this.CertifiedValueSampleGroup = new List<SampleGroup>();
            this.Samples = new List<SampleGroup>();
        }



        /* Public Methods */

        public void Parse ()
        {
            Sample sample = null;
            List<Sample> certifiedValueSamples; // This will be used to create a list of samples to be passed to a SampleGroup
            List<String> strList = new List<String>();

            strList = this.ParseCheckStandards(strList);
            this.CreateCertifiedValueLists();

            // TODO
            // Parse performs the following functions
            // 1. Read each sample from the input stream
            //  1.1 Create Sample
            //  1.2 Add elements to the sample
            //  1.3 Add the sample to the correct list (using Loader.Add<SampleType> see comments in DataLoader by each list)

            sample = null;

            this.Input.ReadLine(); // Consumes the first line of the file that is always empty

            while (this.Input.Peek() >= 0)
            {
                sample = this.ParseHeader();
                this.ParseResults(sample);
                this.ParseInternalStandards(sample);
            }

            // Add sample to correct list here
            // Before adding:
            // 1. Check the sampleType
            // 2. If the sampleType is CertifiedValueSample or Sample check the name of the sample and add to the correct SampleGroup

            if (sample != null)
            {
                if (String.Compare(sample.SampleType, "Cal") == 0)
                {
                    this.AddCalibrationsStandard(sample);
                }
                else if (String.Compare(sample.Name, "Instrument Blank") == 0) // Assuming all Calibration Samples will be name "Instrument Blank"
                {
                    this.AddCalibrationsSample(sample);
                }
                else if (String.Compare(sample.Name, "CCV") == 0)
                {
                    this.AddQualityControlSample(sample);
                }
                else if (String.Compare(sample.SampleType, ) == 0)
                {
                    // this.AddCertifiedValueSample(sample); // Need to add the sample to a sample group before adding it to the CertifiedValue list
                }
                else if (String.Compare(sample.SampleType, "Unk") == 0)
                {

                }
            }
        }



        /* Private Methods */

        private List<String> ParseCheckStandards (List<String> strList)
        {
            // Function reads in Check Standards from the .xlsx file and returns a String [] containing the names of the Check Standards
            return strList;
        }



        private void CreateCertifiedValueLists ()
        {
            // Function takes the names of the Certified Values and creates a SampleGroup adding the SampleGroup to the list CertifiedValueSampleGroup
        }



        private Sample CreateSample(String name, String comment, String runTime, String sampleType, Int32 repeats)
        {
            // TODO More error checking?

            if (name == null || comment == null || runTime == null || sampleType == null || repeats > -1)
                throw new ArgumentNullException("The sample you are trying to create will contain a null member variable\n");

            return new Sample(name, comment, runTime, sampleType, repeats);
        }



        private SampleGroup CreateSampleGroup (List<Sample> sampleList, String name, bool skipFirst)
        {
            // TODO More error checking?

            if (sampleList == null || name == null)
                throw new ArgumentException("The SampleGroup you are trying to create will contain a null memeber variable\n");

            return new SampleGroup(sampleList, name, skipFirst);
        }



        private void AddSample(SampleGroup sample)
        {
            if (sample == null)
                throw new ArgumentNullException("The sample being added to the List<T> Samples is null\n");

            this.Samples.Add(sample);
        }



        private void AddCalibrationsSample(Sample sample)
        {
            if (sample == null)
                throw new ArgumentNullException("The sample being added to the List<T> CalibrationSamples is null\n");

            this.CalibrationSamples.Add(sample);
        }



        private void AddCalibrationsStandard(Sample sample)
        {
            if (sample == null)
                throw new ArgumentNullException("The sample being added to the List<T> CalibrationsStandards is null\n");

            this.CalibrationsStandards.Add(sample);
        }



        private void AddQualityControlSample(Sample sample)
        {
            if (sample != null)
                throw new ArgumentNullException("The sample being added to the List<T> QualityControlSamples is null\n");

            this.QualityControlSamples.Add(sample);
        }



        private void AddCertifiedValueSampleGroup(SampleGroup sample)
        {
            if (sample != null)
                throw new ArgumentNullException("The sample being added to the List<T> CertifiedValueSamples is null\n");

            this.CertifiedValueSampleGroup.Add(sample);
        }



        /*private void PassSampleGroupsToDataLoader()
        {
            
            for ()
            {
                // this.Loader.AddSampleGroup(new SampleGroup(this.Samples, NAME HERE));
            }

            // this.Loader.AddCalibrationSampleGroup(new SampleGroup(this.CalibrationSamples, NAME HERE));
            // this.Loader.AddCalibrationStandard(new SampleGroup(this.CalibrationsStandards, NAME HERE));
            // this.Loader.AddQualityControlSampleGroup(new SampleGroup(this.QualityControlSamples, NAME HERE));

            for ()
            {
                // this.Loader.AddCertifiedValueSampleGroup(new SampleGroup(this.CertifiedValueSamples, NAME HERE));
            }
        }*/



        private Element CreateElement (String line)
        {
            // TODO
            return null;
        }



        private void AddElementToSample ()
        {

        }



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
    }
}
