using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using WaterAnalysisTool.Components;

namespace WaterAnalysisTool.Loader
{

    class DataLoaderParser
    {
        /* Attributes */

        private DataLoader Loader;
        private StreamReader Input;
        private List<Sample> CalibrationSamples;        // Quality Control Solutions (Insturment Blanks) -> Sample Type: QC - These will not always have the same number of elements(analytes) in the input text file
        private List<Sample> CalibrationStandards;     // Calibration Standard -> Sample Type: Cal
        private List<Sample> QualityControlSamples;     // Stated Values (CCV) -> Sample Type: QC
        
        // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC - These will not always have the same number of elements(analytes) in the input text file
        private List<SampleGroup> CertifiedValueSamples; // The names of the different groupings of Certified Values can be anything and there can be any number of different names
        private List<SampleGroup> Samples;



        /* Constructors */

        public DataLoaderParser (DataLoader loader, StreamReader inf)
        {
            this.Loader = loader;
            this.Input = inf;
            this.CalibrationSamples = new List<Sample>();
            this.CalibrationStandards = new List<Sample>();
            this.QualityControlSamples = new List<Sample>();

            this.CertifiedValueSamples = new List<SampleGroup>();
            this.Samples = new List<SampleGroup>();
        }



        /* Public Methods */

        public void Parse ()
        {
            // TODO
            // Parse performs the following functions
            // 1. Read each sample from the input stream
            //  1.1 Create Sample
            //  1.2 Add elements to the sample
            //  1.3 Add the sample to the correct list (using Loader.Add<SampleType> see comments in DataLoader by each list)

            this.Input.ReadLine(); // Consumes the first line of the file that is always empty

            while (this.Input.Peek() >= 0)
            {
                this.ParseHeader();
            }
        }



        /* Private Methods */

        private void CreateCertifiedValueLists () // May not need
        {
            // Read Certified Value names from a separate file and create SampleGroups (TMDW, Soil B, CCV (Continuous Calibration Verification)) to be added to the List<SampleGroup> CertifiedValueSampleGroups
        }



        private Sample CreateSample(String name, String comment, String runTime, String sampleType, Int32 repeats)
        {
            // TODO More error checking?

            if (name == null || comment == null || runTime == null || sampleType == null || repeats > -1)
                throw new ArgumentNullException("The sample you are trying to create will contain a null member variable\n");

            return new Sample(name, comment, runTime, sampleType, repeats);
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

            this.CalibrationStandards.Add(sample);
        }



        private void AddQualityControlSample(Sample sample)
        {
            if (sample != null)
                throw new ArgumentNullException("The sample being added to the List<T> QualityControlSamples is null\n");

            this.QualityControlSamples.Add(sample);
        }



        private void AddCertifiedValueSample(SampleGroup sample)
        {
            if (sample != null)
                throw new ArgumentNullException("The sample being added to the List<T> CertifiedValueSamples is null\n");

            this.CertifiedValueSamples.Add(sample);
        }



        private void PassSampleGroupsToDataLoader()
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
        }



        private Element CreateElement (String line)
        {
            // TODO
            return null;
        }



        private void AddElementToSample ()
        {

        }



        private void ParseHeader ()
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

                    sample = this.CreateSample(stringList[1], stringList[3], stringList[7], stringList[8], int.Parse(stringList[11])); // At this point I'm including the Comment member variable (Sample) whether it is blank in the input file or not

                    // Add sample to correct list here
                    // Before adding:
                        // 1. Check the sampleType
                        // 2. If the sampleType is CertifiedValueSample or Sample check the name of the sample and add to the correct SampleGroup

                    // this.ParseResults(sample);

                    //switch (sample) // Probably need getter for sampleType
                    //{
                        // Add sample to the correct sample list depending on sampleType
                    //}
                }
            }
        }



        private void ParseResults (Sample sample)
        {

        }



        private void ParseInternalStandards ()
        {

        }



    }
}
