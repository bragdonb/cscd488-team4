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
        private List<Sample> CalibrationSamples;        // Quality Control Solutions (Insturment Blanks) -> Sample Type: QC
        private List<Sample> CalibrationsStandards;     // Calibration Standard -> Sample Type: Cal
        private List<Sample> QualityControlSamples;     // Stated Values (CCV) -> Sample Type: QC
        private List<Sample> CertifiedValueSamples;     // Certified Values (SoilB/TMDW/etc.) -> Sample Type: QC
        private List<Sample> Samples;



        /* Constructors */

        public DataLoaderParser (DataLoader loader, StreamReader inf)
        {
            this.Loader = loader;
            this.Input = inf;
            this.CalibrationSamples = new List<Sample>();
            this.CalibrationsStandards = new List<Sample>();
            this.QualityControlSamples = new List<Sample>();
            this.CertifiedValueSamples = new List<Sample>();
            this.Samples = new List<Sample>();
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



        private Sample CreateSample(String name, String comment, String runTime, String sampleType, Int32 repeats)
        {
            // TODO More error checking?

            if (name == null || comment == null || runTime == null || sampleType == null || repeats > -1)
                throw new ArgumentNullException("The sample you are trying to create will contain a null member variable\n");

            return new Sample(name, comment, runTime, sampleType, repeats);
        }



        private void AddSample(Sample sample)
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



        private void AddCertifiedValueSample(Sample sample)
        {
            if (sample != null)
                throw new ArgumentNullException("The sample being added to the List<T> CertifiedValueSamples is null\n");

            this.CertifiedValueSamples.Add(sample);
        }



        private void PassSampleGroupsToDataLoader()
        {
            this.Loader.AddSampleGroup(new SampleGroup(this.Samples));
            this.Loader.AddCalibrationSampleGroup(new SampleGroup(this.CalibrationSamples));
            this.Loader.AddCalibrationStandard(new SampleGroup(this.CalibrationsStandards));
            this.Loader.AddQualityControlSampleGroup(new SampleGroup(this.QualityControlSamples));
            this.Loader.AddCertifiedValueSampleGroup(new SampleGroup(this.CertifiedValueSamples));
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
                                                                                                                                       // If she doesn't want any blank comments in the .xlsx file I will change it. I just figured this way was easier

                    // this.ParseResults(sample);

                    switch (sample) // Probably need getter for sampleType
                    {
                        // Add sample to the correct sample list depending on sampleType
                    }
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
