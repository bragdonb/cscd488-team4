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
            this.ParseHeader();
        }



        public Sample CreateSample(String name, String comment, String runTime, String sampleType, Int32 repeats)
        {
            // TODO Add any needed error checking for passed in Strings
            if (name != null && comment != null && runTime != null && sampleType != null && repeats > -1)
            {

                return new Sample(name, comment, runTime, sampleType, repeats);
            }
            else
                throw new ArgumentNullException("The sample is null \n");
        }



        public void AddCalibrationSample()
        {

        }



        public void AddCalibrationStandard()
        {

        }



        public void AddQualityControlSample()
        {

        }



        public void AddCertifiedValueSample()
        {

        }



        public void AddSample()
        {

        }



        /* Private Methods */

        private Element CreateElement(String line)
        {
            // TODO
            return null;
        }



        private void AddElementToSample ()
        {

        }



        private void ParseHeader ()
        {
            Sample samp;

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

                    samp = this.CreateSample(stringList[1], stringList[3], stringList[7], stringList[8], int.Parse(stringList[11])); // At this point I'm including the Comment member variable (Sample) whether it is blank in the input file or not
                    // Add samp to corresponding sample list
                }                                                                                                                    // If she doesn't want any blank comments in the .xlsx file I will change it. I just figured this way was easier
            }
        }



        private void ParseResults (String line)
        {

        }



        private void ParseInternalStandards (String line)
        {

        }



    }
}
