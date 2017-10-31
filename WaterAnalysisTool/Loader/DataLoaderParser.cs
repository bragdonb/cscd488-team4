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
        private List<Sample> Samples;

        /* Constructors */
        public DataLoaderParser(DataLoader loader, StreamReader inf)
        {
            this.Loader = loader;
            this.Input = inf;
            this.Samples = new List<Sample>();
        }

        /* Public Methods */
        public void Parse()
        {

            // TODO
            // Parse performs the following functions
            // 1. Read each sample from the input stream
            //  1.1 Create Sample
            //  1.2 Add elements to the sample
            //  1.3 Add the sample to the correct list (using Loader.Add<SampleType> see comments in DataLoader by each list)

            string line = "";
            this.Input.ReadLine(); // Consumes the first line of the file that is always empty

            this.ParseHeader(line);
        }

        public Element CreateElement(String line)
        {
            // TODO
            return null;
        }

        public Sample CreateSample(String name, String comment, String runTime, String sampleType, Int32 repeats)
        {
            // TODO
            return null;
        }

        public void AddElementToSample()
        {

        }

        public void AddSampleToDataLoader()
        {

        }

        public Sample ParseHeader(string line)
        {
            Sample samp = null;

            if (this.Input.Peek() >= 0)
            {
                line = this.Input.ReadLine();

                if (string.Compare(line, "[Sample Header]") == 0)
                {
                    List<string> stringList = new List<string>();
                    while (this.Input.Peek() >= 0)
                    {
                        stringList.Add(this.Input.ReadLine());
                    }

                    samp = new Sample(stringList[1], stringList[3], stringList[7], stringList[8], int.Parse(stringList[11])); // At this point I'm including the Comment member variable (Sample) whether it is blank in the input file or not
                }                                                                                                             // If she doesn't want any blank comments in the .xlsx file I will change it. I just figured this way was easier
            }
            return samp;
        }

        public void parseResults()
        {

        }

        public void parseInternalStandards()
        {

        }
    }
}
