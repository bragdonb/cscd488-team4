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
    }
}
