using System;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using WaterAnalysisTool.Components;

namespace WaterAnalysisTool.Loader
{
    class DataLoader
    {
        /* Attributes */
        private List<Sample> CalibrationSamples;
        private List<Sample> QualityControlSamples;
        private List<Sample> CertifiedValueSamples;
        private List<Sample> Samples;

        private StreamReader Input;
        private ExcelPackage Output;

        /* Constructors */
        public DataLoader(StreamReader inf, ExcelPackage outf)
        {
            this.Input = inf;
            this.Output = outf;
        }

        /* Public Methods */
        public void Load()
        {
            // TODO
        }

        // TODO
        // Error checking for Add<Sample> functions
        public void AddCalibrationSample(Sample sample)
        {
            this.CalibrationSamples.Add(sample);
        }

        public void AddQualityControlSample(Sample sample)
        {
            this.QualityControlSamples.Add(sample);
        }

        public void AddCertifiedValueSample(Sample sample)
        {
            this.CertifiedValueSamples.Add(sample);
        }

        public void AddSample(Sample sample)
        {
            this.Samples.Add(sample);
        }
    }
}
