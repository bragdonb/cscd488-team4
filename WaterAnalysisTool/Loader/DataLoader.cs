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
            // tbh I don't remember why I included the input file...
            this.Input = inf;
            this.Output = outf;
        }

        /* Public Methods */
        public void Load()
        {
            // TODO
            // Load performs the following functions:
            // 1. Write QC Sample, Certified Val Sample, and Sample data into the Data worksheet
            //  1.1 Method Header, Analysis Date, and Sample Name Descriptor in as first three rows
            //  1.2 Bolded Element Name headers (x2 one for mg/L and another for RSD)
            //  1.3 Bolded Units (x2 one for mg/L and another for RSD)
            //  1.4 Write QC, 
            // 2. Write Calibration Sample data into the Calibration Standards Worksheet
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
