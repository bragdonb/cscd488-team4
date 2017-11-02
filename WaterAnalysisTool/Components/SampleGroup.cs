using System;
using System.Collections.Generic;

namespace WaterAnalysisTool.Components
{
    class SampleGroup
    {
        /* Attributes */
        private List<Sample> samples;
        private List<Double> average;
        private List<Double> loq;
        private List<Double> lod;
        private List<Double> percentDifference;
        private List<Double> rsd;
        private List<Double> recovery;
        
        #region Properties
        public Double Average
        {
            get { return this.average; }
        }

        public Double LOQ
        {
            get { return this.loq; }
        }

        public Double LOD
        {
            get { return this.lod; }
        }

        public Double PercentDifference
        {
            get { return this.percentDifference; }
        }

        public Double RSD
        {
            get { return this.rsd; }
        }

        public Double Recovery
        {
            get { return this.recovery; }
        }

        public List<Sample> Samples 
        {
            get { return this.samples; }
        }
        #endregion

        /* Constructors */
        public SampleGroup(List<Sample> sampleList)
        {
            this.samples = sampleList;

            this.loq = CalculateLOQ();
            this.lod = CalculateLOD();
            this.percentDifference = CalculatePercentDifference();
            this.rsd = CalculateRSD();
            this.recovery = CalculateRecovery();
        }

        private List<Double> CalculateAverage()
        {
            foreach(Sample s in this.samples)
            {
                foreach(Element e in s.Elements)
                {
                    
                }
            }
            return null;
        }

        private List<Double> CalculateLOQ()
        {
            return null;
        }

        private List<Double> CalculateLOD()
        {
            return null;
        }

        private List<Double> CalculatePercentDifference()
        {
            return null;
        }

        private List<Double> CalculateRSD()
        {
            return null;
        }

        private List<Double> CalculateRecovery()
        {
            return null;
        }
    }
}
