using System;
using System.Collections.Generic;

namespace WaterAnalysisTool.Components
{
    class SampleGroup
    {
        /* Attributes */
        private List<Sample> samples;
        private List<Double> average;
        private List<Double> lod;
        private List<Double> loq;
        private List<Double> percentDifference;
        private List<Double> rsd;
        private List<Double> recovery;
        
        #region Properties
        public List<Double> Average
        {
            get { return this.average; }
        }

        public List<Double> LOQ
        {
            get { return this.loq; }
        }

        public List<Double> LOD
        {
            get { return this.lod; }
        }

        public List<Double> PercentDifference
        {
            get { return this.percentDifference; }
        }

        public List<Double> RSD
        {
            get { return this.rsd; }
        }

        public List<Double> Recovery
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
            CalculateAverage();
            CalculateLODandLOQandRSD();

            //not finished
            CalculatePercentDifference();
            CalculateRecovery();
        }

        private void CalculateAverage()
        {
            this.average = new List<Double>();

            int count = 0, index = 0;

            foreach(Sample s in this.samples)
            {
                count++;
                index = 0;
                foreach(Element e in s.Elements)
                {
                    this.average[index] += e.Average;
                    index++;
                }
            }

            for(index = 0; index < this.average.Count; index++)
                this.average[index] = this.average[index] / count;

        }//end CalculateAverage()

        //maybe change this name.....hahaha
        private void CalculateLODandLOQandRSD()
        {
            this.lod = new List<Double>();
            this.loq = new List<Double>();
            this.rsd = new List<Double>();
            
            int count = 0, index = 0;

            foreach (Sample s in this.samples)
            {
                count++;
                index = 0;
                foreach (Element e in s.Elements)
                {
                    this.lod[index] += Math.Pow((e.Average - this.average[index]), 2);
                    index++;
                }
            }

            for (index = 0; index < this.average.Count; index++)
            {
                this.lod[index] = 3 * Math.Sqrt(this.lod[index] / count);
                this.loq[index] = 10 * Math.Sqrt(this.lod[index] / count);
                this.rsd[index] = (Math.Sqrt(this.lod[index] / count)) / this.average[index] * 100;
            }

        }//end CalculateLODandLOQandRSD()

        private void CalculatePercentDifference()
        {
            this.percentDifference = new List<Double>();
            //TODO Find out how to retrieve Stated Values, which are needed to calculate percent difference
            
        }//end CalculatePercentDifference()

        private void CalculateRecovery()
        {
            this.recovery = new List<Double>();
            //TODO Find out how to retrieve Stated Values, which are needed to calculate recovery(%)

        }//CalculateRecovery()

    }
}
