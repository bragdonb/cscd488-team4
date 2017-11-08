using System;
using System.Collections.Generic;

namespace WaterAnalysisTool.Components
{
    class SampleGroup
    {
        /* Attributes */
        private String name;
        private List<Sample> samples;//first row contains data from Check Standards file
        private List<Double> average;
        private List<Double> lod;
        private List<Double> loq;
        private List<Double> percentDifference;
        private List<Double> rsd;
        private List<Double> recovery;
        
        #region Properties
        public String Name { get {return this.name;} }

        public List<Double> Average { get { return this.average; } }

        public List<Double> LOQ { get { return this.loq; } }

        public List<Double> LOD { get { return this.lod; } }

        public List<Double> PercentDifference { get { return this.percentDifference; } }

        public List<Double> RSD { get { return this.rsd; } }

        public List<Double> Recovery { get { return this.recovery; } }

        public List<Sample> Samples { get { return this.samples; } }
        #endregion

        /* Constructors */
        public SampleGroup(List<Sample> sampleList, String name)
        {
            this.name = name;
            this.samples = sampleList;
            CalculateAverage();
            CalculateLODandLOQandRSD();

            //not finished
            CalculatePercentDifference();
            CalculateRecovery();
        }

        #region Private Methods
        private void CalculateAverage()
        {
            this.average = new List<Double>();

            int count = 0, index = 0;
            bool firstRow = true;

            foreach(Sample s in this.samples)//start at row + 1
            {
                count++;
                index = 0;
                if (!firstRow)
                {
                    foreach (Element e in s.Elements)
                    {
                        this.average[index] += e.Average;
                        index++;
                    }
                }
                firstRow = false;
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
            bool firstRow = true;

            foreach (Sample s in this.samples) // start at row + 1
            {
                count++;
                index = 0;

                if (!firstRow)
                {
                    foreach (Element e in s.Elements)
                    {
                        this.lod[index] += Math.Pow((e.Average - this.average[index]), 2);
                        index++;
                    }
                }
                firstRow = false;
            }

            for (index = 0; index < this.average.Count; index++)
            {
                this.lod[index] = 3 * Math.Sqrt(this.lod[index] / count);
                this.loq[index] = 10 * Math.Sqrt(this.lod[index] / count);
                this.rsd[index] = (Math.Sqrt(this.lod[index] / count)) / this.average[index] * 100;
            }

        }//end CalculateLODandLOQandRSD()

        private void CalculatePercentDifference() // %difference = (mean - certified value) / certified value * 100
        {
            this.percentDifference = new List<Double>();

            for (int x = 0; x < this.average.Count; x++)
            {
                if (this.samples[0].Elements[x].Average == -1)
                    this.percentDifference[x] = -1;
                else
                    this.percentDifference[x] = (this.average[x] - this.samples[0].Elements[x].Average) / this.samples[0].Elements[x].Average * 100;
            }

        }//end CalculatePercentDifference()

        private void CalculateRecovery() // %recovery = mean / certified value * 100
        {
            this.recovery = new List<Double>();

            for (int x = 0; x < this.average.Count; x++)
            {
                if (this.samples[0].Elements[x].Average == -1)
                    this.recovery[x] = -1;
                else
                    this.recovery[x] = this.average[x] / this.samples[0].Elements[x].Average * 100;
            }
        }//CalculateRecovery()

    }
    #endregion
}
