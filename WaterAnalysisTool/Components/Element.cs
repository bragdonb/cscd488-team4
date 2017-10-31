using System;

namespace WaterAnalysisTool.Components
{
    class Element
    {
        /* Attributes */
        private String name;
        private String units;
        private double avg;
        private double stddev;
        private double rsd;

        #region Properties
        public String Name
        {
            get { return this.name; }
        }

        public String Units
        {
            get { return this.units; }
        }

        public Double Average
        {
            get { return this.avg; }
        }

        public Double StandardDeviation
        {
            get { return this.stddev; }
        }

        public Double RSD
        {
            get { return this.rsd; }
        }
        #endregion

        /* Constructors */
        public Element(String name, String units, Double avg, Double stddev, Double rsd)
        {
            this.name = name;
            this.units = units;
            this.avg = avg;
            this.stddev = stddev;
            this.rsd = rsd;
        }

        /* Public Functions */
    }
}
