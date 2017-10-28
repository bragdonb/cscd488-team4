using System;

namespace WaterAnalysisTool.Components
{
    class Element
    {
        /* Attributes */
        private int l = 1;

        #region Properties
        private String Name
        {
            set
            {
                if (l == 0)
                {
                    this.Name = value;
                }
            }
            get { return this.Name; }
        }

        private String Units
        {
            set
            {
                if (l == 0)
                {
                    this.Units = value;
                }
            }
            get { return this.Units; }
        }

        private Double Average
        {
            set
            {
                if (l == 0)
                {
                    this.Average = value;
                }
            }
            get { return this.Average; }
        }

        private Double StandardDeviation
        {
            set
            {
                if (l == 0)
                {
                    this.StandardDeviation = value;
                }
            }
            get { return this.StandardDeviation; }
        }

        private Double RSD
        {
            set
            {
                if (l == 0)
                {
                    this.RSD = value;
                }
            }
            get { return this.RSD; }
        }
        #endregion

        /* Constructors */
        public Element(String name, String units, Double avg, Double stddev, Double rsd)
        {
            l = 0;
            this.Name = name;
            this.Units = units;
            this.Average = avg;
            this.StandardDeviation = stddev;
            this.RSD = rsd;
            l = 1;
        }

        /* Public Functions */
    }
}
