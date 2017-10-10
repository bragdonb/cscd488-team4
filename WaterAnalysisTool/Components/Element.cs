using System;

namespace WaterAnalysisTool.Components
{
    class Element
    {
        /* Attributes */
        private String Name;
        private String Units;
        private Double Average;
        private Double StandardDeviation;
        private Double RSD;

        /* Constructors */
        public Element(String name, String units, Double avg, Double stddev, Double rsd)
        {
            this.Name = name;
            this.Units = units;
            this.Average = avg; ;
            this.StandardDeviation = stddev;
            this.RSD = rsd;
        }

        /* Public Functions */
    }
}
