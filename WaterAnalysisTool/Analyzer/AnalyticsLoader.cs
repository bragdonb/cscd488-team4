using System;
using System.Collections.Generic;
using WaterAnalysisTool.Components;
using OfficeOpenXml;

namespace WaterAnalysisTool.Analyzer
{
    class AnalyticsLoader
    {
        /* Attributes */
        private List<List<Element>> Elements; // each list of elements represents data for one element
        private ExcelPackage DataWorkbook;
        private Double Threshold;

        /* Constructors */
        public AnalyticsLoader(ExcelPackage datawb, Double threshold)
        {
            this.DataWorkbook = datawb;
            this.Threshold = threshold;
            this.Elements = new List<List<Element>>();
        }

        /* Public Methods */
        public void Load()
        {
            int count = 0;
            int index = 0;
            Double CoD; // Coefficient of Determination or r squared

            foreach(List<Element> e1 in Elements)
            {
                List<Element> e2 = null;

                count++;

                while(index < count) // don't need to do calculations multiple times
                {
                    e2 = Elements[index];
                    index++;
                }

                CoD = CalculateCoefiecientOfDetermination(e1, e2);

                if(CoD > this.Threshold)
                {
                    CreateCorrelationMatrix(e1, e2);
                }
            }
        }

        public void AddElements(List<Element> elements)
        {
            if (elements == null)
                throw new ArgumentNullException("List of elements is null.");

            this.Elements.Add(elements);
        }

        /* Private Methods */
        private Double CalculateCoefiecientOfDetermination(List<Element> e1, List<Element> e2)
        {
            // TODO

            return -1.0;
        }

        private void CreateCorrelationMatrix(List<Element> e1, List<Element> e2)
        {
            // TODO
        }
    }
}
