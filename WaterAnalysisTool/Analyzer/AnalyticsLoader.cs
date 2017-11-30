using System;
using System.Collections.Generic;
using WaterAnalysisTool.Components;
using OfficeOpenXml;
using WaterAnalysisTool.Exceptions;

namespace WaterAnalysisTool.Analyzer
{
    class AnalyticsLoader
    {
        /* Attributes */
        private List<List<List<Element>>> Elements; // each list of elements represents data for one element
        private ExcelPackage DataWorkbook;
        private Double Threshold;

        /* Constructors */
        public AnalyticsLoader(ExcelPackage datawb, Double threshold)
        {
            this.DataWorkbook = datawb;
            this.Threshold = threshold;
            this.Elements = new List<List<List<Element>>>();
        }

        /* Public Methods */
        public void Load()
        {
            int count = 0;
            int index = 0;
            Double CoD; // Coefficient of Determination or r squared

            AnalyticsParser parser = new AnalyticsParser(this.DataWorkbook, this);
            parser.Parse();

            if (this.Elements.Count < 2) // something went weird yo, needs to be at least two elements to create a correlation matrix
                throw new ParseErrorException("Cannot create coefficient matrices for input file. Number of elements per sample in the input file must be greater than 1.");

            foreach (List<List<Element>> sg in Elements)
            {
                foreach (List<Element> e1 in sg)
                {
                    List<Element> e2 = null;

                    count++;

                    while (index < count) // don't need to do calculations multiple times
                    {
                        e2 = sg[index];
                        index++;
                    }

                    CoD = CalculateCoefiecientOfDetermination(e1, e2);

                    if (CoD > this.Threshold)
                    {
                        CreateCorrelationMatrix(e1, e2);
                    }
                }
            }
        }

        public void AddElements(List<List<Element>> elements)
        {
            if (elements == null)
                throw new ArgumentNullException("List of elements is null.");

            this.Elements.Add(elements);
        }

        /* Private Methods */
        private Double CalculateCoefiecientOfDetermination(List<Element> e1, List<Element> e2)
        {
            Double stdev1 = CalculateElementStandardDeviation(e1);
            Double stdev2 = CalculateElementStandardDeviation(e2);

            Double coVar = CalculateElementCovariance(e1, e2);

            return Math.Pow((coVar / (stdev1 * stdev2)), 2.0);
        }

        private Double CalculateElementStandardDeviation(List<Element> els)
        {
            if (els.Count < 1)
                throw new ArgumentException("To calculate standard deviation the length of the set must be greater than 0");

            Double avg = 0;
            foreach (Element e in els)
                avg += e.Average;

            avg = avg / els.Count;

            Double sum = 0;
            foreach (Element e in els)
                sum += e.Average * e.Average;

            Double sumavg = sum / (els.Count - 1);

            return Math.Sqrt(sumavg - (avg * avg));
        }

        private Double CalculateElementCovariance(List<Element> e1, List<Element> e2)
        {
            if (e1.Count != e2.Count || e1.Count < 1 || e2.Count < 1)
                throw new ArgumentException("To calculate covariance the length of both sets must be equal and greater than 0.");

            Double avg1 = 0;
            foreach (Element e in e1)
                avg1 += e.Average;

            avg1 = avg1 / e1.Count;

            Double avg2 = 0;
            foreach (Element e in e2)
                avg2 += e.Average;

            avg2 = avg2 / e2.Count;

            int index = 0;
            Double sum = 0;
            foreach(Element e in e1)
            {
                sum += (e.Average - avg1) * (e2[index].Average - avg2);
                index++;
            }

            return sum / (e1.Count - 1);
        }

        private void CreateCorrelationMatrix(List<Element> e1, List<Element> e2)
        {
            // TODO
        }
    }
}
