using System;
using System.Collections.Generic;
using WaterAnalysisTool.Components;
using OfficeOpenXml;
using WaterAnalysisTool.Exceptions;

namespace WaterAnalysisTool.Analyzer
{
    class AnalyticsLoader
    {
        #region Attributes
        private List<List<List<Element>>> Elements; // each list of elements represents data for one element
        private ExcelPackage DataWorkbook;
        private Double Threshold;
        #endregion

        #region Constructors
        public AnalyticsLoader(ExcelPackage datawb, Double threshold)
        {
            this.DataWorkbook = datawb;
            this.Threshold = threshold;
            this.Elements = new List<List<List<Element>>>();
        }
        #endregion

        #region Public Methods
        public void Load()
        {
            int count = 0;
            int index = 0;
            int row = 1;
            int col = 1;
            Double CoD; // Coefficient of Determination or r squared
            List<Element> e2 = null;

            AnalyticsParser parser = new AnalyticsParser(this.DataWorkbook, this);
            parser.Parse();

            if (Elements.Count < 1)
                throw new ParseErrorException("Problem parsing input Excel workbook. No Sample groups found.");

            this.DataWorkbook.Workbook.Worksheets.Add("Correlation");
            var correlationws = this.DataWorkbook.Workbook.Worksheets[this.DataWorkbook.Workbook.Worksheets.Count]; // should be the last workbook

            // Write outline for correlation matrices
            for(int i = 0; i < Elements.Count; i++)
            {
                col = 1;
                count = 0;

                while(col <= Elements[0].Count)
                {
                    col++;
                    correlationws.Cells[row, col].Value = Elements[0][count][0].Name;
                    count++;
                }

                col = 1;
                count = 0;

                while(row <= Elements[0].Count)
                {
                    row++;
                    correlationws.Cells[row, col].Value = Elements[0][count][0].Name;
                    count++;
                }

                row++;
            }

            // Calculate Coefficient of Determination for each element pair for each sample group
            foreach (List<List<Element>> sg in Elements)
            {
                index = 0;
                count = 1;
                row = 1; // TODO figure out how to get the row properly for multiple sample groups
                
                foreach (List<Element> e1 in sg)
                {
 
                    count = index + 1;

                    while (count <= sg.Count)
                    {
                        e2 = sg[count];
                        CoD = CalculateCoeffiecientOfDetermination(e1, e2);

                        // Write CoD in it's cell
                        correlationws.Cells[row, count].Value = CalculateCoeffiecientOfDetermination(e1, e2).ToString();

                        count++;
                    }

                    index++;
                    row++;
                }
            }
        }

        public void AddElements(List<List<Element>> elements)
        {
            if (elements == null)
                throw new ArgumentNullException("List of elements is null.");

            List<List<Element>> sampleGroup = new List<List<Element>>();

            foreach(List<Element> els in elements)
            {
                List<Element> elementList = new List<Element>();

                foreach(Element e in els)
                {
                    elementList.Add((Element)e.Clone());
                }

                sampleGroup.Add(elementList);
            }

            this.Elements.Add(sampleGroup);
        }
        #endregion

        #region Private Methods
        private Double CalculateCoeffiecientOfDetermination(List<Element> e1, List<Element> e2)
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
        #endregion
    }
}
