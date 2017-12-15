using System;
using System.Collections.Generic;
using WaterAnalysisTool.Components;
using OfficeOpenXml;
using WaterAnalysisTool.Exceptions;

namespace WaterAnalysisTool.Analyzer
{
    class AnalyticsParser
    {
        #region Attributes
        private List<List<Element>> elements; // represents all data for one element within at least SampleGroup
        private AnalyticsLoader loader;
        private ExcelPackage dataWorkbook;
        private ExcelWorksheet dataws;
        private int resetRow;
        private int row;
        private int col;
        private List<String> elementNames;
        #endregion

        #region Constructor(s)
        public AnalyticsParser(ExcelPackage datawb, AnalyticsLoader loader)
        {
            this.dataWorkbook = datawb;
            this.loader = loader;
            this.elements = new List<List<Element>>();
            this.dataws = datawb.Workbook.Worksheets[1];
            fillElementNames();
        }
        #endregion

        #region Public Methods
        //TODO parse sample group names as well and pass them to the loader using loader.AddSampleName(sgName)
        public void Parse()
        {
            if (this.dataWorkbook.File.Length < 4 || !this.dataWorkbook.File.Exists)
                throw new ParseErrorException("Data Workbook does not exist or does not have correct contents.");

            this.row = 7;
            this.col = 3;
            int blankLineCount = 0;
            ExcelWorksheet dataws = this.dataWorkbook.Workbook.Worksheets[1]; //data worksheet

            /* Loop reads through file until it encounters the Samples section */
            while (blankLineCount < 2 && blankLineCount >= 0)
            {
                if (this.dataws.Cells[this.row, 1].Value != null)
                {
                    if (!this.dataws.Cells[this.row, 1].Value.ToString().ToLower().Equals("samples"))
                    {
                        this.row++;
                        blankLineCount = 0;
                    }
                    else
                        blankLineCount = -1;
                }
                else
                {
                    blankLineCount++;
                    this.row++;
                }
            }

            if (blankLineCount > 1)
            {
                Console.WriteLine("No samples found in file.");
                return;
            }
            
            /* We have reached the Samples.
               Next line should be the name of SampleGroup, 
               the line after that should be the 
               first sample name within the first SampleGroup.
            */
            this.row += 2;

            while (!isEndOfWorksheet())
            {
                // elements.Add(fillElementList()); // needs to add the sample group's list of element lists to the loader
                fillElementList();
                this.loader.AddElements(this.elements);
                /*TESTING*/
                this.loader.AddSampleName("Test Sample Name");
                this.elements.Clear();
                this.row += 2;
            }
            /*foreach (List<Element> le in elements)
            {
                foreach(Element e in le)
                {
                    Console.WriteLine(e.Name + " " + e.Average);
                }
            }*/
            // this.loader.AddElements(elements);

        }
        #endregion

        #region Private Methods
        private List<Element> fillElementList()
        {
            this.resetRow = this.row;
            int colLength = 0; // TODO I am pretty sure this needs to start at one? Last iteration of first run skips increment for last data point (originally was zero, is probably fine, don't wanna think about it.)
            bool firstRun = true;

            // move inside for loop, new list for each element
            //List<Element> analytes = new List<Element>();

            for (int x = 0; this.dataws.Cells[this.row, this.col].Value != null; x++)
            {
                List<Element> analytes = new List<Element>();

                for (int y = 0; this.dataws.Cells[this.row, this.col].Value != null; y++)
                {
                    /*TODO Testing Console.WriteLine("Trying to parse: " + this.dataws.Cells[this.row, this.col].Value.ToString() + "   " + "\tfor element: " + this.elementNames[x] + "\tx = " + x + "\ty = " + y);*/
                    analytes.Add(new Element(this.elementNames[x], "", Double.Parse(this.dataws.Cells[this.row, this.col].Value.ToString()), 0.0, 0.0));
                    this.row++;
                    if (firstRun)
                        colLength++;
                }
                firstRun = false;
                this.row = this.resetRow;
                this.col++;

                // need to add to the list that represents the sample list then clear the list for re-use, otherwise elements just get added to the same list
                this.elements.Add(analytes);
                //analytes.Clear(); // I lied don't clear
            }

            this.row += colLength;//at blank space after first samplegroup
            this.col = 3;

            // doesn't really need to return now
            //return analytes;
            return null;
        }

        private bool isEndOfWorksheet()
        {
            if (this.dataws.Cells[this.row, this.col].Value != null)
                return false;

            return true;
        }

        private void fillElementNames()
        {
            int column = 3;
            this.elementNames = new List<String>();
            
            while(this.dataws.Cells[5, column].Value != null)
            {
                this.elementNames.Add(this.dataws.Cells[5, column].Value.ToString());
                column++;
            }
        }
        #endregion

    } // end class AnalyticsParser

} // end namespace WaterAnalysisTool.Analyzer
