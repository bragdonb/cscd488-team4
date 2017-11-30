using System;
using System.Collections.Generic;
using WaterAnalysisTool.Components;
using OfficeOpenXml;

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
        public void Parse()
        {
            this.row = 7;
            this.col = 3;
            int blankLineCount = 0;
            ExcelWorksheet dataws = this.dataWorkbook.Workbook.Worksheets[1]; //data worksheet

            while (blankLineCount < 2)
            {
                if (this.dataws.Cells[this.row, 1].Value != null)
                {
                    if (!this.dataws.Cells[this.row, 1].Value.ToString().ToLower().Equals("samples"))
                    {
                        this.row++;
                        blankLineCount = 0;
                    }
                    else
                        blankLineCount = 2;
                }
                else
                {
                    blankLineCount++;
                    this.row++;
                }           
           }

            //now at Samples row. Next line is name of SampleGroup, line after that is first sample name within the first sample group.
            this.row += 2;

            while (!isEndOfWorksheet())
            {
                elements.Add(fillElementList());
                this.row += 2;
            }

            this.loader.AddElements(elements);

        }
        #endregion

        #region Private Methods
        private List<Element> fillElementList()
        {
            this.resetRow = this.row;
            int colLength = 0;
            bool firstRun = true;

            List<Element> analytes = new List<Element>();

            for (int x = 0; this.dataws.Cells[this.row, this.col].Value != null; x++)
            {
                for (int y = 0; this.dataws.Cells[this.row, this.col].Value != null; y++)
                {
                    analytes.Add(new Element(this.elementNames[x], "", Double.Parse(this.dataws.Cells[this.row, this.col].Value.ToString()), 0.0, 0.0));
                    this.row++;
                    if (firstRun)
                        colLength++;
                }
                firstRun = false;
                this.row = this.resetRow;
                this.col++;
            }

            this.row += colLength;//at blank space after first samplegroup
            this.col = 3;

            return analytes;
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
