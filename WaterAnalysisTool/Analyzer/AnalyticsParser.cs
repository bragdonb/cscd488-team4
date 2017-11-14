using System;
using System.Collections.Generic;
using WaterAnalysisTool.Components;
using OfficeOpenXml;

namespace WaterAnalysisTool.Analyzer
{
    class AnalyticsParser
    {
        /* Attributes */
        private List<Element> Elements; // represents all data for one element
        private AnalyticsLoader Loader;
        private ExcelPackage DataWorkbook;

        /* Constructors */
        public AnalyticsParser(ExcelPackage datawb, AnalyticsLoader loader)
        {
            this.DataWorkbook = datawb;
            this.Loader = loader;
            this.Elements = new List<Element>();
        }

        /* Public Methods */
        public void Parse()
        {
            // TODO
        }
    }
}
