using System;
using System.Collections.Generic;

namespace WaterAnalysisTool.Components
{
    class Sample
    {
        /* Attributes */
        private List<Element> Elements;
        private String Name;
        private String Comment;
        private String RunTime;
        private String SampleType;
        private Int32 Repeats;

        /* Constructors */
        public Sample(String name, String comment, String runTime, String sampleType, Int32 rpts)
        {
            this.Name = name;
            this.Comment = comment;
            this.RunTime = runTime;
            this.SampleType = sampleType;
            this.Repeats = rpts;
            this.Elements = new List<Element>();
        }

        public Sample(String name, String runTime, String sampleType, Int32 rpts)
        {
            this.Name = name;
            this.Comment = "";
            this.RunTime = runTime;
            this.SampleType = sampleType;
            this.Repeats = rpts;
            this.Elements = new List<Element>();
        }

        /* Public functions */
        public void AddElement(Element e)
        {
            // TODO
            // Error checking e

            this.Elements.Add(e);
        }
    }
}
