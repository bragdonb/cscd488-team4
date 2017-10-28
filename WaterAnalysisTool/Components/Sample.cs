using System;
using System.Collections.Generic;

namespace WaterAnalysisTool.Components
{
    class Sample
    {
        /* Attributes */
        private int l = 1;

        #region Properties
        private List<Element> Elements
        {
            set
            {
                if(l == 0)
                {
                    this.Elements = value;
                }
            }
            get { return this.Elements; }
        }

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

        private String Comment
        {
            set
            {
                if (l == 0)
                {
                    this.Comment = value;
                }
            }
            get { return this.Comment; }
        }

        private String RunTime
        {
            set
            {
                if (l == 0)
                {
                    this.RunTime = value;
                }
            }
            get { return this.RunTime; }
        }

        private String SampleType
        {
            set
            {
                if (l == 0)
                {
                    this.SampleType = value;
                }
            }
            get { return this.SampleType; }
        }

        private Int32 Repeats
        {
            set
            {
                if (l == 0)
                {
                    this.Repeats = value;
                }
            }
            get { return this.Repeats; }
        }
        #endregion

        /* Constructors */
        public Sample(String name, String comment, String runTime, String sampleType, Int32 rpts)
        {
            l = 0;
            this.Name = name;
            this.Comment = comment;
            this.RunTime = runTime;
            this.SampleType = sampleType;
            this.Repeats = rpts;
            this.Elements = new List<Element>();
            l = 1;
        }

        public Sample(String name, String runTime, String sampleType, Int32 rpts)
        {
            l = 0;
            this.Name = name;
            this.Comment = "";
            this.RunTime = runTime;
            this.SampleType = sampleType;
            this.Repeats = rpts;
            this.Elements = new List<Element>();
            l = 1;
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
