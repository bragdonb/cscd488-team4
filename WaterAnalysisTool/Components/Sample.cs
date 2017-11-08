﻿using System;
using System.Collections.Generic;

namespace WaterAnalysisTool.Components
{
    class Sample
    {
        /* Attributes */
        private List<Element> elements;
        private String method;
        private String name;
        private String comment;
        private String runTime;
        private String sampleType;
        private int repeats;

        #region Properties
        public List<Element> Elements
        {
            get { return this.elements; }
        }

        public String Name
        {
            get { return this.name; }
        }

        public String Comment
        {
            get { return this.comment; }
        }

        public String RunTime
        {
            get { return this.runTime; }
        }

        public String SampleType
        {
            get { return this.sampleType; }
        }

        public Int32 Repeats
        {
            get { return this.repeats; }
        }

        public String Method
        {
            get { return this.method; }
        }
        #endregion

        /* Constructors */
        public Sample(String method, String name, String comment, String runTime, String sampleType, Int32 rpts)
        {
            this.method = method;
            this.name = name;
            this.comment = comment;
            this.runTime = runTime;
            this.sampleType = sampleType;
            this.repeats = rpts;
            this.elements = new List<Element>();
        }

        public Sample(String method, String name, String runTime, String sampleType, Int32 rpts)
        {
            this.method = method;
            this.name = name;
            this.comment = "";
            this.runTime = runTime;
            this.sampleType = sampleType;
            this.repeats = rpts;
            this.elements = new List<Element>();
        }

        /* Public functions */
        public void AddElement(Element e)
        {
            if (e == null)
                throw new ArgumentNullException("Element is Null.\n");

            this.Elements.Add(e);
        }
    }
}
