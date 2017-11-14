using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WaterAnalysisTool.Exceptions
{
    [Serializable]
    class ConfigurationErrorException : Exception
    {
        public ConfigurationErrorException() { }

        public ConfigurationErrorException(String msg) : base(msg) { }

        public ConfigurationErrorException(String msg, Exception inner) : base(msg, inner) { }
    }
}
