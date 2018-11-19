using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace SITSmokeTests
{
    public class ConfigUtils
    {
        public static String Read(string key)
        {
            var getValue = ConfigurationManager.AppSettings[key];
            return getValue;
        }
    }
}
