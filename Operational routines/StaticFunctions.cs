using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOPO.NUnit.ParallelTests
{
    public static class StaticFunctions
    {
         public static string TimeStampReader()
        {
            DateTime now = DateTime.Now;
            var timestamp = now.ToString("yyyyMMddHHmmss");
            return timestamp;
        }

       

    }


}
