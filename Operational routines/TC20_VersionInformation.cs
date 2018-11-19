using BOPO.NUnit.ParallelTests.POM.TestRunup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOPO.NUnit.ParallelTests.Tests.TestRunup
{
    public class TC20_VersionInformation : TestBase
    {
        public Boolean TC20_VersionInformationTest()
        {
            bool finalResult = false;
            Initial_Setup("TC20_Version_InformationTest");

            POM_SizeCurveToolPage Sctpage = new POM_SizeCurveToolPage(BrowserName, Driver, logStack);
           
            Sctpage.printAllVersionInfo();
            Driver.Close();
            return true;
        }
    }
}
