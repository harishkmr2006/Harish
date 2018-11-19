using BOPO.NUnit.ParallelTests.POM.TestRunup;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOPO.NUnit.ParallelTests.Tests.TestRunup
{
    [TestFixture]
    //[Parallelizable]
    public class TC03_SellPriceTool_IEBrowser:TestBase
    {
        public Boolean TC03_SellPriceTool_IE()
        {
            bool finalResult = false;
            Initial_Setup("TC03_SellPriceTool_IE");
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "SellPriceTool");
            try
            {
                System.Environment.SetEnvironmentVariable("TestObjective", "Test cases to check on environment testrunup.");

                //--------------------------------TEST DATA GENERATOR-------------------------------------------------------------------------------------------

                // TestDataGenerator testgen = new TestDataGenerator(BrowserName, Driver, logStack, iExcel.ReadData(irownumber, "ProductificationBuildStatus"), iExcel.ReadData(irownumber, "Season"), iExcel.ReadData(irownumber, "SeasoninYYYYMM"), iExcel.ReadData(irownumber, "PlanMarketcode"));

                //---------------------------------------------------------------------------------------------------------

                POM_SellPriceToolPage Sptpage = new POM_SellPriceToolPage(BrowserName, Driver, logStack);
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), Sptpage.versionInfo.GetAttribute("textContent"));
                VersionConrol.addApplication("SellPriceTool", Sptpage.versionInfo.GetAttribute("textContent"));
                Sptpage.selectFiscalTab();
                Sptpage.selectingSeason(iExcel.ReadData(irownumber, "Season"));
                Sptpage.selectingBrand(iExcel.ReadData(irownumber, "Brand"));
                Sptpage.selectingChannel(iExcel.ReadData(irownumber, "Channel"));
                Sptpage.selectingFiscalCountry(iExcel.ReadData(irownumber, "Fiscal country"));
                Sptpage.enterSellPrice();
                Sptpage.saveSellPrice();
                Sptpage.confirmSaveSellPrice();
                Sptpage.deleteSellPrice();
                Sptpage.confirmDeleteSellPrice();
                Sptpage.saveSellPrice();
                Sptpage.confirmSaveSellPrice();
               



                //------------------------------ REPORTING-----------------------------------------------------------
                System.Environment.SetEnvironmentVariable("Endtime", DateTime.Now.ToString("yyyy:MM:dd_HH:mm:ss"));
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.EndTime.ToString(), VersionConrol.getTimeStamp());
                TestBase.sTestName = stestcaseid;
                ReportCustom swriteresult = new ReportCustom();
                swriteresult.WriteResults(stestcaseid, logStack);
                TestBase.sbrowsertype = System.Environment.GetEnvironmentVariable("Browser1");
                TestBase.DriverCleardown = Driver;

                //----------------------VIDEO RECORDING END--------------------------------
                //stopVideoRecording();
                //-------------------------------------------------------------------------
                finalResult = true;
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.TestResult.ToString(), "Pass");
                return finalResult;

            }

            catch (Exception e)
            {
                System.Environment.SetEnvironmentVariable("Endtime", DateTime.Now.ToString("yyyy:MM:dd_HH:mm:ss"));
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.EndTime.ToString(), VersionConrol.getTimeStamp());
                TestBase.sbrowsertype = System.Environment.GetEnvironmentVariable("Browser2");
                TestBase.sTestName = stestcaseid;
                ReportCustom swriteresult = new ReportCustom();
                swriteresult.WriteResults(stestcaseid, logStack, finalResult);
                Console.WriteLine("StackTrace from Test Class :" + e.StackTrace);
                Console.WriteLine("Message from Test Class" + e.Message);
                //stopVideoRecording();
                TestBase.DriverCleardown = Driver;
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.TestResult.ToString(), "Fail");
                // throw e;
                return finalResult;


            }
        }
    }
}
