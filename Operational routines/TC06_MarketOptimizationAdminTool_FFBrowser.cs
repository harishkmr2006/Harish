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
    public class TC06_MarketOptimizationAdminTool_FFBrowser:TestBase
    {
        public Boolean TC06_MarketOptimizationAdminTool_FF()
        {
            bool finalResult = false;
            Initial_Setup("TC06_MarketOptimizationAdminTool_FF");
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "MarketOptAdminTool");
            try
            {
                System.Environment.SetEnvironmentVariable("TestObjective", "Test cases to check on environment testrunup.");

                //--------------------------------TEST DATA GENERATOR-------------------------------------------------------------------------------------------

                // TestDataGenerator testgen = new TestDataGenerator(BrowserName, Driver, logStack, iExcel.ReadData(irownumber, "ProductificationBuildStatus"), iExcel.ReadData(irownumber, "Season"), iExcel.ReadData(irownumber, "SeasoninYYYYMM"), iExcel.ReadData(irownumber, "PlanMarketcode"));

                //---------------------------------------------------------------------------------------------------------

                POM_MarketOptimizationAdminToolPage moatpage = new POM_MarketOptimizationAdminToolPage(BrowserName, Driver, logStack);
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), moatpage.versionInfo.GetAttribute("textContent"));
                VersionConrol.addApplication("MarketOptimizationAdminTool",moatpage.versionInfo.GetAttribute("textContent"));
                moatpage.gettabnames();
                moatpage.clickManagetab();
                moatpage.clickAddSection();
                moatpage.clickDDButton();
                moatpage.checkDivSection();
                moatpage.clickDDButton();
                moatpage.clickOKButton();
                moatpage.clickSaveUpdateButton();
                moatpage.clickSPPperTab();
                moatpage.checkDIVSECinSPPPPER();
                moatpage.clickBuyperTab();
                moatpage.checkDIVSECinBUYPER();

                




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
                return finalResult;


            }
        }
    }
}
