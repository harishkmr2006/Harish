using NUnit.Framework;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOPO.NUnit.ParallelTests.Tests.TestRunup
{
    [TestFixture]
    public class TC07_PercentageAdminTool_FFBrowser:TestBase
    {
      //  [Test, Category("FFRegression1")]
        public Boolean TC07_PercentAdminToolVerificaton()
        {
            bool finalResult = false;
            Initial_Setup("TC07_PercentAdminToolVerificaton");
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "PercentPMAdminTool");
            try
            {
                System.Environment.SetEnvironmentVariable("TestObjective", "This test case is to verify Percent admin tool.");

                //-------------------Launch URL----------------------------------------
             //   Driver.Navigate().GoToUrl(ConfigUtils.Read("PERCENTADMINTOOL_URL"));

                //------------------Core Test Case-------------------------------------
                POM_PercentageAdminTool pHomepage = new POM_PercentageAdminTool(BrowserName, Driver, logStack);
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), pHomepage.versionInfo.GetAttribute("textContent"));
                VersionConrol._initTestData(pHomepage.versionInfo.GetAttribute("textContent"));
                Assert.IsTrue(pHomepage.PercentPMToolVerification(iExcel.ReadData(irownumber, "Season"), iExcel.ReadData(irownumber, "Brand"), iExcel.ReadData(irownumber, "SubIndex"), iExcel.ReadData(irownumber, "Department"), iExcel.ReadData(irownumber, "Version")));
                test.Log(LogStatus.Pass, "PercentPMToolVerification", "<font color = 'green'><strong>PASS</strong> </font>");

                //------------------------------ REPORTING-----------------------------------------------------------
                Generate_Report_And_Close_Browser("TC07_PercentAdminToolVerificaton");
                finalResult = true;
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.TestResult.ToString(), "Pass");
                return finalResult;

            }

            catch (Exception e)
            {
                // Handle_Exception("TC08_PercentAdminToolVerificaton", ex.StackTrace);
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

