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
    public class TC05_AllocatedAssortmentTool_FFBrowser:TestBase
    {
        
           public Boolean TC05_AllocatedAssortmentTool_FF()
            {
                bool finalResult = false;
                Initial_Setup("TC05_AllocatedAssortmentTool_FF");
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "AllocatedAssortmentTool");
            try
            {
                System.Environment.SetEnvironmentVariable("TestObjective", "Test cases to check on environment testrunup.");

                //--------------------------------TEST DATA GENERATOR-------------------------------------------------------------------------------------------

                // TestDataGenerator testgen = new TestDataGenerator(BrowserName, Driver, logStack, iExcel.ReadData(irownumber, "ProductificationBuildStatus"), iExcel.ReadData(irownumber, "Season"), iExcel.ReadData(irownumber, "SeasoninYYYYMM"), iExcel.ReadData(irownumber, "PlanMarketcode"));

                //---------------------------------------------------------------------------------------------------------

                POM_AllocatedAssortmentToolPage aatpage = new POM_AllocatedAssortmentToolPage(BrowserName, Driver, logStack);
                aatpage.showBuildinfo();
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), aatpage.checkVersion.GetAttribute("textContent"));
                VersionConrol.addApplication("AllocatedAssortmentTool", aatpage.checkVersion.Text);
                aatpage.pickanyindex();
                aatpage.enterDataonpub();
                aatpage.republish();
                aatpage.clearCache();
                aatpage.selectingBrand(iExcel.ReadData(irownumber, "Brand"));
                aatpage.selectingSeason(iExcel.ReadData(irownumber, "Season"));
                aatpage.selectingPM(iExcel.ReadData(irownumber, "Planning Market"));
                aatpage.selectingIndex(iExcel.ReadData(irownumber, "Index"));
                aatpage.nonpublished();
                aatpage.clickCheckbox();
                aatpage.enterSize();
                aatpage.enterData();
                aatpage.clickSave();
                aatpage.clearCache();
                aatpage.selectingBrand(iExcel.ReadData(irownumber, "Brand"));
                aatpage.selectingSeason(iExcel.ReadData(irownumber, "Season"));
                aatpage.selectingPM(iExcel.ReadData(irownumber, "Planning Market"));
                aatpage.selectingIndex(iExcel.ReadData(irownumber, "Index"));
                aatpage.picknonpublished();
                aatpage.clearSize();
                aatpage.clearData();
                aatpage.clickSave();
                aatpage.clearCache();




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


