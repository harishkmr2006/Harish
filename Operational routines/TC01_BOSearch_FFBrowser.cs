using BOPO.NUnit.ParallelTests.Helpers;
using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace BOPO.NUnit.ParallelTests
{
    [TestFixture]
    //[Parallelizable]
 public   class TC01_BOSearch_FFBrowser : TestBase
    {
       public Boolean TC01_BOSearch_FF()
        {
            bool finalResult = false;
            Initial_Setup("TC01_BOSearch_FF");
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "BOSearch");
            try
            {
                System.Environment.SetEnvironmentVariable("TestObjective", "Test cases to check on environment testrunup.");

                //--------------------------------TEST DATA GENERATOR-------------------------------------------------------------------------------------------

               // TestDataGenerator testgen = new TestDataGenerator(BrowserName, Driver, logStack, iExcel.ReadData(irownumber, "ProductificationBuildStatus"), iExcel.ReadData(irownumber, "Season"), iExcel.ReadData(irownumber, "SeasoninYYYYMM"), iExcel.ReadData(irownumber, "PlanMarketcode"));

                //---------------------------------------------------------------------------------------------------------

                POM_BOSearchPage Hbosearchpage = new POM_BOSearchPage(BrowserName, Driver, logStack);
               // VersionConrol._initTestData(TestContext.CurrentContext.Test.Name);
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), Hbosearchpage.versionInfo.GetAttribute("textContent"));
                VersionConrol.addApplication("BOSearch", Hbosearchpage.versionInfo.GetAttribute("textContent"));
                Hbosearchpage.selectDept(iExcel.ReadData(irownumber, "Department"));
              //  string str = Hbosearchpage.versionInfo.Text;
                //string sdfghjk = Hbosearchpage.versionInfo.GetAttribute("textContent");

                
                Hbosearchpage.clickSearch();
                Hbosearchpage.getArticleDetails();
                Hbosearchpage.productClick();


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
