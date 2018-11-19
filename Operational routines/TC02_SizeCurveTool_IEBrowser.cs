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
   public class TC02_SizeCurveTool_IEBrowser : TestBase
    {
        
        public Boolean TC02_SizeCurveTool_IE()
        {
            bool finalResult = false;
            Initial_Setup("TC02_SizeCurveTool_IE");
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "SizeCurveTool");
            try
            {
                System.Environment.SetEnvironmentVariable("TestObjective", "Test cases to check on environment testrunup.");

                //--------------------------------TEST DATA GENERATOR-------------------------------------------------------------------------------------------

                // TestDataGenerator testgen = new TestDataGenerator(BrowserName, Driver, logStack, iExcel.ReadData(irownumber, "ProductificationBuildStatus"), iExcel.ReadData(irownumber, "Season"), iExcel.ReadData(irownumber, "SeasoninYYYYMM"), iExcel.ReadData(irownumber, "PlanMarketcode"));

                //---------------------------------------------------------------------------------------------------------

                POM_SizeCurveToolPage Sctpage = new POM_SizeCurveToolPage(BrowserName, Driver, logStack);
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), Sctpage.checkVersion.GetAttribute("textContent"));
                Sctpage.showBuildinfo();
                VersionConrol.addApplication("SizeCurveTool", Sctpage.checkVersion.GetAttribute("textContent"));
                Sctpage.selectingSeason(iExcel.ReadData(irownumber, "Season"));
                Sctpage.selectingSection(iExcel.ReadData(irownumber,"Section"));
                Sctpage.selectingDepartment(iExcel.ReadData(irownumber, "Department"));
                Sctpage.createNewComparision();
                Sctpage.clearComparision();
                Sctpage.enteringComparisionname();
                Sctpage.selectSizeScale(iExcel.ReadData(irownumber, "SizeScale"));
                Sctpage.createComparision();
                Sctpage.createNewCurvegrp();
                Sctpage.clickCreateNewCurvegrp();
                Sctpage.clickEditbutton();
                Sctpage.productCheck();
                Sctpage.clickAnalysisButton();
                Sctpage.saveCurveGroup();
                Sctpage.closeComp();
                Sctpage.deletecomp();
                

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


