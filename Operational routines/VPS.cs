using System;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.Verifications;
using System.Threading;
using HP.LFT.SDK.WinForms;
using HP.LFT.Report;
using System.Diagnostics;
using BOPO.NUnit.ParallelTests;

namespace SIT_VPS
{
    [TestFixture]
    public class VPS : UnitTestClassBase
    {
        public string imagePath { get; private set; }
        ApplicationModel1 AppObj = new ApplicationModel1();
        [OneTimeSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture
        }

        [SetUp]
        public void SetUp()
        {
            // Before each test
            //VPSClassObject a = new VPS_ProductCreation();
           
        }

        public string getAppVersion()
        {
            string sClientExePath = "C:\\Program Files (x86)\\H & M Hennes & Mauritz AB\\H & M VPS\\HM.Vps.Client.UI.exe";
            var versionInfo = FileVersionInfo.GetVersionInfo(sClientExePath);
            var appver = versionInfo.FileVersion;
            return appver;
        }
        public void launchapp()
        {
            try
            {

                System.Diagnostics.Process.Start("C:\\Program Files (x86)\\H & M Hennes & Mauritz AB\\H & M VPS\\HM.Vps.Client.UI.exe");
                Thread.Sleep(5000);
                if (AppObj.VPSWindow1.OKUiObject.Exists())
                {
                    AppObj.VPSWindow1.OKUiObject.Click();
                    AppObj.VPSWindow.Restore();
                    AppObj.VPSWindow.Maximize();
                    Reporter.ReportEvent("VPS launched successfuly", "Pass");
                }
               else
                {
                    AppObj.VPSWindow.Maximize();
                    Reporter.ReportEvent("VPS launched successfuly", "Pass");
                }
               

            }
            catch (Exception e)
            {
                if(AppObj.SplashFormWindow.Exists())
                {
                    Reporter.ReportEvent("VPS is down", AppObj.SplashFormWindow.VPSErrorWindow.GetVisibleText());
                    AppObj.SplashFormWindow.VPSErrorWindow.CloseUiObject.Click();
                   // Reporter.ReportEvent("Test execution", Status.Failed.ToString());
                    Reporter.ReportEvent("Application Error", "VPS application error", Status.Failed);
                }
                Reporter.ReportEvent("Application Error", e.Message, Status.Failed, imagePath = Environment.CurrentDirectory);

            }

        }

      //  [Test, Description("VPS Flow")]
        public void TC10_VPS_Flow()
        {
            string SeasonToBeSelected = "5-2017";
            var sdept = "1111";
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "VPS");
            launchapp();
            string qtpversion = getAppVersion();
            Reporter.ReportEvent("VPS Version:  ", qtpversion);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), qtpversion);

            VersionConrol.addApplication("VPS", qtpversion);

            var defaultSeason = AppObj.VPSWindow.BoardcardExplorerWindow.PlanningSeason.GetVisibleText();

            if (SeasonToBeSelected == defaultSeason)
            {
                AppObj.VPSWindow.BoardcardExplorerWindow.PlanningSeason.Click();
                Keyboard.SendString(SeasonToBeSelected);
                Keyboard.SendString(SeasonToBeSelected);
            }
            AppObj.VPSWindow.BoardcardExplorerWindow.PlanningSeason.Click();
            Keyboard.SendString(SeasonToBeSelected);
            Reporter.ReportEvent("Season Selection", "User selected seaon as 5-2017", HP.LFT.Report.Status.Passed);
            Thread.Sleep(1000);
           // var sdept = "1111";
            var vNodes = AppObj.VPSWindow.BoardcardExplorerWindow.TreeDepartmentsList.NativeObject.Nodes[0].Nodes;
            var scount = vNodes.Count;
            // var scount = sNodemebers1.Nodes[0].Nodes.Count;
            Thread.Sleep(2000);
            bool bNodefound = false;

            for (int i = 0; i <= scount - 1; i++)
            {
               // sNodeName = appModel.HMOrderWindow.OrderExplorerWindow.UltraTreeDepartmentsUiObject.NativeObject.Nodes[0].Nodes[i].Key;
                var sNodeName = vNodes[i].Key;
                if (sNodeName == sdept)
                {
                    AppObj.VPSWindow.BoardcardExplorerWindow.TreeDepartmentsList.NativeObject.Nodes[0].Nodes[i].Selected = true;
                    //sNodemebers1.Nodes[0].Nodes[i].Selected = true;
                    bNodefound = true;
                    break;
                }
            }

            if (bNodefound == true)
            {
                Reporter.ReportEvent("VPS Department Node Selection", "1111 VPS Department Node Selected - ", HP.LFT.Report.Status.Passed);
            }
            else
            {
                Reporter.ReportEvent("VPS Department Node Selection", "1111 VPS Department Node Selected -", HP.LFT.Report.Status.Failed);
            }
            Thread.Sleep(16000);
            AppObj.VPSWindow.Close();

        }

        [TearDown]
        public void TearDown()
        {
            //String result =TestContext.CurrentContext.Result.Outcome.ToString();
            //VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.EndTime.ToString(), VersionConrol.getTimeStamp());
            //VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.TestResult.ToString(), result);

        }

        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
            
        }
    }
}
