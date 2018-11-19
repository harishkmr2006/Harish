using System;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.Verifications;
using System.Threading;
using System.Diagnostics;
using System.IO;
using HP.LFT.SDK.StdWin;
using System.Drawing;
using System.Configuration;
using BOPO.NUnit.ParallelTests;
//using HP.LFT.SDK.StdWin;

namespace AssortmentPlan
{
    [TestFixture]
    public class AssortmentPlan1 : UnitTestClassBase
    {
        private AssortmentPlanModel assortmentPlanModel = new AssortmentPlanModel();
        [OneTimeSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture
        }

        [SetUp]
        public void SetUp()
        {
            // Before each test 
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());

        }

        //[Test]
        public void VerifyApplicationVersion()
        {
            Process assortmentPlanProcess = OpenApplication();

            ClickFileMenuItem();
            ClickSystemInformation();

            assortmentPlanModel.APWindow.AppVersionInfo.Highlight();

            string a = assortmentPlanModel.APWindow.AppVersionInfo.GetVisibleText();

            string[] VersionInfo = assortmentPlanModel.APWindow.AppVersionInfo.GetVisibleText().Split('\r');
            string AssortmentPlanVersion = VersionInfo[1].Replace('\n', ' ').Trim();
            string PlesVersion = VersionInfo[3].Replace('\n', ' ').Trim();

            assortmentPlanModel.APWindow.AppConnectionInfo.Highlight();

            a = assortmentPlanModel.APWindow.AppConnectionInfo.GetVisibleText();

            string[] ConnectionInfo = assortmentPlanModel.APWindow.AppConnectionInfo.GetVisibleText().Split('\r');
            string PlesSignalRAddress = ConnectionInfo[3].Replace('\n', ' ').Trim();
            string VDSSignalRAddress = ConnectionInfo[7].Replace('\n', ' ').Trim();
            string CirrusAddress = ConnectionInfo[9].Replace('\n', ' ').Trim();
            string GAMAddress = ConnectionInfo[11].Replace('\n', ' ').Trim();
            string PlanDatabase = ConnectionInfo[15].Replace('\n', ' ').Trim();
            string VDSDatabase = ConnectionInfo[17].Replace('\n', ' ').Trim();

            Reporter.ReportEvent("Version Info", "AssortmentPlanVersion: " + AssortmentPlanVersion, HP.LFT.Report.Status.Passed);
            Reporter.ReportEvent("Version Info", "PlesVersion: " + PlesVersion, HP.LFT.Report.Status.Passed);
            Reporter.ReportEvent("Version Info", "PlesSignalRAddress: " + PlesSignalRAddress, HP.LFT.Report.Status.Passed);
            Reporter.ReportEvent("Version Info", "VDSSignalRAddress: " + VDSSignalRAddress, HP.LFT.Report.Status.Passed);
            Reporter.ReportEvent("Version Info", "CirrusAddress: " + CirrusAddress, HP.LFT.Report.Status.Passed);
            Reporter.ReportEvent("Version Info", "GAMAddress: " + GAMAddress, HP.LFT.Report.Status.Passed);
            Reporter.ReportEvent("Version Info", "PlanDatabase: " + PlanDatabase, HP.LFT.Report.Status.Passed);
            Reporter.ReportEvent("Version Info", "VDSDatabase: " + VDSDatabase, HP.LFT.Report.Status.Passed);

            CloseApplication(assortmentPlanProcess);
        }

       // [Test]
        public void VerifyCirrus()
        {
            Process assortmentPlanProcess;

            DeleteImagesFolder();
            assortmentPlanProcess = OpenApplication();
            OpenImageBrowser();
            SelectSeasonAndDepartment();
            bool ImagesPresent = VerifyIfImagesAreLoaded();
            if (ImagesPresent)
                Reporter.ReportEvent("ImagesPresent", "Images Present", HP.LFT.Report.Status.Passed);
            else
                Reporter.ReportEvent("ImagesPresent", "No Images Present", HP.LFT.Report.Status.Failed);
            CloseApplication(assortmentPlanProcess);

        }

      //  [Test]
        public void CreateArticle()
        {
            Process assortmentPlanProcess;

            assortmentPlanProcess = OpenApplication();
            VersionConrol._initTestData("AssortmentPlan");
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "AssortmentPlan");

            string qtpversion = getAppVersion();
            Reporter.ReportEvent("Assortment Plan Vesion:  ", qtpversion);
            VersionConrol.addSubKey("AssortmentPlan", subKeys.keySet.Version.ToString(),qtpversion);
            VersionConrol.addApplication("Assortment Plan", qtpversion);

            assortmentPlanModel.APWindow.AddArticleButton.Highlight();
            assortmentPlanModel.APWindow.AddArticleButton.Click();
            Reporter.ReportEvent("Create Article", "Article Created Successfully", HP.LFT.Report.Status.Passed);
            CloseApplication(assortmentPlanProcess);
            VerifyCirrus();
        }

        private bool VerifyIfImagesAreLoaded()
        {
            return assortmentPlanModel.ImageBrowserWindow.WpfImage.Exists();
        }

        private void SelectSeasonAndDepartment()
        {
            string season = "6-2017";// ConfigurationManager.AppSettings["Season"];
            string department = "1322"; // ConfigurationManager.AppSettings["Department"];
            assortmentPlanModel.ImageBrowserWindow.SeasonsComboBox.Highlight();
            assortmentPlanModel.ImageBrowserWindow.SeasonsComboBox.Click();
            assortmentPlanModel.ImageBrowserWindow.SeasonsComboBox.Select(season);
            Thread.Sleep(5000);
            assortmentPlanModel.ImageBrowserWindow.DepartmentsComboBox.Click();
            assortmentPlanModel.ImageBrowserWindow.DepartmentsComboBox.DoubleClick();
            Keyboard.SendString(department);
            Thread.Sleep(10000);

            Reporter.ReportEvent("Select Season", "Selected Season: " + season, HP.LFT.Report.Status.Passed);
            Reporter.ReportEvent("Select Department", "Selected Department: " + department, HP.LFT.Report.Status.Passed);
        }

        private void OpenImageBrowser()
        {
            assortmentPlanModel.APWindow.OpenImageBrowserButton.Highlight();
            assortmentPlanModel.APWindow.OpenImageBrowserButton.Click();
            Thread.Sleep(10000);
            Reporter.ReportEvent("Open Image Browser", "Image browser opened successfully", HP.LFT.Report.Status.Passed);
        }

        private void DeleteImagesFolder()
        {
            string ImagesPath = Environment.GetEnvironmentVariable("LocalAppData") + "\\H&M\\Images";

            try
            {
                if (Directory.Exists(ImagesPath))
                {
                    Directory.Delete(ImagesPath, true);
                    Reporter.ReportEvent("DeleteImagesFolder", "Deleted Folder: " + ImagesPath, HP.LFT.Report.Status.Passed);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.Message);
            }

        }

        private void ClickFileMenuItem()
        {
            var a = assortmentPlanModel.APWindow.RadRibbonViewTabStrip.GetTextLocations("Build");
            //int x = Convert.ToInt32((srect[0].Width) / 2.0 + srect[0].X) + app.MainQuantificationWindowID.DataHeaderTableUiObject.Location.X;
            //int y = Convert.ToInt32((srect[0].Height) / 2.0 + srect[0].Y) + app.MainQuantificationWindowID.DataHeaderTableUiObject.Location.Y;
            //Mouse.Click(new Point(x, y));
            Reporter.ReportEvent("Click File Menu", "File menu Clicked", HP.LFT.Report.Status.Passed);
        }

        private void CloseApplication(Process assortmentPlanProcess)
        {
            assortmentPlanProcess.CloseMainWindow();
            Reporter.ReportEvent("Application Close", "Application closed successfully", HP.LFT.Report.Status.Passed);
        }

        private Process OpenApplication()
        {

           // string AppPath = ConfigurationManager.AppSettings["ApplicationPath"];
          string AppPath = "C:\\Program Files (x86)\\H & M Hennes & Mauritz AB\\H & M Assortment Plan Client\\HM.Plan.AssortmentPlan.exe";
            Process p = Process.Start(AppPath);
            assortmentPlanModel.APWindow.Exists(10);
            Reporter.ReportEvent("Open Application", "Application opened successfully", HP.LFT.Report.Status.Passed);
            return p;
        }

        public string getAppVersion()
        {
            string sClientExePath = "C:\\Program Files (x86)\\H & M Hennes & Mauritz AB\\H & M Assortment Plan Client\\HM.Plan.AssortmentPlan.exe";
            var versionInfo = FileVersionInfo.GetVersionInfo(sClientExePath);
            var appver = versionInfo.FileVersion;
            return appver;
        }

        private void ClickSystemInformation()
        {
            assortmentPlanModel.APWindow.RibbonItemSystemInformation.Highlight();
            assortmentPlanModel.APWindow.RibbonItemSystemInformation.Click();
            Reporter.ReportEvent("Click System Information", "Clicked System Information successfully", HP.LFT.Report.Status.Passed);
        }

        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
            string result = TestContext.CurrentContext.Result.Outcome.ToString();
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.EndTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addTestDataValue(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.TestResult.ToString(), result);
        }

        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
        }
    }
}
