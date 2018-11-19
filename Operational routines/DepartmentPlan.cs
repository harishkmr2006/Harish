using System;
using NUnit.Framework;

using HP.LFT.Verifications;
using System.Diagnostics;

using HP.LFT.SDK.WPF;
using HP.LFT.SDK;
using System.Drawing;
using HP.LFT.Report;
using BOPO.NUnit.ParallelTests;

namespace LeanFtTestProject1
{
    [TestFixture]
    public class DepartmentPlan : UnitTestClassBase
    {
        Process appProcess = new Process();
        ApplicationModel1 app = new ApplicationModel1();

        public string imagePath { get; private set; }

        public string getAppVersion()
        {
            string sClientExePath = @"C:\Program Files (x86)\H & M Hennes & Mauritz AB\H & M Department Plan Client\HM.Plan.DepartmentPlan.Client.UI.WPF.exe";
            var versionInfo = FileVersionInfo.GetVersionInfo(sClientExePath);
            var appver = versionInfo.FileVersion;
            return appver;
        }

        public void launchApplication()
        {
            try
            {
                appProcess = new Process { StartInfo = { FileName = @"C:\Program Files (x86)\H & M Hennes & Mauritz AB\H & M Department Plan Client\HM.Plan.DepartmentPlan.Client.UI.WPF.exe" } };
                appProcess.Start();
                

                waitSomeTime(3);
                if(app.DepartmentPlanWindow1.Exists())
                {
                    app.DepartmentPlanWindow1.CloseButton.Click();
                    app.DepartmentPlanWindow.Restore();
                    app.DepartmentPlanWindow.Maximize();
                    Reporter.ReportEvent("Department Plan launched successfuly", "Pass");
                }
                else
                {
                    app.DepartmentPlanWindow.GetVisibleText();
                }
                

                //if (app.AuthenticationErrorWindow.Exists())
                //{
                //    app.AuthenticationErrorWindow.OKButton.Click();
                //    Reporter.ReportEvent("Critical Jobs Running", app.AuthenticationErrorWindow.AuthenticationErrorUiObject.GetVisibleText());
                //    Reporter.ReportEvent("Test execution", Status.Failed.ToString());
                //    return 1;

                //}
                //if (app.QptExplorerWindow.Exists())
                //{

                //    writeLog("Launch Application", app.QptExplorerWindow.Text, "Pass");
                //    Console.WriteLine("Application Launched" + appProcess.ProcessName);
                //    return 0;
                //}
                ////else
                ////{
                //    checkException();
                //}
            }
            catch (Exception e)
            {
                if(app.DepartmentPlanWindow1.Exists())
                {
                    Reporter.ReportEvent("DepartmentPlan is down", app.DepartmentPlanWindow1.GetVisibleText());
                    app.DepartmentPlanWindow1.CloseButton.Click();
                   // Reporter.ReportEvent("Test execution", Status.Failed.ToString());
                    Reporter.ReportEvent("Application Error", "Department application error", Status.Failed);
                }
                
                Reporter.ReportEvent("Application Error", e.Message, Status.Failed, imagePath = Environment.CurrentDirectory);

                // Environment.Exit(0);

            }

        }
        public void waitSomeTime(double dTime)
        {
            dTime = dTime * 1000;
            System.Threading.Thread.Sleep(Convert.ToInt32(dTime));
        }

        public void selectableProductGrid()
        {
            try
            {
                app.DepartmentPlanWindow.DepartmentPlanDataTable.Highlight();
                var rowcount = app.DepartmentPlanWindow.DepartmentPlanDataTable.Rows.Count;

                var strloc = app.DepartmentPlanWindow.DepartmentPlanDataTable.GetTextLocations("DEFAULT PLAN");

                int x = Convert.ToInt32((strloc[0].Width) / 2.0 + strloc[0].X) + app.DepartmentPlanWindow.DepartmentPlanDataTable.Location.X;
                int y = Convert.ToInt32((strloc[0].Height) / 2.0 + strloc[0].Y) + app.DepartmentPlanWindow.DepartmentPlanDataTable.Location.Y;
                // Mouse.Click(new Point(x, y));
                Mouse.DoubleClick(new Point(x, y));
                waitSomeTime(20);
                Reporter.ReportEvent("Default plan is selected", "Pass");
            }
            catch(Exception e)
            {
                Reporter.ReportEvent("Not able to select Default plan", e.Message, Status.Failed, imagePath = Environment.CurrentDirectory);
            }
            //var txt = app.DepartmentPlanWindow.DepartmentPlanDataTable.GetVisibleText();
           // Console.WriteLine(txt);
        }

        public void createPlan()
        {
            try
            {
                app.DepartmentPlanWindow.NewButton.Click();
                waitSomeTime(10);
                String S = getUniqueName();
                app.NewPlanWindow.NewBudgetNameEditField.SendKeys(S);
                app.NewPlanWindow.CreateButton.Click();
                waitSomeTime(40);
                dataConfirmation();
                waituntilElementisVisible();
                app.DepartmentPlanWindow.DepartmentPlanDataTable.Highlight();
                var rowcount = app.DepartmentPlanWindow.DepartmentPlanDataTable.Rows.Count;
                var strloc = app.DepartmentPlanWindow.DepartmentPlanDataTable.GetTextLocations(S);
                int x = Convert.ToInt32((strloc[0].Width) / 2.0 + strloc[0].X) + app.DepartmentPlanWindow.DepartmentPlanDataTable.Location.X;
                int y = Convert.ToInt32((strloc[0].Height) / 2.0 + strloc[0].Y) + app.DepartmentPlanWindow.DepartmentPlanDataTable.Location.Y;
                Mouse.Click(new Point(x, y));
                app.DepartmentPlanWindow.DeleteButton.Click();
                app.DeletePlanWindow.YesButton.Click();
                waitSomeTime(15);
                Reporter.ReportEvent("Plan created successfuly", "Pass");
            }
            catch(Exception e)
            {
                Reporter.ReportEvent("Not able to create plan", e.Message, Status.Failed, imagePath = Environment.CurrentDirectory);
            }
                
        }
      
        public void dataConfirmation()
        {
            try
            {
                // app.EverydayCollectionWindow.GetDescription();
                Console.WriteLine(app.EverydayCollectionWindow1.GetDescription());
                app.EverydayCollectionWindow1.FileButton.Click();
                app.EverydayCollectionWindow1.FileButton.ExitUiObject.Click();
                Reporter.ReportEvent("Able to fetch data", "Pass");
            }
            catch(Exception e)
            {
                Reporter.ReportEvent("Data not available", e.Message, Status.Failed, imagePath = Environment.CurrentDirectory);
            }
        }

        public string getUniqueName()
        {
            return "Auto_" + getTimeStamp();
        }

        public string getTimeStamp()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }
        public void waituntilElementisVisible()
        {
            try
            {


                do
                {
                    waitSomeTime(10);
                    if (app.DepartmentPlanWindow.DepartmentPlanDataTable.Rows.Count > 0)
                    {
                        waitSomeTime(10);
                        break;
                    }
                } while (app.DepartmentPlanWindow.DepartmentPlanDataTable.Rows.Count >= 1);
               Reporter.ReportEvent("Versions loaded successfuly", "Pass");

            }
            catch(Exception e)
            {
                Reporter.ReportEvent("Error in loading versions", e.Message, Status.Failed, imagePath = Environment.CurrentDirectory);
            }
        }

        public void getplan()
        {
            var gridObj = app.DepartmentPlanWindow.DepartmentPlanDataTable;
            int rows = gridObj.Rows.Count;
            String gridText = gridObj.GetVisibleText();
            Reporter.ReportEvent("text:   ", gridText);
             


        }
        public void clickFile()
        {
            try
            {
               // app.DepartmentPlanWindow.FileButton.Highlight();
                app.DepartmentPlanWindow.FileButton.Click();
                app.DepartmentPlanWindow.FileButton.HelpFeedbackUiObject.Click();
                waitSomeTime(3);
                app.DepartmentPlanWindow.FileButton.Exists();
                app.DepartmentPlanWindow.FileButton.Click();
                //app.DepartmentPlanWindow.PlanUiObject.Click();
                Reporter.ReportEvent("Closed plan successfuly", "Pass");

            }
            catch(Exception e)
            {
                Reporter.ReportEvent("Error in closing plan", e.Message, Status.Failed, imagePath = Environment.CurrentDirectory);
            }

        }
        public void selectSection()
        {
            app.DepartmentPlanWindow.SectionsComboBox.Select("15 Everyday Collection");
            Reporter.ReportEvent("Successfuly selected Section", "Pass");
        }

        public void selectPLanLevel()
        {
            app.DepartmentPlanWindow.PlanLevelsListBoxList.Select("Plan");
            Reporter.ReportEvent("Successfuly selected Plan Level", "Pass");
        }
        
        public void selectSeason()
        {
            app.DepartmentPlanWindow.SeasonsListBoxList.Select("7-2018");
            Reporter.ReportEvent("Successfuly selected Season", "Pass");
        }

        public void clickGo()
        {
            app.DepartmentPlanWindow.GOButton.Click();
            Reporter.ReportEvent("Successfuly clicked on Go button", "Pass");
        }


        [OneTimeSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture
        }

        [SetUp]
        public void SetUp()
        {
            // Before each test
        }

       // [Test, Description("DepartmentPlan")]
        public void TC08_DepartmentPlanTest()
        {
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "DepartmentPlan");
            launchApplication();
            string qtpversion = getAppVersion();
            Reporter.ReportEvent("DepartmentPlan Version:  ", qtpversion);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), qtpversion);
            VersionConrol.addApplication("DepartmentPlan", qtpversion);
            clickFile();
            selectSection();
            selectPLanLevel();
            selectSeason();
            clickGo();
            waitSomeTime(20);
            selectableProductGrid();
            waitSomeTime(3);
            dataConfirmation();
            createPlan();
            Console.WriteLine("Success");
            waitSomeTime(10);
            app.DepartmentPlanWindow.Close();
        }

        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
        }

        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
        }
    }
}

