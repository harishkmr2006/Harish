 using System;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.SDK.WPF;
using System.Linq;
using System.Diagnostics;
using HP.LFT.Report;
using HP.LFT.Verifications;
using System.Drawing;
//using CustomResport = BOPO.NUnit.ParallelTests.Reporter;
using setuppro = BOPO.NUnit.ParallelTests.TestBase;
using OpenQA.Selenium;
using BOPO.NUnit.ParallelTests;
using System.Collections.Generic;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Remote;
using initsetup = BOPO.NUnit.ParallelTests.TestBase;

namespace QPT
{
    [TestFixture]
    public class QPT1 : UnitTestClassBase
    {
        Process appProcess = new Process();
        private string imagePath;
        IWebDriver mainDriver = null;   



        //DesiredCapabilities capabilities = new DesiredCapabilities("Firefox", "45.8.0", new Platform(PlatformType.Windows));
        //capabilities.SetCapability("username", "TEMPHAHUC");
        //    capabilities.SetCapability("accessKey", "Citrix1234");
        //    capabilities.SetCapability(FirefoxDriver.BinaryCapabilityName, "C:\\Program Files (x86)\\Mozilla Firefox ESR v45.8.0\\firefox.exe");
        //    capabilities.SetCapability("network.automatic-ntlm-auth.trusted-uris", Environment.UserDomainName);
        //    capabilities.SetCapability("signon.autologin.proxy", true);
        //    IWebDriver driver = new FirefoxDriver(capabilities);


        List<Result> stestcase=null;

        ApplicationModel app = new ApplicationModel();
        public string sSeason = "7-2018";
        public string sDepartment = "1111 - Eyes";
        public IMenuItem[] qt { get; private set; }
        public object Checked { get; private set; }
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Function to launch the application -
        public int launchApplication()
        {
            IWebDriver driver = mainDriver;
            List<Result> localtestcase = stestcase;
            try
            {
                appProcess = new Process { StartInfo = { FileName = @"C:\Program Files (x86)\H & M Hennes & Mauritz AB\H & M QPT Client\HM.Plan.QPT.Client.UI.WPF.exe" } };
                appProcess.Start();
          
                waitSomeTime(3);
               
                           
                if (app.AuthenticationErrorWindow.Exists())
                {
                    app.AuthenticationErrorWindow.OKButton.Click();
                    Reporter.ReportEvent("Critical Jobs Running",app.AuthenticationErrorWindow.AuthenticationErrorUiObject.GetVisibleText());
                    Reporter.ReportEvent("Test execution", Status.Failed.ToString());
                   // CustomResport.Report(localtestcase, driver, "Exception in launching QTP: " ,"");
                    return 1;
                    
                }
                else if (app.QptExplorerWindow.Exists())
                {
                    
                    writeLog("Launch Application", app.QptExplorerWindow.Text, "Pass");
                    Console.WriteLine("Application Launched" + appProcess.ProcessName);
                   // CustomResport.Report(localtestcase, driver, "Successfuly launched QTP", "Successfuly launched QTP");
                    return 0;
                }
                else
                {
                    app.QptExplorerWindow.GetVisibleText();
                }
            }
            catch (Exception e)
            {
                if (app.QPTErrorWindow.Exists())
                {

                    Reporter.ReportEvent("QPT is down", app.QPTErrorWindow.GetVisibleText());
                    app.QPTErrorWindow.CloseButton.Click();
                    Reporter.ReportEvent("Application Error", "QPT application error", Status.Failed);
                   // Reporter.ReportEvent("Test execution", Status.Failed.ToString());
                    //Environment.Exit(0);
                    // return 1;
                }
                Reporter.ReportEvent("Application Error", e.Message, Status.Failed, imagePath = Environment.CurrentDirectory);
               // CustomResport.Report(localtestcase, driver, "Exception in launching QTP: ", "");

                Environment.Exit(0);

            }
            return 0;
        }
        //-----------------------------------------------------------------------------------------------------------------------------------------------------------------
        //get the qpt application version
        public string getAppVersion()
        {
            string sClientExePath = "C:\\Program Files (x86)\\H & M Hennes & Mauritz AB\\H & M QPT Client\\HM.Plan.QPT.Client.UI.WPF.exe";
            var versionInfo = FileVersionInfo.GetVersionInfo(sClientExePath);
            var appver = versionInfo.FileVersion;
            return appver;
        }
        //----------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Generic wait function
        public void waitSomeTime(double dTime)
        {
            dTime = dTime * 1000;
            System.Threading.Thread.Sleep(Convert.ToInt32(dTime));
        }
        //Report.
        public void writeLog(string stepName, string Description, string state)
        {

            if (state.ToUpper() == "PASS")
            {
                Reporter.ReportEvent(stepName, Description, Status.Passed);
            }
            else
            {
                Reporter.ReportEvent(stepName, Description, Status.Failed);
            }

        }
        //Exception handling
        public void checkException()
        {
            IWebDriver driver = mainDriver;
            List<Result> localtestcase = stestcase;
            try
            {
                if (app.QPTErrorWindow.Exists())
                {
                    Reporter.ReportEvent("Error occured", Status.Failed.ToString());
                    app.QPTErrorWindow.Close();
                    CustomResport.Report(localtestcase, driver, "Exception in launching QTP: ", "");
                    //  Environment.Exit(0);
                    //execution will 
                }
            }
            catch
            {
                throw new Exception("QTP Exception Occured");
               // Reporter.ReportEvent("Error ", e.Message);
            }
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Method click on RadioButton
        public void clickRadioButton(string RadioButtonName)
        {
            //IWebDriver driver = mainDriver;
            //List<Result> localtestcase = stestcase;
            switch (RadioButtonName.ToUpper())
            {
                case "STORE":
                    app.QptExplorerWindow.StoreRadioButton.Click();
                     break;
                case "ONLINE":
                    app.QptExplorerWindow.OnlineRadioButton.Click();
                    break;
                case "ALL":
                    app.QptExplorerWindow.AllRadioButton.Click();
                    break;
            }
            writeLog("Click RadioButton: ", RadioButtonName, "Pass");
           // CustomResport.Report(localtestcase, driver, "Successfuly clicked Radio button", "Successfuly clicked Radio button");
        }
     
        //Select th Department
        public void selectDepartment(string vDepartmentString)
        {
            try
            {
                int offset = 0;
                string[] nodetextpath = vDepartmentString.Split(';');
                for (int i = 0; i <= nodetextpath.GetUpperBound(0); i++)
                {
                    if (i == nodetextpath.GetUpperBound(0))
                    {
                        app.QptExplorerWindow.TreeView.Select(vDepartmentString);
                        writeLog("Select ", "Department:" + vDepartmentString, "Pass");
                        break;
                    }
                    string nodetoExpand = vDepartmentString.Substring(0, vDepartmentString.IndexOf(';', offset));
                    offset = vDepartmentString.IndexOf(';') + 1;
                    app.QptExplorerWindow.TreeView.GetNode(nodetoExpand).Expand();
                    waitSomeTime(3);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Returns the current Timestamp
        public string getTimeStamp()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //Generates random unique Name 
        //Eg:- Auto_10012017010101
        public string getUniqueName()
        {
            return "Auto_" + getTimeStamp();
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //select selectgrid and check child checkbox
        public void selectableProductGridChild()
        {
            try
            {
            int numberOfIterations = 3;
            var selectProductGrid = app.MainQuantificationWindowID.SelectableProductsDataGridTable;
            for (int l = 0; l < numberOfIterations; l++)
            {
                if (selectProductGrid.Rows.Count > 0) { break; }
                waitSomeTime(3);
            }

            int totalrowCount = selectProductGrid.Rows.Count;
            string gridRowText = selectProductGrid.GetVisibleText();
            string[] delimitgridRowText = gridRowText.Split('►');

            int selectProductNumber = 0273777;

            int rowCellCount = selectProductGrid.Rows[0].Cells.Count;
            bool flag = false;
            for (int m = 0; m < totalrowCount; m++)
            {

                for (int n = 0; n < rowCellCount; n++)
                {
                    int producNumber = Convert.ToInt32(selectProductGrid.Rows[m].Cells[1].Value);

                    if (selectProductNumber == producNumber)
                    {
                            app.MainQuantificationWindowID.SelectableProductsDataGridTable.Describe<IButton>(new ButtonDescription
                            {
                                FullType = @"button",
                                Index = Convert.ToUInt16(m)
                            }).Highlight();
                            selectProductGrid.Describe<IButton>(new ButtonDescription
                            {
                                FullType = @"button",
                                Index = Convert.ToUInt16(m)
                            }).SendKeys(HP.LFT.SDK.Keys.Return);
                            Reporter.ReportEvent("Selected Product Number       ", selectProductNumber.ToString());
                        flag = true;
                        break;
                    }

                }
                if (flag) break;
            }
            var chd = selectProductGrid.FindChildren<ICheckBox>(new CheckBoxDescription
            {
                FullType = @"check box",
                NativeClass = @"System.Windows.Controls.CheckBox",

            });
            long childcount = Convert.ToUInt16(chd.LongCount());
            Console.WriteLine("Child count = " + childcount);
            for (int j = 0; j < childcount; j++)
            {
                  Reporter.ReportEvent("Color Code:     ", chd[j].Text.ToString());
                  chd[j].Click();
                  Reporter.ReportEvent("Color Code:     ", chd[j].Text.ToString());
                }
        }catch(Exception e )
            {
              
                var selectProductGrid = app.MainQuantificationWindowID.SelectableProductsDataGridTable;
                var chd = selectProductGrid.FindChildren<ICheckBox>(new CheckBoxDescription
                {
                    FullType = @"check box",
                    NativeClass = @"System.Windows.Controls.CheckBox",

                });
                long childcount = Convert.ToUInt16(chd.LongCount());
                Console.WriteLine("Child count = " + childcount);
                for (int j = 0; j < childcount; j++)
                {
                  chd[j].Click();
                }
                Reporter.ReportEvent("Error Message", e.Message);
            }
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------------------------
       
        //[Test,Description("QPT Test")]
        public void TC09_QPTFlow()
        {
            //SDK.Init(new SdkConfiguration());
            //Reporter.Init(new ReportConfiguration

            //{
            //    Title = "Execution Report",
            //    Description = "Feasibility Test Results",
            //    IsOverrideExisting = true,
            //    ReportFolder = "TestResults",
            //    ReportLevel = ReportLevel.All,
            //    SnapshotsLevel = CaptureLevel.All,

            //});
            

           
            try
            {
                VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "QPT");
                var applaunchflag = launchApplication();
                 if (applaunchflag != 1)
                {
                    string qtpversion = getAppVersion();
                    Reporter.ReportEvent("QTP Version:  ", qtpversion);
                    VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), qtpversion);
                    VersionConrol.addApplication("QPT", qtpversion);


                    clickRadioButton("store");
                    string sSeasonDepartment = sSeason + ";" + sDepartment;
                    selectDepartment(sSeasonDepartment);
                    app.QptExplorerWindow.NewButton.Click();
                    string testName = getUniqueName();
                    app.NewFolderWindow.NameEditField.SetText(testName);
                    app.NewFolderWindow.FolderDescription.SetText("Creating Folder for Automation Testing purpose");
                    app.NewFolderWindow.Save.Click();
                    waitSomeTime(3);
                    string folderName = sSeasonDepartment + ";" + testName;
                    selectDepartment(folderName);

                    var men = app.QptExplorerWindow.NewMenu.BuildMenuPath("New;Store One.*");
                    waitSomeTime(2);
                    var oMenu = app.QptExplorerWindow.QuantificationNewButton;
                    oMenu.Click();
                    var IDProp = As.RegExp(@".*Two Seasons.*");
                    var sStoreOneQuantification = oMenu.Describe<IUiObject>(new UiObjectDescription { ObjectName = "Store One Season Quantification", FullNamePath = "Store One Season Quantification;Store One Season Quantification;PART_SubMenuScrollViewer;" });
              
                    var sStoreTwoQuantification = oMenu.Describe<IUiObject>(new UiObjectDescription { ObjectName = IDProp, FullNamePath = IDProp });
              
                    sStoreOneQuantification.Click();
              
                    app.NewQuantificationWindow.CreateButton.Click();
                    waitSomeTime(3);
                    app.MainQuantificationWindowID.Click();
                     if (app.MainQuantificationWindowID.Exists())
                    {
                        app.MainQuantificationWindowID.Level1TabControlTabStrip.Click();
                        int tabcount = app.MainQuantificationWindowID.Level1TabControlTabStrip.Tabs.Count;

                        Console.WriteLine("tabcount: " + tabcount);
                        app.MainQuantificationWindowID.SelectProductButton.Click();
                        try
                        {

                            if (app.QPTErrorWindow.Exists())
                            {
                                Reporter.ReportEvent("Error Occured", " List is empty", Status.Failed);
                               Environment.Exit(0);
                                app.QPTErrorWindow.Close();
                                app.MainQuantificationWindowID.Close();
                                app.QptExplorerWindow.Close();
                                
                            }
                        }
                        catch (Exception e)
                        {
                            Reporter.ReportEvent("Exceptopn", e.Message);
                            
                        }
                        try
                        {
                       var itemscount =  app.MainQuantificationWindowID.DepartmentsListBoxComboBox.Items[0].Count();
                            if(itemscount==1)
                            {
                                Reporter.ReportEvent("Selected Department in product", " List is empty", Status.Failed);
                               
                            }

                        }catch (Exception e)
                        {
                            Reporter.ReportEvent("Exceptopn", e.Message);
                        }
                        waitSomeTime(5);
                        selectableProductGridChild();
                       
                        app.MainQuantificationWindowID.QuantifyButton.Click();
                        waitSomeTime(10);
                       // app.MainQuantificationWindowID.SaveButton.Highlight();
                        app.MainQuantificationWindowID.SaveButton.Click();
                        waitSomeTime(10);
                        app.MainQuantificationWindowID.Level2TabControlTabStrip.Select("SE-01");
                        app.MainQuantificationWindowID.TimeHeaderTableUiObject.Highlight();
                        app.MainQuantificationWindowID.Close();

                        Reporter.ReportEvent("Test Name selected:  ", testName);
                        app.QptExplorerWindow.DeleteButton.Click();
                        app.DeleteQuantificationFolderWindow.YesButton.Click();
                        Reporter.ReportEvent(testName, " Deleted Successfully");
                        app.QptExplorerWindow.Close();
                        Reporter.ReportEvent("QTP application  ", "closed successfully", Status.Passed);
                        // Reporter.GenerateReport();
                        // SDK.Cleanup();

                    }
                    else
                    {
                        return;

                    }
                }
            }catch(HP.LFT.SDK.GeneralReplayException )
            {
                Reporter.ReportEvent("Exceptopn", Status.Failed.ToString());
            }
        }
        
        [TearDown]
        public void TearDown()
        {
           
        }

        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            
        }
    }

    
}

