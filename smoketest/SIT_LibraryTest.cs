//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace SITSmokeTests
//{
//    class SIT_LibraryTest
//    {
//    }
//}


using System;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.Verifications;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System.Threading;
using System.Linq;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using System.Configuration;
using InteropLibrary;
using OpenQA.Selenium.Interactions;
using System.Drawing.Imaging;

namespace SITSmokeTests
{
    class SIT_LibraryTest : UnitTestClassBase
    {
        IWebDriver driver;
        WebDriverWait wait;
        ExcelHelper xCellFileHelper;
        string datafilePath;
        private SIT_Library_UI libraryUi;
        private ICCBAMPageObj iccPortal;

        public SIT_LibraryTest()
        {
            var binary = new FirefoxBinary(ConfigUtils.Read("FirefoxPath"));
            // string path = ReadFirefoxProfile();
            FirefoxProfile ffprofile = new FirefoxProfile();
            driver = new FirefoxDriver(binary, ffprofile);
            // driver.Manage().Cookies.DeleteAllCookies();
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            Thread.Sleep(1000);
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);
            // Before each test
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            datafilePath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + ConfigUtils.Read("TestDataPath");
            libraryUi = new SIT_Library_UI(driver);
            iccPortal = new ICCBAMPageObj(driver);
        }
        

        [OneTimeSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture
        }

        [SetUp]
        public void SetUp()
        {
            //var binary = new FirefoxBinary(ConfigUtils.Read("FirefoxPath"));
            //// string path = ReadFirefoxProfile();
            //FirefoxProfile ffprofile = new FirefoxProfile();
            //driver = new FirefoxDriver(binary, ffprofile);
            //// driver.Manage().Cookies.DeleteAllCookies();
            //driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            //driver.Manage().Window.Maximize();
            //Thread.Sleep(1000);
            // Before each test
            //  options.BrowserExecutableLocation = ConfigUtils.Read("FirefoxPath");
            // driver = new FirefoxDriver(options);
            Thread.Sleep(1000);
            var binary = new FirefoxBinary(ConfigUtils.Read("FirefoxPath"));
            // // string path = ReadFirefoxProfile();
            FirefoxProfile ffprofile = new FirefoxProfile(ConfigUtils.Read("FirefoxPath"));
            ffprofile.SetPreference("binary", ConfigUtils.Read("FirefoxPath"));
            driver = new FirefoxDriver(binary, ffprofile);
            //  FirefoxOptions sfopt = new FirefoxOptions();
            //sfopt.Profile = ffprofile;
            // sfopt.AddAdditionalCapability("binary", binary);
            // sfopt.SetPreference("binary", ConfigUtils.Read("FirefoxPath"));
            // //  sfopt.BrowserExecutableLocation = ConfigUtils.Read("FirefoxPath");
            // // sfopt.ToCapabilities.binary=ConfigUtils.Read("FirefoxPath");
            // driver = new FirefoxDriver(sfopt);
            // driver.Manage().Cookies.DeleteAllCookies();
            // driver.Manage().Timeouts().ImplicitWait = (TimeSpan.FromSeconds(15));
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            //var x = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width;
            //var y = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height;

            //driver.Manage().Window.Size = new System.Drawing.Size(x,y);
            //Thread.Sleep(1000);
            //// Before each test
            driver.Manage().Window.Maximize();
           
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            datafilePath = ConfigUtils.Read("TestDataPath");
            libraryUi = new SIT_Library_UI(driver);
            iccPortal = new ICCBAMPageObj(driver);
        }

      //  [Test]

        public void TC26_SIT_Library_License_Companies()
        {

            try
            {
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(10000);
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                string username = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAUSER");//"admtempjavas";
                string password = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAPWD");//"admtempjavas";
                string season = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");// "7-2018";
                Thread.Sleep(1000);
               
                driver.Manage().Window.Maximize();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                List<string> lswins = driver.WindowHandles.ToList();
                sGMethods.GetLatestWindow(driver);
                Castorpages castorobjs = new Castorpages(driver);
                //castorobjs.CastorLogin(
                //xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "librarianUser"),
                //xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "librarianPassword"));

                castorobjs.CastorLogin(username, password);
                Thread.Sleep(20000);
                //  ********************************************  License Companies ****************************************************************************
                libraryUi.LaunchLicenseCompanies();
                libraryUi.traverseContentFrame();
                libraryUi.traverseToContentBodyFrame();
                string companyName = null;
                if (libraryUi.get_LicenseCompaniesTableRowRecord().Count > 0)
                {
                    libraryUi.get_Checkbox(Int32.Parse(libraryUi.get_LicenseCompaniesTableRowRecord()[1].GetAttribute("rowIndex"))).Click();
                    companyName = libraryUi.get_LicenseCompaniesTableRowRecordName()[1].Text;
                    libraryUi.traverseContentFrame();
                    libraryUi.traverseToContentBodyFrameTableSettings();
                    libraryUi.get_Activate().Click();
                    libraryUi.traverseContentFrame();
                    libraryUi.traverseToContentBodyFrame();
                    for (int i = 0; i < 20; i++)
                    {
                        if (libraryUi.get_StatusByCompanyName(companyName).Text.Equals("Active"))
                        {

                            break;
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine(libraryUi.get_StatusByCompanyName(companyName).Text); Thread.Sleep(1000);
                        }
                    }
                }
                else
                {
                    // no inactive compnay
                }

                //check export
                string libraryWindow = libraryUi.get_libraryWindow();
                libraryUi.get_NewLaunchedWindow();

                iccPortal.LaunchICCWindow();
                               
                try {
                    iccPortal.SelectValueFromApplicationDropDown("LicenceCompany_Castor");
                    iccPortal.VerifySearchResult("LicenceCompany_Castor");
                    iccPortal.VerifyPortInSearchResult("OFS");
                    iccPortal.VerifyPortInSearchResult("Showroom");
                    iccPortal.VerifyPortInSearchResult("HMOrder");
                    iccPortal.VerifyPortInSearchResult("VPS");
                    Reporter.ReportEvent("ICC portal 'LicenceCompany_Castor' Processed for OFS, Showroom, HMOrder, VPS",
                        "ICC portal  LicenceCompany_Castor Processed for OFS, Showroom, HMOrder, VPS",
                        HP.LFT.Report.Status.Passed);
                }
                catch
                {
                    Reporter.ReportEvent("ICC portal 'LicenceCompany_Castor' Processed for OFS, Showroom, HMOrder, VPS",
                       "ICC portal  LicenceCompany_Castor Processed for OFS, Showroom, HMOrder, VPS",
                       HP.LFT.Report.Status.Failed);
                }
                driver.SwitchTo().Window(libraryWindow);
                // make it incative
                libraryUi.traverseContentFrame();
                libraryUi.traverseToContentBodyFrame();
                // libraryUi.get_Checkbox(Int32.Parse(libraryUi.get_RowByCompanyName(companyName).GetAttribute("rowIndex")) + 1).Click();
                //companyName = libraryUi.get_LicenseCompaniesTableRowRecordName()[1].Text;
                libraryUi.traverseContentFrame();
                libraryUi.traverseToContentBodyFrameTableSettings();
                libraryUi.get_Drop().Click();
            }
            catch (Exception ex)
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                System.Diagnostics.Debug.WriteLine("Message*********************" + ex.Message);
                System.Diagnostics.Debug.WriteLine("StackTrace*********************" + ex.StackTrace);
                Reporter.ReportEvent("TC26_SIT_Library_License_Companies script fail",
                     "TC26_SIT_Library_License_Companies script fail- exception message: " + ex.Message,
                     HP.LFT.Report.Status.Failed);
                Reporter.ReportEvent("TC26_SIT_Library_License_Companies script fail",
                    "TC26_SIT_Library_License_Companies script fail- exception StackTrace " + ex.StackTrace,
                    HP.LFT.Report.Status.Failed, ESSpath);
            }

        }

      //  [Test]

        public void TC28_SIT_Library_TransportPacking()
        {

            try
            {
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(10000);
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                string username = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAUSER");//"admtempjavas";
                string password = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAPWD");//"admtempjavas";
                string season = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");// "7-2018";
                string office = "CNSH";
                Thread.Sleep(1000);
              
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);
            driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                List<string> lswins = driver.WindowHandles.ToList();
            sGMethods.GetLatestWindow(driver);
            Castorpages castorobjs = new Castorpages(driver);
            //castorobjs.CastorLogin(
            //xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "librarianUser"),
            //xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "librarianPassword"));

            castorobjs.CastorLogin(username, password);
            Reporter.ReportEvent("Login to Application with user : " + username, "Login Pass", HP.LFT.Report.Status.Passed);

            Thread.Sleep(20000);
            //  *******************************************TransportPacking******************************************************************************
            libraryUi.LaunchTransportPacking();
            Reporter.ReportEvent("Launch TransportPacking", "Launch TransportPacking Pass", HP.LFT.Report.Status.Passed);

            libraryUi.traverseContentFrame();
            libraryUi.traverseToContentBodyFrameTableSettings();
            libraryUi.get_ActionLinkForDropDown().Click();
            libraryUi.get_CreateTransportPackingForDropDown().Click();
            Reporter.ReportEvent("Create new TransportPacking", "Create new TransportPacking Pass", HP.LFT.Report.Status.Passed);

            string libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
                Thread.Sleep(20000);
                if (libraryUi.get_CreateTransportPackingText().Displayed)
                System.Diagnostics.Debug.WriteLine("******* PASS");
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_showTypeSelector().Displayed) Thread.Sleep(3000);
            libraryUi.get_showTypeSelector().Click();
            string type = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
                Thread.Sleep(20000);
                libraryUi.get_TransportPackingExpand().Click();
            Thread.Sleep(2000);
            libraryUi.get_PolybagRadioButton().Click();
            libraryUi.get_SelectButton().Click();
                Thread.Sleep(10000);
                driver.SwitchTo().Window(type);
            driver.SwitchTo().Frame("pagecontent");
            libraryUi.get_txtDescription().Click();
            libraryUi.get_txtDescription().SendKeys("Automation Test");
            libraryUi.get_Concept().Click();
            libraryUi.selectSeason(season);
            driver.SwitchTo().DefaultContent();
            libraryUi.get_DoneButton().Click();
            Reporter.ReportEvent("Create new TransportPacking in new popup", "Create new TransportPacking in new popup Failed", HP.LFT.Report.Status.Passed);
                Thread.Sleep(20000);
                driver.SwitchTo().Window(libraryWindow);
            libraryUi.traverseContentFrame();


            libraryUi.switchToFrame();
            libraryUi.get_EditButton().Click();
            Reporter.ReportEvent("Edit created TransportPacking", "Edit created TransportPacking Failed", HP.LFT.Report.Status.Passed);

            libraryUi.get_btnSupplier().Click();
            libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("searchPane");
            sGMethods.SelectDropDownByValue(libraryUi.get_OfficeIdDropDown(), office);
            driver.SwitchTo().DefaultContent();
            libraryUi.get_FindButton().Click();
            Thread.Sleep(2000); IAlert alert = driver.SwitchTo().Alert();
            alert.Accept();
            driver.SwitchTo().Frame("listDisplay");
            Thread.Sleep(2000);
            libraryUi.get_btnSupplierCheckBox().Click();
            Reporter.ReportEvent("Selected supplier for new TransportPacking", "Selected supplier for new TransportPacking Failed",
                HP.LFT.Report.Status.Passed);
            driver.SwitchTo().DefaultContent();
            libraryUi.get_SubmitButton().Click();
            driver.SwitchTo().Window(libraryWindow);
            System.Diagnostics.Debug.WriteLine(driver.Title);
            libraryUi.traverseContentFrame();
            // driver.SwitchTo().Frame("detailsDisplay");
            libraryUi.switchToFrame();
            if (libraryUi.get_DoneButtonCastor().Displayed) Thread.Sleep(3000);
            libraryUi.get_DoneButtonCastor().Click();
            libraryUi.get_Initiatedlink().Click();
            driver.SwitchTo().DefaultContent();
            libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_stateNameHighlight("Initiated").Displayed)
                System.Diagnostics.Debug.WriteLine("::::::::::::::::::::: PASS");
            Reporter.ReportEvent("Status is Initiated", "Status is Initiated Pass", HP.LFT.Report.Status.Passed);

            driver.SwitchTo().DefaultContent();

            libraryUi.get_StatePromote().Click();
            Thread.Sleep(3000);
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_stateNameHighlight("Review").Displayed)
                System.Diagnostics.Debug.WriteLine("::::::::::::::::::::: PASS");
            Reporter.ReportEvent("Status is Review", "Status is Review Pass", HP.LFT.Report.Status.Passed);

            driver.SwitchTo().DefaultContent();

            libraryUi.get_StatePromote().Click();
            Thread.Sleep(3000);
            driver.SwitchTo().Frame("pagecontent"); if (libraryUi.get_stateNameHighlight("Released").Displayed)
                System.Diagnostics.Debug.WriteLine("::::::::::::::::::::: PASS");
            Reporter.ReportEvent("Status is Released", "Status is Released Pass", HP.LFT.Report.Status.Passed);

            driver.Close();
            driver.SwitchTo().Window(libraryWindow);

            iccPortal.LaunchICCWindow();

            iccPortal.SelectValueFromApplicationDropDown("TransportPacking_Castor");
            iccPortal.VerifySearchResult("TransportPacking");
            iccPortal.VerifyPortInSearchResult("TransportPacking.Shredder");
            Reporter.ReportEvent("ICC portal 'PlanInformationOnArticleLevel_PLES' Processed for CDW and Fantomen",
                "ICC portal  PlanInformationOnArticleLevel_PLES Processed for CDW and Fantomen",
                HP.LFT.Report.Status.Passed);
            }
            catch (Exception ex)
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                System.Diagnostics.Debug.WriteLine("Message*********************" + ex.Message);
                System.Diagnostics.Debug.WriteLine("StackTrace*********************" + ex.StackTrace);
                Reporter.ReportEvent("TC28_SIT_Library_TransportPacking script fail",
               "TC28_SIT_Library_TransportPacking Script fail " + ex.Message,
               HP.LFT.Report.Status.Failed, ESSpath);
            }
        }

      //  [Test]

        public void TC25_SIT_Library_ConsumerPackaging()
        {

            try
            {
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(10000);
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                string username = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAUSER");//"admtempjavas";
            string password = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAPWD");//"admtempjavas";
            string season = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");// "7-2018";
            string office = "CNSH";
                Thread.Sleep(1000);
             
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(4000);
                try
                {
                    sGMethods.AlertAccept(driver);
                }
                catch
                {

                }
               
            List<string> lswins = driver.WindowHandles.ToList();
            sGMethods.GetLatestWindow(driver);
            Castorpages castorobjs = new Castorpages(driver);
            //castorobjs.CastorLogin(
            //xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "librarianUser"),
            //xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "librarianPassword"));

            castorobjs.CastorLogin(username, password);
            Reporter.ReportEvent("Login to Application with user : " + username, "Login Pass", HP.LFT.Report.Status.Passed);

            Thread.Sleep(20000);

            //  ********************************************Consumer Packaging*****************************************************************************



            libraryUi.LaunchConsumerPackaging();
            Reporter.ReportEvent("Launch Consumer Packaging", "Launch Consumer Packaging Pass", HP.LFT.Report.Status.Passed);
            libraryUi.traverseContentFrame();
            libraryUi.traverseToContentBodyFrameTableSettings();
            libraryUi.get_ActionLinkForDropDown().Click();
            libraryUi.get_CreateConsumerPackagingForDropDown().Click();
            Reporter.ReportEvent("Create new Consumer Packaging", "Create new Consumer Packaging Pass", HP.LFT.Report.Status.Passed);
            string libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
            if (libraryUi.get_CreateConsumerPackagingText().Displayed)
                System.Diagnostics.Debug.WriteLine("******* PASS");
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_showTypeSelector().Displayed) Thread.Sleep(3000);
            libraryUi.get_showTypeSelector().Click();
            string type = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
            libraryUi.get_PackagingExpand().Click();
            Thread.Sleep(2000);
            libraryUi.get_BoxRadioButton().Click();
            libraryUi.get_SelectButton().Click();
            Reporter.ReportEvent("Create new Consumer Packaging in new popup", "Create new Consumer Packaging in new popup Pass", HP.LFT.Report.Status.Passed);
            driver.SwitchTo().Window(type);
            driver.SwitchTo().Frame("pagecontent");
            libraryUi.get_txtDescription().Click();
            libraryUi.get_txtDescription().SendKeys("Automation Test");
            libraryUi.get_Concept().Click();
            libraryUi.selectSeason(season);
            driver.SwitchTo().DefaultContent();
            libraryUi.get_DoneButton().Click();
            driver.SwitchTo().Window(libraryWindow);
                Thread.Sleep(4000);
            libraryUi.traverseContentFrame();
            driver.SwitchTo().Frame("detailsDisplay");
            libraryUi.get_EditButton().Click();
            Reporter.ReportEvent("Edit created Consumer Packaging", "Edit created Consumer Packaging Pass", HP.LFT.Report.Status.Passed);
            libraryUi.get_btnSupplier().Click();
                Thread.Sleep(5000);
                libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
                Thread.Sleep(15000);
                driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("searchPane");
            sGMethods.SelectDropDownByValue(libraryUi.get_OfficeIdDropDown(), office);
            driver.SwitchTo().DefaultContent(); libraryUi.get_FindButton().Click();
            Thread.Sleep(10000);
                IAlert alert = driver.SwitchTo().Alert();
            alert.Accept();
            driver.SwitchTo().Frame("listDisplay");
            Thread.Sleep(5000);
            libraryUi.get_btnSupplierCheckBox().Click();
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
            libraryUi.get_SubmitButton().Click();
            Reporter.ReportEvent("Selected supplier for new Consumer Packaging", "Selected supplier for new Consumer Packaging Pass",
                HP.LFT.Report.Status.Passed);
            driver.SwitchTo().Window(libraryWindow);
            System.Diagnostics.Debug.WriteLine(driver.Title);
                Thread.Sleep(4000);
            libraryUi.traverseContentFrame();
            driver.SwitchTo().Frame("detailsDisplay");
            if (libraryUi.get_DoneButtonCastor().Displayed) Thread.Sleep(3000);
            libraryUi.get_DoneButtonCastor().Click();
            libraryUi.get_Initiatedlink().Click();
            libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_stateNameHighlight("Initiated").Displayed)
                System.Diagnostics.Debug.WriteLine("::::::::::::::::::::: PASS");
            Reporter.ReportEvent("Status is Initiated", "Status is Initiated Pass", HP.LFT.Report.Status.Passed);
            driver.SwitchTo().DefaultContent();
            libraryUi.get_StatePromote().Click();
            Thread.Sleep(3000);
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_stateNameHighlight("Review").Displayed)
                System.Diagnostics.Debug.WriteLine("::::::::::::::::::::: PASS");
            Reporter.ReportEvent("Status is Review", "Status is Review Pass", HP.LFT.Report.Status.Passed);

            driver.SwitchTo().DefaultContent();
            libraryUi.get_StatePromote().Click();
            Thread.Sleep(3000);
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_stateNameHighlight("Released").Displayed)
                System.Diagnostics.Debug.WriteLine("::::::::::::::::::::: PASS");
            Reporter.ReportEvent("Status is Released", "Status is Released Pass", HP.LFT.Report.Status.Passed);

            driver.Close();
            driver.SwitchTo().Window(libraryWindow);

            iccPortal.LaunchICCWindow();

            iccPortal.SelectValueFromApplicationDropDown("ConsumerPackaging_Castor");
            iccPortal.VerifySearchResult("ConsumerPackaging");
            iccPortal.VerifyPortInSearchResult("ConsumerPackaging.Shredder");
            Reporter.ReportEvent("ICC portal 'ConsumerPackaging_Castor' Processed for Shredder",
                "ICC portal  'ConsumerPackaging_Castor' Processed for Shredder",
                HP.LFT.Report.Status.Passed);
            }
           catch (Exception ex)
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                System.Diagnostics.Debug.WriteLine("Message*********************" + ex.Message);
               System.Diagnostics.Debug.WriteLine("StackTrace*********************" + ex.StackTrace);
                Reporter.ReportEvent("TC25_SIT_Library_ConsumerPackaging script fail",
               "TC25_SIT_Library_ConsumerPackaging Script fail" + ex.Message,
               HP.LFT.Report.Status.Failed, ESSpath);
            }
        }




       // [Test]

        public void TC27_SIT_Library_Materials()
        {
            try
            {
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(10000);
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                string username = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAUSER");//"admtempjavas";
                string password = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAPWD");//"admtempjavas";
                string season = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");// "7-2018";
                string office = "CNSH";
                Thread.Sleep(1000);
               // GeneralMethods sGMethods = new GeneralMethods();
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);
            driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                List<string> lswins = driver.WindowHandles.ToList();
            sGMethods.GetLatestWindow(driver);
            Castorpages castorobjs = new Castorpages(driver);
            //castorobjs.CastorLogin(
            //xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "librarianUser"),
            //xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "librarianPassword"));
            castorobjs.CastorLogin(username, password);
            Reporter.ReportEvent("Login to Application with user : " + username, "Login Pass", HP.LFT.Report.Status.Passed);

            Thread.Sleep(20000);

            //*******************************************************************************************
            libraryUi.Materials();
            Reporter.ReportEvent("Launch Materials", "Launch Materials Pass", HP.LFT.Report.Status.Passed);

            libraryUi.traverseContentFrame();
            libraryUi.tabFrame();
            libraryUi.get_MaterialTab("Denim").Click();
            Reporter.ReportEvent("Launch Materials Denim", "Launch Materials Denim Pass", HP.LFT.Report.Status.Passed);

            libraryUi.traverseContentFrame();
            libraryUi.tabFrame();
            // libraryUi.tabFrame();
            libraryUi.traverseToContentBodyFrameTableSettings();
            libraryUi.get_ActionLinkForDropDown().Click();
            libraryUi.get_CreateCreateFancyForDropDown().Click();
            Reporter.ReportEvent("Create new Materials", "Create new Materials Pass", HP.LFT.Report.Status.Passed);

            string libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
                Thread.Sleep(20000);
                if (libraryUi.get_MaterialText().Displayed)
                System.Diagnostics.Debug.WriteLine("******* PASS");
            sGMethods.SelectDropDownByValue(libraryUi.get_FabricTypeDropDown(), "Comfort Stretch Denim");
            libraryUi.get_txtDescriptionTextBox().Click();
            libraryUi.get_txtDescriptionTextBox().SendKeys("Automation Test");
            sGMethods.SelectDropDownByValue(libraryUi.get_OfficeIdDropDown(), office);
            sGMethods.SelectDropDownByValue(libraryUi.get_SeasonIdDropDown(), season);
            libraryUi.get_DoneButton().Click();
                Thread.Sleep(30000);
                driver.SwitchTo().Window(libraryWindow);
                Thread.Sleep(4000);
            libraryUi.traverseContentFrame();
            driver.SwitchTo().Frame("detailsDisplay");
            Reporter.ReportEvent("Create new Materials in new popup", "Create new Materials in new popup Failed", HP.LFT.Report.Status.Passed);

            libraryUi.get_EditButton().Click();
            Reporter.ReportEvent("Edit created Materials", "Edit created Materials Failed", HP.LFT.Report.Status.Passed);
                Thread.Sleep(1000);
                libraryUi.get_btnSupplier().Click();
                Thread.Sleep(1000);
                libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("searchPane");
            sGMethods.SelectDropDownByValue(libraryUi.get_OfficeIdDropDown(), office);
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
            libraryUi.get_FindButton().Click();
            Thread.Sleep(5000);
                IAlert alert = driver.SwitchTo().Alert();
            alert.Accept();
            driver.SwitchTo().Frame("listDisplay");
            Thread.Sleep(2000);
            libraryUi.get_btnSupplierCheckBox().Click();
                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();
            Reporter.ReportEvent("Selected supplier for new Materials", "Selected supplier for new Materials Failed",
                HP.LFT.Report.Status.Passed);
            libraryUi.get_SubmitButton().Click();
                Thread.Sleep(2000);
                driver.SwitchTo().Window(libraryWindow);
            System.Diagnostics.Debug.WriteLine(driver.Title);
                Thread.Sleep(3000);
                libraryUi.traverseContentFrame();
            driver.SwitchTo().Frame("detailsDisplay");
            if (libraryUi.get_btnFiberContent().Displayed) Thread.Sleep(3000);
            libraryUi.get_btnFiberContent().Click();
            Reporter.ReportEvent("Selected FiberContent for new Materials", "Selected FiberContent for new Materials Failed",
                HP.LFT.Report.Status.Passed);
            libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
            driver.SwitchTo().Frame("fabricContentTable");
            driver.SwitchTo().Frame("tableContentFrame");
            driver.SwitchTo().Frame("tableBodyRight");
            libraryUi.get_materialCompositionType()[2].Click();
            libraryUi.get_materialCompositionType()[2].Clear();
            libraryUi.get_materialCompositionType()[2].SendKeys("100");
            driver.SwitchTo().DefaultContent();
            libraryUi.get_DoneButtonPopUp().Click();
            driver.SwitchTo().Window(libraryWindow);
            libraryUi.traverseContentFrame();
            driver.SwitchTo().Frame("detailsDisplay");
            libraryUi.get_btnPurchaseCost().Clear();
            libraryUi.get_btnPurchaseCost().SendKeys("4");
            sGMethods.SelectDropDownByValue(libraryUi.get_btnCurrencyId(), "EUR");
            sGMethods.SelectDropDownByValue(libraryUi.get_btnPurchaseUOMId(), "cm");
                Thread.Sleep(10000);
                libraryUi.traverseContentFrame();
            driver.SwitchTo().Frame("detailsDisplay");
            libraryUi.get_DoneButtonCastor().Click();
            libraryUi.get_Initiatedlink().Click();
            libraryWindow = libraryUi.get_libraryWindow();
            libraryUi.get_NewLaunchedWindow();
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_stateNameHighlight("Initiated").Displayed)
                System.Diagnostics.Debug.WriteLine("::::::::::::::::::::: PASS");
            Reporter.ReportEvent("Status is Initiated", "Status is Initiated Pass", HP.LFT.Report.Status.Passed);
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
            libraryUi.get_StatePromote().Click();
            Thread.Sleep(3000);
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_stateNameHighlight("Review").Displayed)
                System.Diagnostics.Debug.WriteLine("::::::::::::::::::::: PASS");
            Reporter.ReportEvent("Status is Review", "Status is Review Pass", HP.LFT.Report.Status.Passed);
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
            libraryUi.get_StatePromote().Click();
            Thread.Sleep(3000);
            driver.SwitchTo().Frame("pagecontent");
            if (libraryUi.get_stateNameHighlight("Released").Displayed)
                System.Diagnostics.Debug.WriteLine("::::::::::::::::::::: PASS");
            Reporter.ReportEvent("Status is Released", "Status is Released Pass", HP.LFT.Report.Status.Passed);

            driver.Close();
            driver.SwitchTo().Window(libraryWindow);

                iccPortal.LaunchICCWindow();
            
                try
                {
                    iccPortal.SelectValueFromApplicationDropDown("MaterialBooking_Castor");
                    iccPortal.VerifySearchResult("MaterialBooking_Castor");
                    iccPortal.VerifyPortInSearchResult("MaterialBooking.RM");
                    Reporter.ReportEvent("ICC portal 'MaterialBooking_Castor' Processed for MaterialBooking.RM",
                        "ICC portal  'MaterialBooking_Castor' Processed for MaterialBooking.RM",
                        HP.LFT.Report.Status.Passed);
                }
                catch
                {
                    Reporter.ReportEvent("ICC portal 'MaterialBooking_Castor' Processed for MaterialBooking.RM",
                       "ICC portal  'MaterialBooking_Castor' Processed for MaterialBooking.RM",
                       HP.LFT.Report.Status.Failed);
                }

            }
            catch (Exception ex)
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                System.Diagnostics.Debug.WriteLine("Message*********************" + ex.Message);
                System.Diagnostics.Debug.WriteLine("StackTrace*********************" + ex.StackTrace);
                Reporter.ReportEvent("TC27_SIT_Library_Materials script fail",
               "TC27_SIT_Library_Materials Script fail",
               HP.LFT.Report.Status.Failed, ESSpath);
            }
        }



        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
            xCellFileHelper.CleanUp();
            driver.Quit();
            Thread.Sleep(2000);
        }

        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
        }
    }
}
