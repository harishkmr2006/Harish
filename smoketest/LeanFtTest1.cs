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
   [TestFixture]
    public class LeanFtTest1 : UnitTestClassBase
    {
        IWebDriver driver;
        WebDriverWait wait;
        ExcelHelper xCellFileHelper;
        string datafilePath;
        private PrePlanPage pPlanPage;
        string siticcbamenv = "Acc2010";
        string siticcbamenv1 = "AccTest3";


        public LeanFtTest1()
        {
            Thread.Sleep(1000);
            var binary = new FirefoxBinary(ConfigUtils.Read("FirefoxPath"));
            // // string path = ReadFirefoxProfile();
            FirefoxProfile ffprofile = new FirefoxProfile(ConfigUtils.Read("FirefoxPath"));
            ffprofile.SetPreference("binary", ConfigUtils.Read("FirefoxPath"));
            driver = new FirefoxDriver(binary, ffprofile);
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            Thread.Sleep(1000);
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            datafilePath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + ConfigUtils.Read("TestDataPath");
        }

        
        [OneTimeSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture
        }

        [SetUp]
        public void SetUp()
        {
            try {
               // var options = new FirefoxOptions();
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
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);

                System.Diagnostics.Debug.WriteLine(e.StackTrace);
                Reporter.ReportEvent("Set up issue for firefix", "Set up issue for firefix", HP.LFT.Report.Status.Failed);
            }

        }

      //  [Test, Order(1)]

        public void TC01_Selenium_Castor_ProductDevelopmentandPublish()
        {
            try {
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                //  driver.Manage().Window.Maximize();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                //  sGMethods.WaitAlerttoPresent(driver);
                //  sGMethods.AlertAccept(driver);
                Thread.Sleep(1000);
                List<string> lswins = driver.WindowHandles.ToList();


                sGMethods.GetLatestWindow(driver);

                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "BuyerUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "BuyerPwd"));

               // driver.FindElement(By.XPath(".//input[@value=\"SEBO PD Buyer\"]")).Click();

                Thread.Sleep(20000);
                sGMethods.GetLatestWindow(driver);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                //  sGMethods.WaitFrametoSwitch(driver, "content");
                // sGMethods.GetLatestWindow(driver);
                // sGMethods.WaitFrametoSwitch(driver, "content");
                 
                //  sGMethods.SwitchFrame_Name(driver, "content");


                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");

                 wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='tvcTabs0']//a[@title=\"Product Developments\"]")));
              //  wait.Until(ExpectedConditions.ElementExists(By.XPath(".//li[@title=\"Product Developments\"]")));

                sGMethods.WebButton_Click(castorobjs.Tab_ProductDevelopments, "Product Development Tab1");
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
               // sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMProductDevelopmentBuyerStartPageTabcontentFrame");
                sGMethods.SwitchFrame_Name(driver, castorobjs.frame_tvcTabs0_ProductDevelopmentTabcontentFrame);
                try
                {
                    sGMethods.WebButton_Click(castorobjs.Tab_ProductDevelopments, "Product Development Tab2");
                }
                catch
                {

                }
              
               // sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMProductDevelopmentBuyerStartPagecontentFrame");
               // sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                 sGMethods.SwitchFrame_Name(driver, castorobjs.frame_tvcTabs0_ProductDevelopmentcontentFrame);
                 sGMethods.SwitchFrame_WebElelemnt(driver, castorobjs.CreateProduct_FrameElelemnt);
                Thread.Sleep(2000);
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//span[text()='Create Product Development']")));
                Thread.Sleep(2000);
                sGMethods.WebLink_Click(castorobjs.Link_CreateProductDevelopment, "Create Product Development Link");

                driver.SwitchTo().DefaultContent();
                // sGMethods.WaitFrametoSwitch(driver,);
                sGMethods.GetLatestWindow(driver);

                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//input[@id=\"txtProductName\"]")));

                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_ProductDevelopmentName, "Product Development Name field", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductDevelopName"));
                try
                {
                    sGMethods.selectByVisibleText(castorobjs.DropDown_OwningOffice, "SEBO");
                }
                catch
                {

                }
                try
                {
                    sGMethods.selectByVisibleText(castorobjs.DropDown_Brand, "HM");
                }
                catch
                {

                }
              

                Thread.Sleep(1000);
                // driver.FindElement(By.XPath(".//a[contains(.,\"Create\")]")).Click();
                sGMethods.WebButton_Click(castorobjs.WebButton_Create, "Create");
                Thread.Sleep(10000);

                sGMethods.GetLatestWindow(driver);

                sGMethods.SwitchFrame_Name(driver, "content");

                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                Thread.Sleep(10000);
                string productID = driver.FindElement(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")).Text;
                xCellFileHelper.SetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID", productID);
                Reporter.ReportEvent("ProductID", "Product ID is " + productID, HP.LFT.Report.Status.Passed);
                Thread.Sleep(5000);
                sGMethods.WebButton_Click(castorobjs.EditButton_ProductDevelopment, "Product Development Edit");

                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(text(),\"Season\")]/..//div[@class=\"selectize-input items not-full has-options\"]")));
                sGMethods.WebButton_Click(castorobjs.DropDown_Season, "Season Dropdown");
                // driver.FindElement(By.XPath(".//label[contains(text(),\"Season\")]/..//div[@class=\"selectize-input items not-full has-options\"]")).Click();
                sGMethods.WebButton_Click(driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='" + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season") + "']")), "Season Option");

                // Thread.Sleep(1000);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(text(),\"Department\")]/..//div[@class=\"selectize-input items not-full has-options\"]")));
                sGMethods.WebButton_Click(castorobjs.DropDown_Department, "Department Dropdown");

                sGMethods.WebButton_Click(driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='" + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Dept") + "']")), "Department Option");
                //  Thread.Sleep(2000);

                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(text(),\"Concept\")]/..//div[contains(@class,\"selectize-input items\")]")));
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.DropDown_Concept, "Concepts Dropdown");
                Thread.Sleep(1000);
                string ssection = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Section");
                string[] sasection = ssection.Split('-');
                sGMethods.WebButton_Click(driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='" + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Concept").Trim() + "']")), "Concepts Option");



                // driver.FindElement(By.XPath(".//label[contains(text(),\"Concept\")]/..//div[@class=\"selectize-input items not-full\"]")).Click();
                // driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='Everyday Collection']")).Click();
                // Thread.Sleep(2000);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(text(),\"Product Type\")]/..//div[contains(@class,\"selectize-input items not-full\")]")));
                Thread.Sleep(2000);
                sGMethods.WebButton_Click(castorobjs.DropDown_ProductType, "Product Type Dropdown");
                Thread.Sleep(1000);
              
                // driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='" + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductType") + "']")).

                IJavaScriptExecutor jse = (IJavaScriptExecutor)driver;

                jse.ExecuteScript("arguments[0].scrollIntoView(true);", driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='" + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductType") + "']")));

                sGMethods.WebButton_Click(driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='" + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductType") + "']")), "Product Type Option");
                //driver.FindElement(By.XPath(".//label[contains(text(),\"Product Type\")]/..//div[@class=\"selectize-input items not-full\"]")).Click();
                // driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='Other jacket - Top']")).Click();
                // Thread.Sleep(1000);

                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//label[contains(text(),\"Sales Mode\")]/..//div[@class=\"selectize-input items full has-options has-items\"]")));
                sGMethods.WebButton_Click(castorobjs.DropDown_SalesMode, "SalesMode Dropdown");
                sGMethods.WebButton_Click(castorobjs.Options_SalesMode, "SalesMode Option");


                //  driver.FindElement(By.XPath(".//label[contains(text(),\"Sales Mode\")]/..//div[@class=\"selectize-input items full has-options has-items\"]")).Click();
                //  driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='Single']")).Click();
                //  Thread.Sleep(1000);
                try
                {
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//label[contains(text(),\"Licensed\")]/..//div[@class=\"selectize-input items full has-options has-items\"]")));
                    sGMethods.WebButton_Click(castorobjs.DropDown_Licensed, "Licensed Dropdown");
                    sGMethods.WebButton_Click(castorobjs.Options_Licensed, "Licensed Option");
                }
                catch
                {

                }


                // driver.FindElement(By.XPath(".//label[contains(text(),\"Licensed\")]/..//div[@class=\"selectize-input items full has-options has-items\"]")).Click();
                // driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='No']")).Click();
                //  Thread.Sleep(1000);
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_TargetPrice, "Target Price", "100");
                // driver.FindElement(By.XPath(".//label[contains(text(),\"Currency\")]/..//div[contains(@class,\"selectize-input items not-full has-options\")]")).Click();
                //  driver.FindElement(By.XPath("//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='SEK']")).Click();
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//label[contains(text(),\"Currency\")]/..//div[contains(@class,\"selectize-input items not-full has-options\")]")));
                sGMethods.WebButton_Click(castorobjs.DropDown_Currency, "Licensed Dropdown");
                Thread.Sleep(500);
                sGMethods.WebButton_Click(castorobjs.Options_Currency, "Licensed Option");

                //  driver.FindElement(By.XPath(".//label[contains(text(),\"Target Price\")]/../input[@name=\"section6_targetPrice\"]")).SendKeys("100");
                // sGMethods.WebEdit_SetValue(castorobjs.WebEdit_TargetPrice,"Target Price", "100");
                Thread.Sleep(2000);
                sGMethods.WebButton_Click(castorobjs.WebButton_Save, "Save");
                Thread.Sleep(10000);
                //  driver.FindElement(By.XPath(".//button[contains(text(),'Save')]")).Click();

                //  Thread.Sleep(10000);
                //  driver.FindElement(By.XPath(".//div[@class=\"he-tabs-bar\"]/ul[@class=\"tabs\"]//a[@controls=\"tab_Description\"]")).Click();

                sGMethods.WebLink_Click(castorobjs.Tab_Description, "Dscription tab");
                Thread.Sleep(2000);
                // driver.FindElement(By.XPath(".//*[@id='Appearances']//a/span[contains(.,\"Create Appearance(s)\")]")).Click();
                sGMethods.WebLink_Click(castorobjs.Link_CreateAppearance, "Create Appearance(s)");
                //  driver.Navigate().GoToUrl("http://iccportal/ICCPortal/BAM/BPOC");

                //lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}
                sGMethods.GetLatestWindow(driver);
                string scoloroption = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ColorApperance");
                string[] sColoredit = scoloroption.Split('-');
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_Colour, "Color Edit box", sColoredit[0]);
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//table/tbody/tr//td[contains(.,\"" + scoloroption + "\")]")));
                Thread.Sleep(1000);
                // driver.FindElement(By.XPath(".//*[@id='colorName']")).SendKeys("99");
                //  driver.FindElement(By.XPath(".//table/tbody/tr//td[contains(.,\"99-101 Green Yellowish Dark\")]")).Click();
                sGMethods.WebLink_Click(driver.FindElement(By.XPath(".//table/tbody/tr//td[contains(.,\"" + scoloroption + "\")]")), "Color option");
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_Appearance, "Apperance Edibox", "COLOR");
                // driver.FindElement(By.XPath(".//*[@id='txtAppearanceName']")).SendKeys("COLOR");
                // driver.FindElement(By.XPath(".//*[@id='.tvc.btn.continue.a']/a")).Click();
                sGMethods.WebButton_Click(castorobjs.WebButton_Appearance_Create, "Create");
                Thread.Sleep(10000);
                
                lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }

                // driver.SwitchTo().Window(sparentwindow);
                // Thread.Sleep(1000);
                // Actions sact = new Actions(driver);
                // sact.SendKeys(Keys.Tab);
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame("content");
                //            driver.SwitchTo().Frame("detailsDisplay");
                sGMethods.WebLink_Click(castorobjs.WebLink_tab_LabelAndPackaging, "tab_LabelAndPackaging");
                // driver.FindElement(By.XPath(".//div[@class=\"he-tabs-bar\"]/ul[@class=\"tabs\"]//a[@controls=\"tab_LabelAndPackaging\"]")).Click();
                Thread.Sleep(2000);
                sGMethods.WebLink_Click(castorobjs.WebLink_Label_Actions, "Actions");
                //  driver.FindElement(By.XPath("//div[@id=\"tab_LabelAndPackaging\"]//ul[@class=\"he-tb-toolbar\"]//span[contains(.,\"Actions\")]")).Click();
                Thread.Sleep(2000);
                sGMethods.WebLink_Click(castorobjs.WebLink_AddLabel, "Add Label");
                // driver.FindElement(By.XPath("//div[@id=\"tab_LabelAndPackaging\"]//ul[@class=\"he-tb-toolbar\"]//ul[@class=\"he-tb-menu\"]//span[text()=\"Add Label\"]")).Click();

                Thread.Sleep(5000);
                //lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}
                sGMethods.GetLatestWindow(driver);



                // sGMethods.WebEdit_SetValue(castorobjs.WebEdit_AddLabel_Search, "Add Label Search Edit box", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "LabelSearch"));
                // driver.FindElement(By.XPath(".//*[@id='txtTextSearch']")).SendKeys("11");
                castorobjs.PDMerch_Options_LabelSelect();
                
                 wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='mx_btn-search']")));
                sGMethods.WebButton_Click(castorobjs.WebButton_AddLabel_Search, "Addlabel Search");
                castorobjs.waitForSpinner();
                castorobjs.waitforDataLoadSpinner();
                //  driver.FindElement(By.XPath(".//*[@id='mx_btn-search']")).Click();
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame("structure_browser");
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                //  driver.FindElement(By.XPath("//table[@id=\"tbl\"]/tbody/tr/td[1]/input[@type=\"checkbox\"]")).Click();
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//table[@id=\"tbl\"]/tbody/tr/td[1]/input[@type=\"checkbox\"])[1]")));
                sGMethods.WebLink_Click(castorobjs.WebCheckBox_AddLabel_SearchResult1, "Add Label search Checkbox");
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame("structure_browser");
                driver.SwitchTo().Frame("tableContentFrame");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='toolbar-container']//span[contains(.,\"Add Selected\")]")));
                sGMethods.WebButton_Click(castorobjs.WebButton_AddLabel_AddSelected, "AddSelected button");
                Thread.Sleep(5000);
                lswins = driver.WindowHandles.ToList();
                int iswin = 1;
                while ((lswins.Count == 2) && (iswin <= 3000))
                {
                    lswins = driver.WindowHandles.ToList();
                    iswin = iswin + 1;
                    Thread.Sleep(1000);
                }


                Thread.Sleep(2000);
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");

                sGMethods.WebLink_Click(castorobjs.WebLink_tab_Offices, "Offices Tab");
                Thread.Sleep(5000);
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    // Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }

                //  driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@class=\"he - iframe \"]")));
                //  driver.SwitchTo().DefaultContent();
                //  sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("tableContentFrame");
                //  driver.FindElement(By.XPath(".//*[@id='toolbar-container']/ul[1]/li[1]")).Click();


                sGMethods.WebLink_Click(castorobjs.WebLink_Label_AddOffice, "Add Office Link");


                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tvcInlineFrame_0");
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                // driver.FindElement(By.XPath(".//table[@id=\"tbl\"]/tbody/tr[1]/td[1]")).Click();
                // sGMethods.WaitElementToExists(driver, ".//table[@id=\"tbl\"]/tbody/tr/td/div[contains(.,\"Name: CNHK\")]/../../td/input[@type=\"checkbox\"]");
                castorobjs.WebCheck_Office(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office"));
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tvcInlineFrame_0");
                driver.SwitchTo().Frame("tableContentFrame");
                // driver.SwitchTo().Frame("tableContentFrame");
                sGMethods.WebLink_Click(castorobjs.WebLink_Office_AddSelected, "Office Add Selected");
                Thread.Sleep(20000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                // driver.SwitchTo().Frame("tvcInlineFrame_0");
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                sGMethods.WebLink_Click(castorobjs.WebCheck_Office_Publish, "Office WebCheckbox for Publish");
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                // driver.SwitchTo().Frame("tvcInlineFrame_0");
                driver.SwitchTo().Frame("tableContentFrame");
                sGMethods.WebLink_Click(castorobjs.WebLink_Office_Publish, "Office WebLink for Publish");
                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                //            driver.SwitchTo().Frame("content");
                driver.SwitchTo().Frame("overlayContentFrame");
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_Publish_TOField, "Publish To Field", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "MailId"));
                sGMethods.WebLink_Click(castorobjs.Weboption_Publish_Searchoption1, xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "MailId"));
                sGMethods.WebButton_Click(castorobjs.Webbutton_Publish_Publish);

                Thread.Sleep(20000);

                lswins = driver.WindowHandles.ToList();
                iswin = 1;




                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    castorobjs.CastorLogout();
                    driver.Close();

                }


            }
            catch
            {
               
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory")+"ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Castor ProductDevelopment till Publish Script failed", "Castor ProductDevelopment till Publish Script failed", HP.LFT.Report.Status.Failed, ESSpath);
                List<string> lswins = driver.WindowHandles.ToList();
               
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Assert.IsTrue(2 == 3);
            }
        }

      //  [Test, Order(2)]
        public void TC02_Selenium_ICCVerificationexportsTillPDPublish()
        {
            try {
                Thread.Sleep(3000);
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(3000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_ICC"));
                Thread.Sleep(1000);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                ICCBAMPageObj ICCBamobjs = new ICCBAMPageObj(driver);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Acc2010");
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "BopoAcc2010");
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCATENV"));
                }
                if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCSITENV"));
                   // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Acc2010");
                }
                if (ConfigUtils.Read("ENV") == "DIT")
                {
                  
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Int2010");
                }
                string ssendtime = "";
                try
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Application, "Application DropDown", "ProductDevelopment_Castor");
                    Thread.Sleep(2000);
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "ProductDevelopment");
                    /// sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "AssignedProductDevelopment");
                    sGMethods.WebEdit_SetValue(ICCBamobjs.WebEdit_ProductID, "Product ID", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                    sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                    wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                    Thread.Sleep(2000);

                    try
                    {
                        ssendtime = ICCBamobjs.Time_ProductDevelopment_Sendports();
                    }
                    catch
                    {

                    }
                }
                catch
                {

                }
               
               

                if (ssendtime == "")
                {
                    ssendtime = DateTime.Now.AddMinutes(-2).ToLongDateString().ToString();
                }

                //// string ssendtime = "2017-07-03 12:47:29";
                //  bool bval= ICCBamobjs.Castor_ExportElementExistinTimeRange("ProductDevelopment.RM.Primary.SendPort", ssendtime);

                sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "AssignedProductDevelopment");

                sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                Thread.Sleep(2000);

                ICCBamobjs.Verify_ProductDevelopment_SendportswithTimestamp(ssendtime);
                driver.Close();

                // ICCBamobjs.WebEdit_ProductID.GetCssValue("color");
            }
            catch
            {
                Reporter.ReportEvent("Selenium_ICCVerificationexportsTillPDPublish Script failed", "Selenium_ICCVerificationexportsTillPDPublish Script failed", HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }

      //  [Test,Order(3)]

        public void TC03_Selenium_Castor_Merch_PublishToTillHandover()
        {
            try {
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                driver.Manage().Window.Maximize();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                //  sGMethods.WaitAlerttoPresent(driver);
                //  sGMethods.AlertAccept(driver);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchPwd"));

                //   driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                Thread.Sleep(20000);
               

                string sparentwindow = sGMethods.GetCurrentWindow(driver);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='GlobalNewTEXT']")));
              
                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox,"SearchBox Text" ,xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();
                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }
                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }

               

                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                string offices = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office");
                string[] offciesarr = offices.Split(':');
                castorobjs.PDMerch_EditClick_OfficeProduct(offciesarr[1].Trim());

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);
                // driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                Thread.Sleep(5000);
                sGMethods.GetLatestWindow(driver);
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                Thread.Sleep(5000);
                sGMethods.WebButton_Click(castorobjs.EditButton_ProductDevelopment, "Product Development Edit for PD Merch");
                castorobjs.PDMerch_PDEdit_Weight.Clear();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_PDEdit_Weight, "100");
                sGMethods.WebButton_Click(castorobjs.WebButton_Save, "Save");
                Thread.Sleep(5000);
                // sGMethods.GetLatestWindow(driver);
                sGMethods.WebLink_Click(castorobjs.PDMerch_WebLink_tab_SupplierandRequest, "Supplier and Request Tab");
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                Thread.Sleep(5000);
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    // Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                sGMethods.WebLink_Click(castorobjs.PDMerch_WebLink_AddSupplier, "Add Supplier ImageLink");
                Thread.Sleep(5000);
                sGMethods.GetLatestWindow(driver);
                castorobjs.addSupplierDetails(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductionGroup"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductionType"));

                //castorobjs.PDMerch_SearchSuppliersByAttribute("Production Type");
                //Thread.Sleep(5000);
                //castorobjs.PDMerch_AddSupplierByProductTypeSearch("Bag Leather");
                //Thread.Sleep(5000);
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='mx_btn-search']")));
                sGMethods.WebButton_Click(castorobjs.PDMerch_Search_AddSupplier, "Search button on Add Supplier window");
                //Thread.Sleep(5000);
                //sGMethods.WebButton_Click(castorobjs.PDMerch_miniClose_AddSupplieroptions, "MiniClose button for Add Supplier Option window");
                //Thread.Sleep(10000);
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame("structure_browser");
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//table[@id=\"tbl\"]/tbody/tr/td[1]/input[@type=\"checkbox\"])[1]")));
                sGMethods.WebLink_Click(castorobjs.WebCheckBox_AddLabel_SearchResult1, "Add Supplier search Checkbox1");
                string ssuppliername = driver.FindElement(By.XPath(".//table[@id=\"tbr\"]/tbody/tr[1]/td[1]")).Text;

                System.Environment.SetEnvironmentVariable("SupplierName", ssuppliername);

                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame("structure_browser");
                driver.SwitchTo().Frame("tableContentFrame");
                sGMethods.WebButton_Click(castorobjs.WebButton_AddLabel_AddSelected, "AddSelected button");
                Thread.Sleep(1000);
                lswins = driver.WindowHandles.ToList();
                int iswin = 1;
                while ((lswins.Count == 2) && (iswin <= 360))
                {
                    lswins = driver.WindowHandles.ToList();
                    iswin = iswin + 1;
                    Thread.Sleep(1000);
                }
                //  string ssuppliername = "APEX LINGERIE LIMITED";
                Thread.Sleep(1000);
                lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");

                sGMethods.WebLink_Click(castorobjs.PDMerch_WebLink_tab_options, "Options Tab");
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                string ssup1 = ssuppliername;
                if (ssuppliername.Length > 20)
                {
                     ssup1 = ssuppliername.Substring(0, 19);
                    Thread.Sleep(1000);
                }
                try
                {
                    if (driver.FindElement(By.XPath(".//table//tr//td[contains(.,'" + ssup1 + "')]")).Displayed == true)
                    {
                        Reporter.ReportEvent("Supplier name display check in Options page", "Supplier Name displayed in Option page " + ssuppliername, HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("Supplier name display check in Options page", "Supplier Name displayed in Option page " + ssuppliername, HP.LFT.Report.Status.Failed);
                    }

                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("Supplier name display check in Options page", "Supplier Name" + ssuppliername + " displayed in Option page check raised exception- " + e.Message, HP.LFT.Report.Status.Failed);
                }

                castorobjs.PDMerch_Options_ProductGroup_Click(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductDevelopName"));
                Thread.Sleep(5000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");
                driver.SwitchTo().Frame("TopPanelContentFrame");
                sGMethods.WebButton_Click(castorobjs.PDMerch_WebButton_OptionGroupEditbutton, "OptionGroupEdit button");
                Thread.Sleep(1000);
                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_WebList_OptionGroupEditPage_ProductGroup, "Production Group", "Woven");

                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_WebList_OptionGroupEditPage_productionType, "Production Type", "Dresses");
                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_WebList_OptionGroupEditPage_productionApperance, "Production Appearance", "Casual - Dresses");

                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_WebList_OptionGroupEditPage_COP, "Country of Production", "China");
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");

                driver.SwitchTo().Frame("tvcTabs0_HMOptionGroupDescriptionTabcontentFrame");
                driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"bom-list\"]")));
                driver.SwitchTo().Frame("tableContentFrame");
                sGMethods.WebButton_Click(castorobjs.PDMerch_Action_OptionGroupEditPage_BOM, "OptionGroupEdit BOM Action Button");

                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");

                driver.SwitchTo().Frame("tvcTabs0_HMOptionGroupDescriptionTabcontentFrame");
                driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"bom-list\"]")));
                driver.SwitchTo().Frame("tableContentFrame");
                // driver.SwitchTo().Frame("tableBodyRight");
                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_OptionGroupEditPage_BOM, "OptionGroupEdit BOM Action option Add Leather Button");
                Thread.Sleep(5000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }

                sGMethods.WebButton_Click(castorobjs.PDMerch_LeatherSearch_OptionGroupEditPage_BOM, "Add Leather Window Search Button");
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "structure_browser");
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                sGMethods.WebLink_Click(castorobjs.WebCheckBox_AddLabel_SearchResult1, "Add Leather search Checkbox1");

                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "structure_browser");
                driver.SwitchTo().Frame("tableContentFrame");

                sGMethods.WebButton_Click(castorobjs.WebButton_AddLabel_AddSelected, "AddSelected button");
                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_WebList_Position_OptionGroupEditPage_BOM, "Shell");
                sGMethods.WebButton_Click(castorobjs.PDMerch_WebButton_PositionAdd_OptionGroupEditPage_BOM, "Add button for Position");
                Thread.Sleep(1000);
                lswins = driver.WindowHandles.ToList();
                iswin = 1;
                while ((lswins.Count == 3) && (iswin <= 360))
                {
                    lswins = driver.WindowHandles.ToList();
                    iswin = iswin + 1;
                    Thread.Sleep(1000);
                }
                Thread.Sleep(1000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                    driver.Manage().Window.Maximize();
                }

                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");
                driver.SwitchTo().Frame("tvcTabs0_HMOptionGroupDescriptionTabcontentFrame");
                driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"bom-list\"]")));
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                sGMethods.WebLink_Click(castorobjs.PDMerch_WebCheck_SelectAll_OptionGroupEditPage_BOM, "WebCheck box for Select All BOMs");
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");

                driver.SwitchTo().Frame("tvcTabs0_HMOptionGroupDescriptionTabcontentFrame");
                driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"bom-list\"]")));
                driver.SwitchTo().Frame("tableContentFrame");
                sGMethods.WebButton_Click(castorobjs.PDMerch_Action_OptionGroupEditPage_BOM, "OptionGroupEdit BOM Action Button");

                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");

                driver.SwitchTo().Frame("tvcTabs0_HMOptionGroupDescriptionTabcontentFrame");
                driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"bom-list\"]")));
                driver.SwitchTo().Frame("tableContentFrame");
                // driver.SwitchTo().Frame("tableBodyRight");
                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionEditOption_OptionGroupEditPage_BOM, "OptionGroupEdit BOM Action option Edit Button");
                Thread.Sleep(5000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }


                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_EditOption_Placement_OptionGroupEditPage_BOM, "BAG");
                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_EditOption_countryOfOrigin_OptionGroupEditPage_BOM, "China");
               /// sGMethods.WebEdit_SetValue(castorobjs.PDMerch_EditOption_Consefficiency_OptionGroupEditPage_BOM, "ConsEfficiency", "100");
                //driver.FindElement(By.Id("efficiency")).SendKeys("100");
                sGMethods.WebButton_Click(castorobjs.PDMerch_WebButton_PositionAdd_OptionGroupEditPage_BOM, "update button for Position");
                try
                {
                    Thread.Sleep(10000);
                    // sGMethods.WaitAlerttoPresent(driver);
                    sGMethods.AlertAccept(driver);
                }
                catch
                {

                }
               
                Thread.Sleep(2000);
                lswins = driver.WindowHandles.ToList();
                Thread.Sleep(10000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                driver.SwitchTo().Window(lswins[(lswins.Count) - 1]).Close();
                Thread.Sleep(6000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");

                //driver.SwitchTo().DefaultContent();
                //sGMethods.WaitFrametoSwitch(driver, "content");
                //driver.SwitchTo().Frame("detailsDisplay");
                //driver.SwitchTo().Frame("tvcTabs0_HMOptionGroupDescriptionTabcontentFrame");
                //driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"bom-list\"]")));
                //driver.SwitchTo().Frame("tableContentFrame");
                //driver.SwitchTo().Frame("tableBodyRight");
                sGMethods.WebLink_Click(castorobjs.PDMerch_WebCheck_SelectAll_OptionGroupEditPage_BOM, "WebCheck box for Select All BOMs");

                sGMethods.WebLink_Click(castorobjs.PDMerch_OptionGroup_SupplierCheckbox, "WebCheck box for Select Supplier to add Option");

                //driver.SwitchTo().DefaultContent();
                //sGMethods.WaitFrametoSwitch(driver, "content");
                //driver.SwitchTo().Frame("detailsDisplay");

                //driver.SwitchTo().Frame("tvcTabs0_HMOptionGroupDescriptionTabcontentFrame");
                //driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"bom-list\"]")));
                //driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                sGMethods.WebButton_Click(castorobjs.PDMerch_Action_OptiontoAdd, "OptionGroup Action Button to Create Option");
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_OptiontoAdd, "OptionGroup Action Option to Create Option");
                Thread.Sleep(10000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");
                driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"top-panel\"]")));
                driver.SwitchTo().Frame("TopPanelContentFrame");

                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_EditButton, "OptionGroup Action Option to Create Option");
                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_ActionOption_CountryOfProduction, "China");
                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_ActionOption_CountryOfDelivery, "China");
                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_ActionOption_CountryOfOrigin, "China");



                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");

                driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"cost\"]")));
                driver.SwitchTo().Frame("tableContentFrame");

                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_AddCostButton, "OptionGroup Action Option Add Cost button");
                Thread.Sleep(10000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);

                //  driver.SwitchTo().DefaultContent();
                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_AddCost_OrderPlacementDate, "OptionGroup Action Option Add Cost Order Placement date");
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_AddCost_OrderPlacementDateSelect, "OptionGroup Action Option Add Cost Order Placement Today date selection");
                castorobjs.PDMerch_ActionOption_AddCost_MeterialCost.Clear();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_ActionOption_AddCost_MeterialCost, "MeterialCost", "999");
                sGMethods.TAB_Click(castorobjs.PDMerch_ActionOption_AddCost_MeterialCost, driver);
                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_AddCost_SaveCost, "OptionGroup Action Option Add Cost Save Button");
                Thread.Sleep(5000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");
                driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"top-panel\"]")));
                //  driver.SwitchTo().Frame("TopPanelContentFrame");

                /// sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_OptionDetails_SetStatus, "OptionGroup Action Option ");
                castorobjs.PDMerch_Options_SetSecureStatus();
                Thread.Sleep(2000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("detailsDisplay");
                driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadget\"]//iframe[@data-config-id=\"top-panel\"]")));
                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_OptionDetails_Lock, "OptionGroup OptionDetails Lock button");
                Thread.Sleep(2000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                driver.SwitchTo().Window(lswins[(lswins.Count) - 1]).Close();
                Thread.Sleep(2000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                sGMethods.WebLink_Click(castorobjs.PDMerch_ActionOption_SupplertoHandOver, "WebCheck box for Select Supplier to HandOver");

                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                sGMethods.WebLink_Click(castorobjs.PDMerch_ActionOption_HandOverLink, "WebCheck box for Select Supplier to HandOver");
                Thread.Sleep(10000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("overlayContentFrame");
                sGMethods.WebButton_Click(castorobjs.PDMerch_ActionOption_MailHandOverButton, "OptionGroup Mail Handover Button");

                Thread.Sleep(5000);

                lswins = driver.WindowHandles.ToList();
                iswin = 1;
                while ((lswins.Count == 3) && (iswin <= 120))
                {
                    lswins = driver.WindowHandles.ToList();
                    iswin = iswin + 1;
                    Thread.Sleep(1000);
                }

                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");

                if (castorobjs.PDMerch_ActionOption_HandOvertextToVerify.Displayed)
                {
                    Reporter.ReportEvent("OptionsGroup Handover text verification", "Handover text displayed", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OptionsGroup Handover text verification", "Handover text not displayed", HP.LFT.Report.Status.Failed);
                }

                lswins = driver.WindowHandles.ToList();
                iswin = 1;

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Thread.Sleep(2000);
            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Selenium_Castor_Merch_PublishToTillHandover Script failed", "Selenium_Castor_Merch_PublishToTillHandover Script failed", HP.LFT.Report.Status.Failed, ESSpath);
                List<string> lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Assert.IsTrue(2 == 3);
            }
        }
       // [Test,Order(4)]
        public void TC04_Selenium_ICCExportsverificationforHandOver()
        {
            try
            {
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(3000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_ICC"));
                Thread.Sleep(1000);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                ICCBAMPageObj ICCBamobjs = new ICCBAMPageObj(driver);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Acc2010");
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "BopoAcc2010");
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCATENV"));
                }
                if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCSITENV"));
                  //  sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Acc2010");
                }
                if (ConfigUtils.Read("ENV") == "DIT")
                {

                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Int2010");
                }
                string ssendtime = "";
                try
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Application, "Application DropDown", "ProductDevelopment_Castor");
                    Thread.Sleep(1000);
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "ProductDevelopment");
                    /// sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "AssignedProductDevelopment");
                    sGMethods.WebEdit_SetValue(ICCBamobjs.WebEdit_ProductID, "Product ID", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                    sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                    wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                    Thread.Sleep(1000);

                    try
                    {
                        ssendtime = ICCBamobjs.Time_ProductDevelopment_Sendports();
                    }
                    catch
                    {

                    }
                }
                catch
                {

                }
               
               
                if ((ssendtime == "")|| (ssendtime.Equals(null)))
                {
                    ssendtime = DateTime.Now.AddMinutes(-2).ToLongDateString().ToString();
                }
                //// string ssendtime = "2017-07-03 12:47:29";
                //  bool bval= ICCBamobjs.Castor_ExportElementExistinTimeRange("ProductDevelopment.RM.Primary.SendPort", ssendtime);

                sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "DevelopmentOptionGroup");
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                Thread.Sleep(1000);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                Thread.Sleep(1000);
                ICCBamobjs.Verify_DevelopmentOptionGroup_SendportswithTimestamp(ssendtime);
                driver.Close();
                Thread.Sleep(1000);
            }
            catch
            {
                Reporter.ReportEvent("Selenium_ICCExportsverificationforHandOver Script failed", "Selenium_ICCExportsverificationforHandOver Script failed", HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }
       // [Test,Order(5)]
        public void TC05_Selenium_Castor_Buyer_PDRTO()
        {
            try {
                Thread.Sleep(10000);
                driver.Manage().Window.Maximize();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();

                Thread.Sleep(1000);
                // driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                // sGMethods.WaitAlerttoPresent(driver);
                // sGMethods.AlertAccept(driver);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "BuyerUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "BuyerPwd"));

                // driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"SEBO PD Buyer\"]")).Click();

                Thread.Sleep(10000);
                //  wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='AEFGlobalFullTextSearch']")));
                //  wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                //  sGMethods.WaitFrametoSwitch(driver, "content");
                // sGMethods.GetLatestWindow(driver);
                // sGMethods.WaitFrametoSwitch(driver, "content");
                //  sGMethods.GetLatestWindow(driver);
                // sGMethods.SwitchFrame_Name(driver, "content");

                string sparentwindow = sGMethods.GetCurrentWindow(driver);

                wait.Until(ExpectedConditions.ElementExists(By.XPath("//div[@id=\"GTBsearchDiv\"]//input[@id=\"GlobalNewTEXT\"]")));
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                sGMethods.GetLatestWindow(driver);
                sGMethods.SwitchFrame_Name(driver, "content");

                sparentwindow = sGMethods.GetCurrentWindow(driver);

                //var swait = new WebDriverWait(driver, TimeSpan.FromSeconds(60000));

                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='tvcTabs0']//a[@title=\"Product Developments\"]")));
                // Thread.Sleep(10000);
                //  wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='toolbar-container']/ul/li/span[contains(.,\"Filter Objects\")]")));
                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, "SearchBox Text", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                // castorobjs.PDMerch_SearchBox.SendKeys(OpenQA.Selenium.Keys.Tab);
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();
                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }
                Thread.Sleep(2000);
                //  wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='AEFGlobalFullTextSearch']")));

          
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }

              
                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                castorobjs.PDMerch_EditClick_BuyerProduct(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);
                // driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                Thread.Sleep(1000);


                //driver.SwitchTo().DefaultContent();
                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");
                ////            sGMethods.SwitchFrame_Name(driver, "gsFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                ////sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                //castorobjs.PDMerch_EditClick_BuyerProduct(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                //driver.SwitchTo().DefaultContent();
                //driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                sGMethods.GetLatestWindow(driver);
                sGMethods.SwitchFrame_Name(driver, "content");
                Thread.Sleep(2000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);

                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");

                sGMethods.WebLink_Click(castorobjs.PDMerch_WebLink_tab_options, "Options Tab");
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                if (castorobjs.PDBuyer_ActionOption_ReceivedtextToVerify.Displayed)
                {
                    Reporter.ReportEvent("OptionsGroup Received text verification", "Received text displayed", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OptionsGroup Received text verification", "Received text not displayed", HP.LFT.Report.Status.Failed);
                }
                sGMethods.WebLink_Click(castorobjs.PDMerch_ActionOption_SupplertoHandOver, "WebCheck box for Select Supplier to Proceed");
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                sGMethods.WebLink_Click(castorobjs.PDBuyer_ActionOption_ProceedLink, "WebLink Proceed to Proceed");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.PDBuyer_ActionOption_ProceedLink_ProccedOption, "WebLink Proceedoption to Proceed");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                driver.SwitchTo().Frame("overlayContentFrame");

                sGMethods.WebButton_Click(castorobjs.PDBuyer_ActionOption_ProceedLink_RTOButton, "WebButton RTO to Proceed");
                Thread.Sleep(10000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                if (castorobjs.PDBuyer_ActionOption_ProceedtextToVerify.Displayed)
                {
                    Reporter.ReportEvent("OptionsGroup Proceed text verification", "Proceed text displayed", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OptionsGroup Proceed text verification", "Proceed text not displayed", HP.LFT.Report.Status.Failed);
                }

                lswins = driver.WindowHandles.ToList();


                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Selenium_Castor_Buyer_PDRTO script failed", "Selenium_Castor_Buyer_PDRTO script failed", HP.LFT.Report.Status.Failed, ESSpath);
                 List<string> lswins = driver.WindowHandles.ToList();


                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Assert.IsTrue(2 == 3);

            }

        }

       // [Test, Order(6)]
        public void TC06_Selenium_ICCExportsverificationforRTO()
        {
            try
            {
                Thread.Sleep(3000);
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_ICC"));
                Thread.Sleep(1000);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                ICCBAMPageObj ICCBamobjs = new ICCBAMPageObj(driver);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Acc2010");
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "BopoAcc2010");
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCATENV"));
                }
                if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCSITENV"));
                   // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Acc2010");
                }
                if (ConfigUtils.Read("ENV") == "DIT")
                {

                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Int2010");
                }
                try
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Application, "Application DropDown", "ProductDevelopment_Castor");
                    Thread.Sleep(2000);
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "ProductDevelopment");
                    /// sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "AssignedProductDevelopment");
                    sGMethods.WebEdit_SetValue(ICCBamobjs.WebEdit_ProductID, "Product ID", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                    sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                    wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                    //  Thread.Sleep(30000);
                    ICCBamobjs.Verify_ProductDevelopment_Sendports_3();
                }
                catch
                {
                    Reporter.ReportEvent("Selenium_ICCExportsverificationforHandOver Script failed", "Selenium_ICCExportsverificationforHandOver Script failed", HP.LFT.Report.Status.Failed);
                }
               
               
                driver.Close();
            }
            catch
            {
                Reporter.ReportEvent("Selenium_ICCExportsverificationforHandOver Script failed", "Selenium_ICCExportsverificationforHandOver Script failed", HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }




       // [Test, Order(9)]

        public void TC09_Selenium_PPlan_EditQunatity_Colorvalidation()
        {
            try {
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                string sSeason = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");
                string sSection = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Section");
                string sDept = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Dept");
                string sPplanProduct = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PPlanProduct");
                string tabName = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "TabName");
                Thread.Sleep(1000);
                pPlanPage = new PrePlanPage(driver);
                Thread.Sleep(1000);
                driver.Manage().Window.Maximize();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_PPlan"));
                Thread.Sleep(1000);
                pPlanPage.waitForPageToLoad();
                pPlanPage.waitForSpinnerToDisappear();
              
               
                if (!pPlanPage.seasonButtonStatus.Text.Contains(sSeason))
                {
                    sGMethods.WebButton_Click(pPlanPage.get_seasonSelectionDirectiveButton(sSeason));
                }
                Reporter.ReportEvent("Season selected", "Season selected : " + sSeason, HP.LFT.Report.Status.Passed);
                pPlanPage.waitForPageToLoad();
                Assert.IsTrue(pPlanPage.get_ProductNameTableHeaderElement().Displayed);
                if (!pPlanPage.SectionDirectrivesDropDownElementSelectedElement.Text.Contains(sSection))
                {
                    sGMethods.SelectDropDownByValue(pPlanPage.SectionDirectrivesDropDownElement,
                        sSection);
                    Thread.Sleep(4000);
                    pPlanPage.waitForPageToLoad();
                }
                Reporter.ReportEvent("Directive selected from dropdown", "Directive selected from dropdown : " + sSection,
                    HP.LFT.Report.Status.Passed);
                try
                {
                    driver.FindElement(By.XPath("//button[text()='Full']")).Click();
                    pPlanPage.waitForPageToLoad();
                    pPlanPage.waitForSpinnerToDisappear();
                }
                catch
                {
                }
                if (pPlanPage.get_departmentTabElement(tabName).Displayed)
                {
                    if (!pPlanPage.DepartmentTabActivElement.Text.Contains(tabName))
                    {
                        pPlanPage.get_departmentTabElement(tabName).Click();
                        pPlanPage.waitForPageToLoad();
                        pPlanPage.waitForSpinnerToDisappear();
                    }
                }
                Reporter.ReportEvent("Tab selected", "Tab selected : " + tabName, HP.LFT.Report.Status.Passed);
                Thread.Sleep(4000);
                Assert.IsTrue(pPlanPage.get_ProductNameTableHeaderElement().Displayed);
                pPlanPage.waitForSpinnerToDisappear();
                driver.FindElement(By.XPath(".//*[@id='center']/div/div[4]/div[1]/div/div[1]/div[5]")).Click();
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='center']/div/div[4]/div[3]/div/div/div[1]")).Click();


                    List<IWebElement> scrolls = driver.FindElements(By.XPath(".//*[@id='center']/div/div[4]/div[3]/div/div/div")).ToList();
                    Console.WriteLine(scrolls.Count);
                    int j = 1;
                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                    while (j < scrolls.Count)
                    {
                        if (j == (scrolls.Count) - 1)
                        {
                            Thread.Sleep(250);
                        }
                        js.ExecuteScript("arguments[0].scrollIntoView();", scrolls[j]);

                        scrolls = driver.FindElements(By.XPath(".//*[@id='center']/div/div[4]/div[3]/div/div/div")).ToList();
                        j = j + 1;
                    }

                }
                catch
                {


                }
                driver.FindElement(By.XPath(".//*[@id='center']//div[contains(.,\"" + sPplanProduct + "\")]/span")).Click();
                Thread.Sleep(10000);
                if (driver.FindElement(By.XPath(".//*[@id='center']//div/span[@class=\"total-valid-plan-week-quantity market-level preliminary normal-size\"]")).GetCssValue("Color") == "rgb(0, 0, 153)")
                {
                    Reporter.ReportEvent("PPlan Order Qty Color Verification", "PPlan Order Qty Color in Blue color when we echange in qty in HMOrder", HP.LFT.Report.Status.Passed);
                }
                else
                {

                    Reporter.ReportEvent("PPlan Order Qty Color Verification", "PPlan Order Qty Color is not showing in Blue color when we echange in qty in HMOrder", HP.LFT.Report.Status.Failed);
                }
                driver.Close();
            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Selenium_PPlan_EditQunatity_Colorvalidation script failed", "Selenium_PPlan_EditQunatity_Colorvalidation script failed", HP.LFT.Report.Status.Failed, ESSpath);
                driver.Close();
            }

        }
      //  [Test,Order(11)]
        public void TC11_Selenium_PPlan_Definite_Colorvalidation()
        {
            try {
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                string sSeason = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");
                string sSection = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Section");
                string sDept = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Dept");
                string sPplanProduct = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PPlanProduct");
                string tabName = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "TabName");
                Thread.Sleep(1000);
                pPlanPage = new PrePlanPage(driver);
                Thread.Sleep(1000);
                driver.Manage().Window.Maximize();
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_PPlan"));
                Thread.Sleep(1000);
                pPlanPage.waitForPageToLoad();
                pPlanPage.waitForSpinnerToDisappear();


                if (!pPlanPage.seasonButtonStatus.Text.Contains(sSeason))
                {
                    sGMethods.WebButton_Click(pPlanPage.get_seasonSelectionDirectiveButton(sSeason));
                }
                Reporter.ReportEvent("Season selected", "Season selected : " + sSeason, HP.LFT.Report.Status.Passed);
                pPlanPage.waitForPageToLoad();
                Assert.IsTrue(pPlanPage.get_ProductNameTableHeaderElement().Displayed);
                if (!pPlanPage.SectionDirectrivesDropDownElementSelectedElement.Text.Contains(sSection))
                {
                    sGMethods.SelectDropDownByValue(pPlanPage.SectionDirectrivesDropDownElement,
                        sSection);
                    Thread.Sleep(4000);
                    pPlanPage.waitForPageToLoad();
                }
                Reporter.ReportEvent("Directive selected from dropdown", "Directive selected from dropdown : " + sSection,
                    HP.LFT.Report.Status.Passed);
                try
                {
                    driver.FindElement(By.XPath("//button[text()='Full']")).Click();
                    Thread.Sleep(4000);
                    pPlanPage.waitForPageToLoad();
                    pPlanPage.waitForSpinnerToDisappear();
                }
                catch
                {
                }
                if (pPlanPage.get_departmentTabElement(tabName).Displayed)
                {
                    if (!pPlanPage.DepartmentTabActivElement.Text.Contains(tabName))
                    {
                        pPlanPage.get_departmentTabElement(tabName).Click();
                        Thread.Sleep(4000);
                        pPlanPage.waitForPageToLoad();
                        pPlanPage.waitForSpinnerToDisappear();
                    }
                }
                Reporter.ReportEvent("Tab selected", "Tab selected : " + tabName, HP.LFT.Report.Status.Passed);
                Thread.Sleep(4000);
                Assert.IsTrue(pPlanPage.get_ProductNameTableHeaderElement().Displayed);
                pPlanPage.waitForSpinnerToDisappear();
                driver.FindElement(By.XPath(".//*[@id='center']/div/div[4]/div[1]/div/div[1]/div[5]")).Click();

                try
                {
                    driver.FindElement(By.XPath(".//*[@id='center']/div/div[4]/div[3]/div/div/div[1]")).Click();


                    List<IWebElement> scrolls = driver.FindElements(By.XPath(".//*[@id='center']/div/div[4]/div[3]/div/div/div")).ToList();
                    Console.WriteLine(scrolls.Count);
                    int j = 1;
                    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                    while (j < scrolls.Count)
                    {
                        if (j == (scrolls.Count) - 1)
                        {
                            Thread.Sleep(250);
                        }
                        js.ExecuteScript("arguments[0].scrollIntoView();", scrolls[j]);

                        scrolls = driver.FindElements(By.XPath(".//*[@id='center']/div/div[4]/div[3]/div/div/div")).ToList();
                        j = j + 1;
                    }

                }
                catch
                {


                }
                driver.FindElement(By.XPath(".//*[@id='center']//div[contains(.,\"" + sPplanProduct + "\")]/span")).Click();
                Thread.Sleep(20000);
                //  if (driver.FindElement(By.XPath(".//*[@id='center']//div/span[@class=\"total-valid-plan-week-quantity market-level preliminary normal-size\"]")).GetCssValue("Color") == "rgb(0, 0, 0)")
                //  if (driver.FindElement(By.XPath("(.//*[@id='center']//span[contains(.,\"600 001\")])[2]")).GetCssValue("Color") == "rgb(0, 0, 0)")
                if (driver.FindElement(By.XPath("(.//*[@id='center']//div[contains(@class,\"cell-value definitive normal-size sum-week-quantity text-right readonly sum-week-column context-menu-area\")])[1]/span[2]")).GetCssValue("Color") == "rgb(0, 0, 0)")
                {
                    Reporter.ReportEvent("PPlan Order Qty Color Verification after Definite", "PPlan Order Qty Color in Black color when we change to Defnite the HMOrder", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("PPlan Order Qty Color Verification after Definite", "PPlan Order Qty Color is not showing in Black color when we change to Defnite the HMOrder", HP.LFT.Report.Status.Failed);
                }
                driver.Close();
            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Selenium_PPlan_Definite_Colorvalidation script failed", "Selenium_PPlan_Definite_Colorvalidation script failed", HP.LFT.Report.Status.Failed, ESSpath);
                driver.Close();
            }

        }
      //  [Test,Order(13)]
        public void TC13_Selenium_SmokeTest_HMORder_BOPOInternalOrderexportsVerification()
        {
            try {
                Thread.Sleep(3000);
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(3000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_ICC"));
                Thread.Sleep(2000);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                ICCBAMPageObj ICCBamobjs = new ICCBAMPageObj(driver);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Acc2010");
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCATENV"));
                }
                if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCSITENV"));
                }
                if (ConfigUtils.Read("ENV") == "DIT")
                {

                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Int2010");
                }
                Thread.Sleep(2000);
                sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Application, "Application DropDown", "BOPOInternalOrder_HMOrder");
                Thread.Sleep(2000);
                //  sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "ProductDevelopment");
                sGMethods.WebEdit_SetValue(driver.FindElement(By.XPath("//table[@class=\"searchFieldsTable\"]//td[contains(.,\"OrderNumber\")]/following-sibling::td/input")), "OrderNumber", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber"));

                sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                Thread.Sleep(10000);

                ICCBamobjs.HMOrder_BOPOInternalOrderExports_Verification();

                driver.Close();

                // ICCBamobjs.WebEdit_ProductID.GetCssValue("color");
            }
            catch
            {
                
                    Reporter.ReportEvent("Selenium_SmokeTest_HMORder_BOPOInternalOrderexportsVerification script failed", "Selenium_SmokeTest_HMORder_BOPOInternalOrderexportsVerification script failed", HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }
     //   [Test,Order(14)]
        public void TC14_Selenium_SmokeTest_HMORder_GenericOrderExports_Verification()
        {
            try {
                Thread.Sleep(3000);
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();

                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_ICC"));

                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                ICCBAMPageObj ICCBamobjs = new ICCBAMPageObj(driver);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Acc2010");
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCATENV"));
                }
                if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCSITENV"));
                }
                if (ConfigUtils.Read("ENV") == "DIT")
                {

                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Int2010");
                }
                Thread.Sleep(2000);
                sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Application, "Application DropDown", "GenericOrder_HMOrder");
                Thread.Sleep(2000);
                sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "GenericOrder");
                sGMethods.WebEdit_SetValue(driver.FindElement(By.XPath("//table[@class=\"searchFieldsTable\"]//td[contains(.,\"OrderNumber\")]/following-sibling::td/input")), "OrderNumber", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber"));

                sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                Thread.Sleep(10000);

                ICCBamobjs.HMOrder_GenericOrderExports_Verification();

                driver.Close();

                // ICCBamobjs.WebEdit_ProductID.GetCssValue("color");
            }
            catch
            {
               Reporter.ReportEvent("Selenium_SmokeTest_HMORder_GenericOrderExports_Verification script failed", "Selenium_SmokeTest_HMORder_GenericOrderExports_Verification script failed", HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }
       // [Test,Order(15)]
        public void TC15_Selenium_SmokeTest_HMORder_OrderedProductSpecification_Verification()
        {
            try {
                Thread.Sleep(3000);
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_ICC"));
                Thread.Sleep(1000);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                ICCBAMPageObj ICCBamobjs = new ICCBAMPageObj(driver);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Acc2010");
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCATENV"));
                }
                if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCSITENV"));
                }
                if (ConfigUtils.Read("ENV") == "DIT")
                {

                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Int2010");
                }
                Thread.Sleep(2000);
                sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Application, "Application DropDown", "OrderedProductSpecification_Castor");
                Thread.Sleep(2000);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "GenericOrder");
                sGMethods.WebEdit_SetValue(driver.FindElement(By.XPath("//table[@class=\"searchFieldsTable\"]//td[contains(.,\"ProductId\")]/following-sibling::td/input")), "ProductId", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PPlanNumber"));

                sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                Thread.Sleep(10000);

                ICCBamobjs.HMOrder_OrderedProductSpecification_Verification();

                driver.Close();

                // ICCBamobjs.WebEdit_ProductID.GetCssValue("color");
            }
            catch
            {
                
               Reporter.ReportEvent("Selenium_SmokeTest_HMORder_OrderedProductSpecification_Verification script failed", "Selenium_SmokeTest_HMORder_OrderedProductSpecification_Verification script failed", HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }
       // [Test]
        public void TC16_Selenium_Castor_CapacityBooking()
        {
            try {
              
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                List<string> lswins = driver.WindowHandles.ToList();
              
                Thread.Sleep(4000);
                sGMethods.GetLatestWindow(driver);
                //  sGMethods.AlertAccept(driver);
                //  Castorpages castorobjs = new Castorpages(driver);
                //  //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                //  Thread.Sleep(1000);

                ////  driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                //  Thread.Sleep(5000);

                //  sGMethods.GetLatestWindow(driver);
                //  string sparentwindow = sGMethods.GetCurrentWindow(driver);
                //  wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='AEFGlobalFullTextSearch']")));
                //  string xval = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID");
                //  //   castorobjs.PDMerch_SearchBox.SendKeys(xval);
                //  sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                //  Thread.Sleep(3000);
                //  driver.SwitchTo().DefaultContent();
                //  sGMethods.WaitFrametoSwitch(driver, "gsFrame");
                //  //            sGMethods.SwitchFrame_Name(driver, "gsFrame");
                //  sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                //  sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                //  castorobjs.PDMerch_EditClick_OfficeProduct("CNHK");
                //  Thread.Sleep(1000);
                //  driver.SwitchTo().DefaultContent();
                //  driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();

                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchPwd"));

                //   driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                Thread.Sleep(20000);


                string sparentwindow = sGMethods.GetCurrentWindow(driver);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='GlobalNewTEXT']")));

                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, "SearchBox Text", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();
                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }

                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }



                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                string offices = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office");
                string[] offciesarr = offices.Split(':');
                castorobjs.PDMerch_EditClick_OfficeProduct(offciesarr[1].Trim());

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);

                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_CapacityBooking, "tab_CapacityBooking");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"capacity-booking-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='toolbar-container']//span[contains(.,\"Create New Booking\")]")));
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_CreateNewBooking, "tab_CreateNewBooking");
                Thread.Sleep(1000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }

                sGMethods.WebButton_Click(castorobjs.PDMerch_Button_CreateCapacityBooking, "Create Capacity Booking");
                //  sGMethods.WaitAlerttoPresent(driver);
                try
                {
                    sGMethods.AlertAccept(driver);
                }
                catch
                {

                }
              
                //Thread.Sleep(1000);
                //HMOrderObject appModel = new HMOrderObject();

                //appModel.MozillaFirefoxWindow.AuthenticationRequiredDialog.Activate();
                //appModel.MozillaFirefoxWindow.AuthenticationRequiredDialog.SendKeys("tempsrmal");
                //appModel.MozillaFirefoxWindow.AuthenticationRequiredDialog.SendKeys(HP.LFT.SDK.Keys.Tab);
                //appModel.MozillaFirefoxWindow.AuthenticationRequiredDialog.SendKeys("Summer2017");
                //appModel.MozillaFirefoxWindow.AuthenticationRequiredDialog.SendKeys(HP.LFT.SDK.Keys.Return);

                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                wait.Until(ExpectedConditions.ElementExists(By.XPath("//div[contains(@data-bind,\"click: addBooking\")]/i")));
                sGMethods.WebLink_Click(castorobjs.PDMerch_Button_CapacityBooking_AddBooking, "Capacity Booking Add Booking");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.PDMerch_Button_CapacityBooking_Supplierspan, "Supplier span field");
                Thread.Sleep(1000);
                //  sGMethods.WebLink_Click(castorobjs.PDMerch_Button_CapacityBooking_Supplierspan, "Supplier span field");
                // sGMethods.SelectDropDownByValue(castorobjs.PDMerch_Button_CapacityBooking_SupplierSelection, "Supplier selection DropDown", "0338-Kei Lock Newage (HK) Limited");

                new SelectElement(castorobjs.PDMerch_Button_CapacityBooking_SupplierSelection).SelectByIndex(1);
                sGMethods.WebLink_Click(castorobjs.PDMerch_Button_CapacityBooking_OPD, "Capacity Booking OPD");
                sGMethods.WebButton_Click(castorobjs.PDMerch_Button_CapacityBooking_TodayDate, "Today date");

                sGMethods.WebLink_Click(castorobjs.PDMerch_Button_CapacityBooking_TOD, "Capacity Booking TOD");
                sGMethods.WebButton_Click(castorobjs.PDMerch_Button_CapacityBooking_TodayDate, "Today date");
                sGMethods.WebLink_Click(castorobjs.PDMerch_Button_CapacityBooking_ISW, "Capacity Booking ISW");
                sGMethods.WebButton_Click(castorobjs.PDMerch_Button_CapacityBooking_TodayDate, "Today date");

                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_Button_CapacityBooking_QTY, "1");
                castorobjs.PDMerch_Button_CapacityBooking_QTY.SendKeys(OpenQA.Selenium.Keys.Tab);
                castorobjs.PDMerch_Button_CapacityBooking_Channelselect.Click();
                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_Button_CapacityBooking_Channelselect, "Store");
                castorobjs.PDMerch_Button_CapacityBooking_BaseTrailelect.Click();
                sGMethods.SelectDropDownByValue(castorobjs.PDMerch_Button_CapacityBooking_BaseTrailelect, "Trail");
                Thread.Sleep(1000);
                //  sGMethods.WebLink_Click(castorobjs.PDMerch_Button_CapacityBooking_COP, "Supplier span field");
                // Thread.Sleep(1000);
                //  Keyboard.SendString(HP.LFT.SDK.Keys.Down);
                //  Keyboard.SendString(HP.LFT.SDK.Keys.Return);
                try
                {
                    sGMethods.selectByVisibleText(castorobjs.PDMerch_Button_CapacityBooking_COP, "China");
                }
                catch
                {

                }
              
              //  new SelectElement(castorobjs.PDMerch_Button_CapacityBooking_COP).SelectByIndex(1);
                Thread.Sleep(6000);
                //Keyboard.SendString(HP.LFT.SDK.Keys.Return);
               // Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.PDMerch_Button_CapacityBooking_SaveButton, "Capacity Save");
                //sGMethods.WaitAlerttoPresent(driver);
                Thread.Sleep(10000);
                sGMethods.AlertAccept(driver);
                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_CapacityBooking, "tab_CapacityBooking");
                Thread.Sleep(15000);
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_CapacityBooking, "tab_CapacityBooking");
                Thread.Sleep(5000);
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_CapacityBooking, "tab_CapacityBooking");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception:::::::::::" + e.Message);
                    Console.WriteLine("Exception:::::::::::" + e.StackTrace);
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"capacity-booking-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                string sbookingid = castorobjs.CapcityBooking_Verifyrecord(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductDevelopName"));

              //  xCellFileHelper.SetCellValueByRowAndColumn("Selenium_SmokeTest1", "BookingID", sbookingid);

                lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }


            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Selenium_Castor_CapacityBooking script failed", "Selenium_Castor_CapacityBooking script failed", HP.LFT.Report.Status.Failed, ESSpath);
                List<string> lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Assert.IsTrue(2 == 3);

            }
        }
      //  [Test]
        public void TC17_Selenium_SmokeTest_CapacityBooking_CastorExports_Verification()
        {
            try {
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_ICC"));
                Thread.Sleep(1000);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                ICCBAMPageObj ICCBamobjs = new ICCBAMPageObj(driver);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Acc2010");
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCATENV"));
                }
                if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCSITENV"));
                }
                if (ConfigUtils.Read("ENV") == "DIT")
                {

                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Int2010");
                }

                sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Application, "Application DropDown", "CapacityBooking_Bistro");
                Thread.Sleep(2000);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "GenericOrder");
               // sGMethods.WebEdit_SetValue(driver.FindElement(By.XPath(".//*[@id='main']//tr[contains(.,\"BookingId\")]//input")), "Booking iD", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "BookingID"));

                sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                Thread.Sleep(10000);

                ICCBamobjs.CapacityBooking_Castor_Verification();

                driver.Close();

                // ICCBamobjs.WebEdit_ProductID.GetCssValue("color");
            }
            catch
            {
                Reporter.ReportEvent("Selenium_SmokeTest_CapacityBooking_CastorExports_Verification script failed", "Selenium_SmokeTest_CapacityBooking_CastorExports_Verification script failed", HP.LFT.Report.Status.Failed);
                
                driver.Close();
            }
        }
       // [Test]
        public void TC18_Selenium_Castor_SampleOrder_CounterSample_Creation()
        {
            try
            {
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                List<string> lswins = driver.WindowHandles.ToList();

                //Thread.Sleep(4000);
                //sGMethods.GetLatestWindow(driver);
                //sGMethods.AlertAccept(driver);
                //Castorpages castorobjs = new Castorpages(driver);
                ////  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                //Thread.Sleep(1000);

                //driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                //Thread.Sleep(5000);

                //sGMethods.GetLatestWindow(driver);
                //string sparentwindow = sGMethods.GetCurrentWindow(driver);
                //wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='AEFGlobalFullTextSearch']")));
                //string xval = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID");
                ////   castorobjs.PDMerch_SearchBox.SendKeys(xval);
                //sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                //Thread.Sleep(3000);
                //driver.SwitchTo().DefaultContent();
                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");
                ////            sGMethods.SwitchFrame_Name(driver, "gsFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                //castorobjs.PDMerch_EditClick_OfficeProduct("CNHK");
                //Thread.Sleep(1000);
                //driver.SwitchTo().DefaultContent();
                //driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                //Thread.Sleep(1000);
                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchPwd"));

                //   driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                Thread.Sleep(20000);


                string sparentwindow = sGMethods.GetCurrentWindow(driver);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='GlobalNewTEXT']")));

                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, "SearchBox Text", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();
                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }

                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }



                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                string offices = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office");
                string[] offciesarr = offices.Split(':');
                castorobjs.PDMerch_EditClick_OfficeProduct(offciesarr[1].Trim());

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);

                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);
            
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_Samples, "tab_Samples");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {
                    
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"sample-orders-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='toolbar-container']/ul/li[contains(.,\"Order Sample\")]/img")));
                Thread.Sleep(3000);
                sGMethods.WebLink_Click(castorobjs.WebLink_Samples_OrderSample, "tab_OrderSample");
                Thread.Sleep(3000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='orderContext']")));
                Thread.Sleep(1000);
                string scurrentwindow = driver.CurrentWindowHandle;
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WebRadio_CreateSampleOrder_Order, "CreateSampleOrder_Order");
                Thread.Sleep(1000);

                sGMethods.SelectDropDownByValue(castorobjs.WebList_CreateSampleOrder_OrderList, "Order List", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber"));
                Thread.Sleep(1000);
                sGMethods.SelectDropDownByValue(castorobjs.WebList_CreateSampleOrder_SelectSampletype, "Sample Type", "Counter Sample");
                Thread.Sleep(1000);
                sGMethods.SelectDropDownByValue(castorobjs.WebList_CreateSampleOrder_LTCondition, "LT Condition", "1");
               
                try
                {
                    string sSize = "";
                    sSize = castorobjs.Webedit_CreateSampleOrder_Size.GetAttribute("placeholder");
                    sGMethods.WebEdit_SetValue(castorobjs.Webedit_CreateSampleOrder_Size, "Size field", sSize);
                }
                catch
                {
                   
                }
                
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.Webedit_CreateSampleOrder_Qty, "Quantity field", "1");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WedDate_CreateSampleOrder_toBeReceivedDate, "To Be Received Date field");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WedDate_CreateSampleOrder_ReceivedDateSelection, "To Be Received Date Selection table date");
                sGMethods.WebButton_Click(castorobjs.WebList_CreateSampleOrder_SendButton, "Send Button");
                Thread.Sleep(10000);
                Actions saction = new Actions(driver);
                saction.SendKeys(castorobjs.WebEdit_CreateSampleOrder_SendToField, OpenQA.Selenium.Keys.Enter);
                sGMethods.WebLink_Click(castorobjs.WebEdit_CreateSampleOrder_SendToField, "SendTo Field");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WebList_CreateSampleOrder_CreateNewcontactOption, "Create New Contact option");
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_CreateSampleOrder_CreateNewcontact_Emailaddress, "Email address field","hari.msgbox@gmail.com");
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_CreateSampleOrder_CreateNewcontact_Firstname, "Email address field", "Srihari");
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_CreateSampleOrder_CreateNewcontact_Lastname, "Email address field", "Mallineni");
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_CreateSampleOrder_CreateNewcontact_AddButton, "Add Button");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "ui-tinymce-1_ifr");
                sGMethods.WebEdit_SetValue(castorobjs.WebButton_CreateSampleOrder_CreateNewcontact_Message, "Message field", "Test");
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                IJavaScriptExecutor jse = (IJavaScriptExecutor)driver;

                jse.ExecuteScript("arguments[0].scrollIntoView(true);", driver.FindElement(By.XPath(".//*[@id='sendButtonId']")));
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_CreateSampleOrder_CreateNewcontact_SendButton, "Send Button");
                Thread.Sleep(3000);
                driver.SwitchTo().Window(lswins[(lswins.Count) - 1]).Close();
                Thread.Sleep(6000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                

            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Castor_SampleOrder_CounterSample_Creation script failed", "Castor_SampleOrder_CounterSample_Creation script failed", HP.LFT.Report.Status.Failed, ESSpath);
                List<string> lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Assert.IsTrue(2 == 3);

            }
        }
      //  [Test]
        public void TC19_Selenium_Castor_SampleOrder_ProductionSample_Creation()
        {
            try
            {
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(2000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(2000);
                List<string> lswins = driver.WindowHandles.ToList();

                //Thread.Sleep(4000);
                //sGMethods.GetLatestWindow(driver);
                //sGMethods.AlertAccept(driver);
                //Castorpages castorobjs = new Castorpages(driver);
                ////  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                //Thread.Sleep(1000);

                //driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                //Thread.Sleep(5000);

                //sGMethods.GetLatestWindow(driver);
                //string sparentwindow = sGMethods.GetCurrentWindow(driver);
                //wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='AEFGlobalFullTextSearch']")));
                //string xval = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID");
                ////   castorobjs.PDMerch_SearchBox.SendKeys(xval);
                //sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                //Thread.Sleep(3000);
                //driver.SwitchTo().DefaultContent();
                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");
                ////            sGMethods.SwitchFrame_Name(driver, "gsFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                //castorobjs.PDMerch_EditClick_OfficeProduct("CNHK");
                //Thread.Sleep(1000);
                //driver.SwitchTo().DefaultContent();
                //driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                //Thread.Sleep(1000);
                //sGMethods.GetLatestWindow(driver);
                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchPwd"));

                //   driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                Thread.Sleep(10000);


                string sparentwindow = sGMethods.GetCurrentWindow(driver);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='GlobalNewTEXT']")));
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, "SearchBox Text", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();
                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }

                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }



                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                string offices = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office");
                string[] offciesarr = offices.Split(':');
                castorobjs.PDMerch_EditClick_OfficeProduct(offciesarr[1].Trim());

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);

                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);

                driver.SwitchTo().DefaultContent();
               
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_Samples, "tab_Samples");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"sample-orders-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='toolbar-container']/ul/li[contains(.,\"Order Sample\")]/img")));
                Thread.Sleep(3000);
                sGMethods.WebLink_Click(castorobjs.WebLink_Samples_OrderSample, "tab_OrderSample");
                Thread.Sleep(3000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }

                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='orderContext']")));
                Thread.Sleep(1000);
                string scurrentwindow = driver.CurrentWindowHandle;
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WebRadio_CreateSampleOrder_Order, "CreateSampleOrder_Order");
                Thread.Sleep(1000);

                sGMethods.SelectDropDownByValue(castorobjs.WebList_CreateSampleOrder_OrderList, "Order List", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber"));
                Thread.Sleep(1000);
                sGMethods.SelectDropDownByValue(castorobjs.WebList_CreateSampleOrder_SelectSampletype, "Sample Type", "Production Sample");
                Thread.Sleep(1000);
                sGMethods.SelectDropDownByValue(castorobjs.WebList_CreateSampleOrder_LTCondition, "LT Condition", "1");

                try
                {
                    string sSize = "";
                    sSize = castorobjs.Webedit_CreateSampleOrder_Size.GetAttribute("placeholder");
                    sGMethods.WebEdit_SetValue(castorobjs.Webedit_CreateSampleOrder_Size, "Size field", sSize);
                }
                catch
                {

                }

                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.Webedit_CreateSampleOrder_Qty, "Quantity field", "1");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WedDate_CreateSampleOrder_toBeReceivedDate, "To Be Received Date field");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WedDate_CreateSampleOrder_ReceivedDateSelection, "To Be Received Date Selection table date");
                sGMethods.WebButton_Click(castorobjs.WebList_CreateSampleOrder_SendButton, "Send Button");
                Thread.Sleep(10000);
                Actions saction = new Actions(driver);
                saction.SendKeys(castorobjs.WebEdit_CreateSampleOrder_SendToField, OpenQA.Selenium.Keys.Enter);
                sGMethods.WebLink_Click(castorobjs.WebEdit_CreateSampleOrder_SendToField, "SendTo Field");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WebList_CreateSampleOrder_CreateNewcontactOption, "Create New Contact option");
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_CreateSampleOrder_CreateNewcontact_Emailaddress, "Email address field", "hari.msgbox@gmail.com");
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_CreateSampleOrder_CreateNewcontact_Firstname, "Email address field", "Srihari");
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_CreateSampleOrder_CreateNewcontact_Lastname, "Email address field", "Mallineni");
                Thread.Sleep(3000);
                sGMethods.WebButton_Click(castorobjs.WebButton_CreateSampleOrder_CreateNewcontact_AddButton, "Add Button");
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "ui-tinymce-1_ifr");
                sGMethods.WebEdit_SetValue(castorobjs.WebButton_CreateSampleOrder_CreateNewcontact_Message, "Message field", "Test");
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                IJavaScriptExecutor jse = (IJavaScriptExecutor)driver;

                jse.ExecuteScript("arguments[0].scrollIntoView(true);", driver.FindElement(By.XPath(".//*[@id='sendButtonId']")));
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_CreateSampleOrder_CreateNewcontact_SendButton, "Send Button");
                Thread.Sleep(5000);
                driver.SwitchTo().Window(lswins[(lswins.Count) - 1]).Close();
                Thread.Sleep(5000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);

                lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }

            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Castor_SampleOrder_ProductionSample_Creation script failed", "Castor_SampleOrder_ProductionSample_Creation script failed", HP.LFT.Report.Status.Failed, ESSpath);
                List<string> lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Assert.IsTrue(2 == 3);

            }
        }
      //  [Test]
        public void TC20_Selenium_Castor_SampleOrder_RegisterSample()
        {
            try
            {
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                List<string> lswins = driver.WindowHandles.ToList();

                Thread.Sleep(4000);
                //sGMethods.GetLatestWindow(driver);
                //sGMethods.AlertAccept(driver);
                //Castorpages castorobjs = new Castorpages(driver);
                ////  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                //Thread.Sleep(1000);

                //driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                //Thread.Sleep(5000);

                //sGMethods.GetLatestWindow(driver);
                //string sparentwindow = sGMethods.GetCurrentWindow(driver);
                //wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='AEFGlobalFullTextSearch']")));
                //string xval = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID");
                ////   castorobjs.PDMerch_SearchBox.SendKeys(xval);
                //sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                //Thread.Sleep(3000);
                //driver.SwitchTo().DefaultContent();
                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");
                ////            sGMethods.SwitchFrame_Name(driver, "gsFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                //castorobjs.PDMerch_EditClick_OfficeProduct("CNHK");
                //Thread.Sleep(1000);
                //driver.SwitchTo().DefaultContent();
                //driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                //Thread.Sleep(1000);
                //sGMethods.GetLatestWindow(driver);
                //driver.SwitchTo().DefaultContent();
                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchPwd"));

                //   driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                Thread.Sleep(20000);


                string sparentwindow = sGMethods.GetCurrentWindow(driver);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='GlobalNewTEXT']")));

                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, "SearchBox Text", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();

                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }
                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }



                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                string offices = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office");
                string[] offciesarr = offices.Split(':');
                castorobjs.PDMerch_EditClick_OfficeProduct(offciesarr[1].Trim());

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);

                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);

                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_Samples, "tab_Samples");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"received-samples-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='toolbar-container']/ul[@class=\"uip-toolbar-row\"]/li/span[text()=\"Sample\"]")));
                Thread.Sleep(3000);
                sGMethods.WebLink_Click(castorobjs.WebButton_ReceivedSample_Sample, "ReceivedSample_Sample");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WebButton_ReceivedSample_RegisterSample, "ReceivedSample_Sample_RegisterSample");
                Thread.Sleep(3000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }

                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='contextSupplier']//input[@value=\"Order\"]")));
               // string scurrentwindow = driver.CurrentWindowHandle;
                sGMethods.WebLink_Click(castorobjs.WebRadio_RegisterSample_OrderRadio, "RegisterSample_Order");
                Thread.Sleep(1000);

                sGMethods.SelectDropDownByValue(castorobjs.WebList_RegisterSample_OrderidList, "Order List", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber"));
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WebCheck_RegisterSample_AllCheckboxes, "RegisterSample_Sampletype all check boxes");
                sGMethods.WebLink_Click(castorobjs.WebLink_RegisterSample_MoveRighArrow, "RegisterSample_Moveright Arrow");
                sGMethods.WebButton_Click(castorobjs.WebLink_RegisterSample_RegisterButton, "RegisterSample_Registerbutton");
                Thread.Sleep(15000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }

                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"received-samples-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");

                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                if (castorobjs.WebLink_RegisterSample_Countersamplerecord.Displayed == true)
                {
                    Reporter.ReportEvent("RegisterSample_Countersamplerecord is Displayed in received mails", "RegisterSample_Countersamplerecord is Displayed in received mails", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("RegisterSample_Countersamplerecord is not Displayed in received mails", "RegisterSample_Countersamplerecord is not Displayed in received mails", HP.LFT.Report.Status.Failed);
                }

                if (castorobjs.WebLink_RegisterSample_Productionsamplerecord.Displayed == true)
                {
                    Reporter.ReportEvent("WebLink_RegisterSample_Productionsamplerecord is Displayed in received mails", "WebLink_RegisterSample_Productionsamplerecord is Displayed in received mails", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("WebLink_RegisterSample_Productionsamplerecord is not Displayed in received mails", "WebLink_RegisterSample_Productionsamplerecord is not Displayed in received mails", HP.LFT.Report.Status.Failed);
                }

                sGMethods.WebLink_Click(castorobjs.WebLink_RegisterSample_SelectAllCheckbox, "RegisterSample_SelectAll checkbox");
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"received-samples-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='toolbar-container']/ul[@class=\"uip-toolbar-row\"]/li/span[text()=\"Sample\"]")));
                Thread.Sleep(3000);
                sGMethods.WebLink_Click(castorobjs.WebButton_ReceivedSample_Sample, "ReceivedSample_Sample");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WebButton_ReceivedSample_HandOverToLab, "ReceivedSample_Sample_HandOvertoLab");
                Thread.Sleep(3000);
                try
                {
                    sGMethods.AlertAccept(driver);
                    Thread.Sleep(3000);
                }
                catch
                {

                }
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"received-samples-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                if (castorobjs.Tooltip_ReceivedSample_HandOverToLabrecord.Displayed == true)
                {
                    Reporter.ReportEvent("ReceivedSample_HandOverToLabrecord is Displayed in received mails", "ReceivedSample_HandOverToLabrecord is Displayed in received mails", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ReceivedSample_HandOverToLabrecord is not Displayed in received mails", "ReceivedSample_HandOverToLabrecord is not Displayed in received mails", HP.LFT.Report.Status.Failed);
                }

                lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                //Reporter.ReportEvent("Sample registration in the Received Email section Script failed", "Sample registration in the Received Email section Script failed", HP.LFT.Report.Status.Failed);
                Reporter.ReportEvent("Castor_SampleOrder_RegisterSample script failed", "Castor_SampleOrder_RegisterSample script failed", HP.LFT.Report.Status.Failed, ESSpath);
                List<string> lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Assert.IsTrue(2 == 3);
            }
        }
      //  [Test]
        public void TC21_Selenium_Castor_SampleOrder_LabTest()
        {
            try
            {
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                List<string> lswins = driver.WindowHandles.ToList();

                Thread.Sleep(4000);
                //sGMethods.GetLatestWindow(driver);
                //sGMethods.AlertAccept(driver);
                //Castorpages castorobjs = new Castorpages(driver);
                ////  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                //Thread.Sleep(1000);

                //driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                //Thread.Sleep(5000);

                //sGMethods.GetLatestWindow(driver);
                //string sparentwindow = sGMethods.GetCurrentWindow(driver);
                //wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='AEFGlobalFullTextSearch']")));
                //string xval = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID");
                ////   castorobjs.PDMerch_SearchBox.SendKeys(xval);
                //sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                //Thread.Sleep(3000);
                //driver.SwitchTo().DefaultContent();
                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");
                ////            sGMethods.SwitchFrame_Name(driver, "gsFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                //castorobjs.PDMerch_EditClick_OfficeProduct("CNHK");
                //Thread.Sleep(1000);
                //driver.SwitchTo().DefaultContent();
                //driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                //Thread.Sleep(1000);
                //sGMethods.GetLatestWindow(driver);
                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAUSER"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAPWD"));

                //   driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                Thread.Sleep(20000);


                string sparentwindow = sGMethods.GetCurrentWindow(driver);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='GlobalNewTEXT']")));

                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, "SearchBox Text", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();
                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }

                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }



                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                string offices = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office");
                string[] offciesarr = offices.Split(':');
                castorobjs.PDMerch_EditClick_OfficeProduct(offciesarr[1].Trim());

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);

                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);

                driver.SwitchTo().DefaultContent();

                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WebLink_Click(castorobjs.Tooltab_QATAB, "Quality Assurance TAB");
                Thread.Sleep(5000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//a[text()=\"Samples Ready for Test\"]")));
                sGMethods.WebLink_Click(castorobjs.Tooltab_SampleReadyTestTab, "Sample test ready TAB");
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");

                sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMQualityAssuranceSamplesReadyForTestcontentFrame");
                sGMethods.SwitchFrame_Name(driver, "tvcInlineFrame_0");
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_Search_Ordernumber,"Order Number", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber"));
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_Search_Apply, "Apply button");
                Thread.Sleep(30000);
                sGMethods.WebButton_Click(castorobjs.WebButton_Search_Apply, "Apply button");
                Thread.Sleep(15000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");
                sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMQualityAssuranceSamplesReadyForTestcontentFrame");
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                Thread.Sleep(1000);
                List<IWebElement> lsSampleorders = driver.FindElements(By.XPath(".//table[@class=\"contentTable\"]/tbody/tr/td[1]")).ToList();
                int samplecount = lsSampleorders.Count();
                for (int i=0;i< samplecount; ++i)
                {
                    Thread.Sleep(1000);
                    lsSampleorders = driver.FindElements(By.XPath(".//table[@class=\"contentTable\"]/tbody/tr/td[1]")).ToList();
                    lsSampleorders[i].Click();
                    Thread.Sleep(3000);
                    driver.SwitchTo().DefaultContent();
                    sGMethods.WaitFrametoSwitch(driver, "content");
                    sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMQualityAssuranceSamplesReadyForTestcontentFrame");

                    wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='s2id_catSel']/a")));
                    Thread.Sleep(1000);
                    sGMethods.WebButton_Click(castorobjs.WebList_LabTest_Test, "LabTest Type");
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath(".//*[@id='select2-drop']//div[contains(.,\"Chlorophenols\")]")).Click();
                    Thread.Sleep(1000);
                    sGMethods.WebButton_Click(castorobjs.WebList_LabTest_Performer, "Performer field");
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath(".//*[@id='select2-drop']//div[contains(.,\"H&M Internal\")]")).Click();
                    Thread.Sleep(1000);
                    sGMethods.WebButton_Click(castorobjs.WebButton_LabTest_InitiateTest, "Initiate Test");
                    Thread.Sleep(3000);
                    int srecord = i + 1;
                    if(driver.FindElement(By.XPath("//tbody/tr/td/div[contains(.,\"Initiated\")]")).Displayed== true)
                    {

                        Reporter.ReportEvent("Lab test intiated for record "+ srecord, "Lab test intiated for record " + srecord, HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("Lab test is not intiated for record " + srecord, "Lab test is not intiated for record " + srecord, HP.LFT.Report.Status.Failed);
                    }

                    sGMethods.WebButton_Click(castorobjs.WebButton_LabTest_HandleQatest, "Handle QA Test");
                    Thread.Sleep(3000);
                    driver.SwitchTo().DefaultContent();
                    sGMethods.WaitFrametoSwitch(driver, "content");
                    sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMQualityAssuranceQATestcontentFrame");
                    sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                    sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                    sGMethods.WebButton_Click(castorobjs.WebButton_QATest_SelectAllCheckbox, "Select All check box in QA test");
                        Thread.Sleep(1000);
                    driver.SwitchTo().DefaultContent();
                    sGMethods.WaitFrametoSwitch(driver, "content");
                    sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMQualityAssuranceQATestcontentFrame");
                    sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                    sGMethods.WebButton_Click(castorobjs.WebButton_QATest_Enterresult, "Enter results for QA Test");
                    Thread.Sleep(1000);
                    driver.SwitchTo().DefaultContent();
                    sGMethods.WaitFrametoSwitch(driver, "content");
                    sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMQualityAssuranceQATestcontentFrame");
                    sGMethods.WebButton_Click(castorobjs.WebButton_QATestresult_Pass, "QA Test result pass selection");
                    Thread.Sleep(2000);
                    sGMethods.WebButton_Click(castorobjs.WebButton_QATestresult_Confirm, "QA Test result Confirm Button");
                    Thread.Sleep(2000);
                    sGMethods.WebButton_Click(castorobjs.WebButton_QATestresult_BacktoQAList, "Back to QA List");
                    Thread.Sleep(2000);
                    driver.SwitchTo().DefaultContent();
                    sGMethods.WaitFrametoSwitch(driver, "content");
                    Thread.Sleep(2000);
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//a[text()=\"Samples Ready for Test\"]")));
                   sGMethods.WebLink_Click(castorobjs.Tooltab_SampleReadyTestTab, "Sample test ready TAB");
                    Thread.Sleep(5000);
                    driver.SwitchTo().DefaultContent();
                    sGMethods.WaitFrametoSwitch(driver, "content");
                    sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMQualityAssuranceSamplesReadyForTestcontentFrame");
                    sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                    sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                }

                lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Castor_SampleOrder_LabTest script failed", "Castor_SampleOrder_LabTest script failed", HP.LFT.Report.Status.Failed, ESSpath);
                List<string> lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Assert.IsTrue(2 == 3);
            }
        }
      //  [Test]
        public void TC22_Selenium_Castor_SampleOrder_SampleReportgeneration()
        {
            try
            {
                Thread.Sleep(1000);
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(2000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                List<string> lswins = driver.WindowHandles.ToList();

                Thread.Sleep(4000);
                //sGMethods.GetLatestWindow(driver);
                //sGMethods.AlertAccept(driver);
                //Castorpages castorobjs = new Castorpages(driver);
                ////  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                //Thread.Sleep(1000);

                //driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                //Thread.Sleep(5000);

                //sGMethods.GetLatestWindow(driver);
                //string sparentwindow = sGMethods.GetCurrentWindow(driver);
                //wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='AEFGlobalFullTextSearch']")));
                //string xval = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID");
                ////   castorobjs.PDMerch_SearchBox.SendKeys(xval);
                //sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                //Thread.Sleep(3000);
                //driver.SwitchTo().DefaultContent();
                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");
                ////            sGMethods.SwitchFrame_Name(driver, "gsFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                //castorobjs.PDMerch_EditClick_OfficeProduct("CNHK");
                //Thread.Sleep(1000);
                //driver.SwitchTo().DefaultContent();
                //driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                //Thread.Sleep(1000);
                //sGMethods.GetLatestWindow(driver);
                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                 castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchPwd"));
               // castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAUSER"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDQAPWD"));
                //   driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                Thread.Sleep(20000);


                string sparentwindow = sGMethods.GetCurrentWindow(driver);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='GlobalNewTEXT']")));

                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, "SearchBox Text", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();

                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }
                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }



                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                string offices = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office");
                string[] offciesarr = offices.Split(':');
                castorobjs.PDMerch_EditClick_OfficeProduct(offciesarr[1].Trim());

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);

                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);

                driver.SwitchTo().DefaultContent();

                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_Samples, "tab_Samples");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"received-samples-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='toolbar-container']/ul[@class=\"uip-toolbar-row\"]/li/span[text()=\"Sample\"]")));
                Thread.Sleep(3000);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                sGMethods.WebLink_Click(castorobjs.PDMerch_ReceivedSample_FirstCheckBox, "PDMerch_ReceivedSample_FirstCheckBox");
                Thread.Sleep(1000);
                //  IJavaScriptExecutor jse = (IJavaScriptExecutor)driver;

                //  jse.ExecuteScript("arguments[0].scrollIntoView(true);", castorobjs.PDMerch_ReceivedSample_FirstCheckBox);
                castorobjs.PDMerch_ReceivedSample_FirstCheckBox.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                castorobjs.PDMerch_ReceivedSample_FirstCheckBox.SendKeys(OpenQA.Selenium.Keys.ArrowUp);
                castorobjs.PDMerch_ReceivedSample_FirstCheckBox.SendKeys(OpenQA.Selenium.Keys.ArrowUp);

                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
               // sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_Samples, "tab_Samples");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//iframe[@data-config-id=\"received-samples-list\"]")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//div[@class=\"gwt-technia-DashboardGadgetContent\"]//iframe")));
                }
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                sGMethods.WebLink_Click(castorobjs.WebButton_ReceivedSample_Sample, "ReceivedSample_Sample");
                Thread.Sleep(1000);

                sGMethods.WebLink_Click(castorobjs.WebButton_ReceivedSample_CreateSampleReport, "WebButton_ReceivedSample_CreateSampleReport");
                Thread.Sleep(5000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='buttons']/input[@value=\"Next >>\"]")));
                sGMethods.WebButton_Click(castorobjs.WebButton_Samplereport_Next, "Sample Report Next button");

                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_Samplereport_Next, "Sample Report Next button");
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_Samplereport_Next, "Sample Report Next button");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WebRadio_Samplereport_Approve, "Sample Report Approved Radio");
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_Samplereport_Next, "Sample Report Next button");
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_Samplereport_Send, "Sample Report Send button");
                Thread.Sleep(1000);
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.WebEdit_CreateSampleOrder_SendToField, "SendTo Field");
                sGMethods.WebLink_Click(castorobjs.WebList_CreateSampleOrder_CreateNewcontactOption, "Create New Contact option");
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_CreateSampleOrder_CreateNewcontact_Emailaddress, "Email address field", "hari.msgbox@gmail.com");
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_CreateSampleOrder_CreateNewcontact_Firstname, "Email address field", "Srihari");
                Thread.Sleep(1000);
                sGMethods.WebEdit_SetValue(castorobjs.WebEdit_CreateSampleOrder_CreateNewcontact_Lastname, "Email address field", "Mallineni");
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_CreateSampleOrder_CreateNewcontact_AddButton, "Add Button");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "ui-tinymce-1_ifr");
                sGMethods.WebEdit_SetValue(castorobjs.WebButton_CreateSampleOrder_CreateNewcontact_Message, "Message field", "Test");
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                IJavaScriptExecutor jse = (IJavaScriptExecutor)driver;

                jse.ExecuteScript("arguments[0].scrollIntoView(true);", driver.FindElement(By.XPath(".//*[@id='sendButtonId']")));
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.WebButton_CreateSampleOrder_CreateNewcontact_SendButton, "Send Button");
                Thread.Sleep(3000);
                driver.SwitchTo().Window(lswins[(lswins.Count) - 1]).Close();
                Thread.Sleep(5000);
                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);

                lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }

            }
            catch
            {
                Reporter.ReportEvent("Castor_SampleOrder_SampleReportgeneration script failed", "Castor_SampleOrder_SampleReportgeneration script failed", HP.LFT.Report.Status.Failed);
                List<string> lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
            }
        }
      //  [Test]
        public void TC23_Selenium_Castor_AssignKoreaCertificate()
        {
            try
            {
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(3000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(3000);
                List<string> lswins = driver.WindowHandles.ToList();

                Thread.Sleep(4000);
                //sGMethods.GetLatestWindow(driver);
                //sGMethods.AlertAccept(driver);
                //Castorpages castorobjs = new Castorpages(driver);
                ////  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                //Thread.Sleep(1000);

                //driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                //Thread.Sleep(5000);

                //sGMethods.GetLatestWindow(driver);
                //string sparentwindow = sGMethods.GetCurrentWindow(driver);
                //wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='AEFGlobalFullTextSearch']")));
                //string xval = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID");
                ////   castorobjs.PDMerch_SearchBox.SendKeys(xval);
                //sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                //Thread.Sleep(3000);
                //driver.SwitchTo().DefaultContent();
                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");
                ////            sGMethods.SwitchFrame_Name(driver, "gsFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                //sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                //castorobjs.PDMerch_EditClick_OfficeProduct("CNHK");
                //Thread.Sleep(1000);
                //driver.SwitchTo().DefaultContent();
                //driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                //Thread.Sleep(1000);
                //sGMethods.GetLatestWindow(driver);
                sGMethods.GetLatestWindow(driver);

                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchPwd"));

                //   driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                Thread.Sleep(20000);


                string sparentwindow = sGMethods.GetCurrentWindow(driver);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='GlobalNewTEXT']")));

                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, "SearchBox Text", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();

                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }
                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }



                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                string offices = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office");
                string[] offciesarr = offices.Split(':');
                castorobjs.PDMerch_EditClick_OfficeProduct(offciesarr[1].Trim());

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);
                // driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);

                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_Certificates, "TABS_Certificates");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }

                sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMProductDevelopmentKoreaCertificatescontentFrame");
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='toolbar-container']//span[contains(.,\"Show Options\")]")));
                Thread.Sleep(5000);
                sGMethods.WebLink_Click(castorobjs.PDMerch_Certificates_ShowOptions, "TABS_Certificates_ShowOptions");
                Thread.Sleep(10000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }

                sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMProductDevelopmentKoreaCertificatescontentFrame");
                sGMethods.SwitchFrame_Name(driver, "tvcInlineFrame_0");
                sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMKoreaCertificateMyOptionsTabcontentFrame");
                sGMethods.WebLink_Click(castorobjs.PDMerch_Certificates_AssignCertificates, "PDMerch_Certificates_AssignCertificates");
                Thread.Sleep(20000);

                sGMethods.GetLatestWindow(driver);
                driver.SwitchTo().DefaultContent();
                /// sGMethods.SwitchFrame_Name(driver, "content");

                sGMethods.SwitchFrame_Name(driver, "source");
                sGMethods.SwitchFrame_Name(driver, "iSource");
                sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMKoreaCertificatePowerViewLibraryTabcontentFrame");
                sGMethods.SwitchFrame_Name(driver, "tvcInlineFrame_0");
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.PDMerch_AssignCertificates_ApplyButton, "PDMerch_AssignCertificates_ApplyButton");
                Thread.Sleep(20000);
               
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "source");
                sGMethods.SwitchFrame_Name(driver, "iSource");

                sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMKoreaCertificatePowerViewLibraryTabcontentFrame");
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                Thread.Sleep(10000);
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//table[@class=\"contentTable\"]//tr[1]/td[1]/input[1]")));
                Thread.Sleep(1000);
                sGMethods.WebButton_Click(castorobjs.PDMerch_AssignCertificates_FirstCheckBox, "PDMerch_AssignCertificates_FirstCheckBox");
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "source");
                sGMethods.SwitchFrame_Name(driver, "iSource");

                sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMKoreaCertificatePowerViewLibraryTabcontentFrame");
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                sGMethods.WebButton_Click(castorobjs.PDMerch_AssignCertificates_AssignButton, "PDMerch_AssignCertificates_AssignButton");


                lswins = driver.WindowHandles.ToList();

                driver.SwitchTo().Window(lswins[lswins.Count()-1]);
                driver.Close();
                Thread.Sleep(5000);
                sGMethods.GetLatestWindow(driver);
                driver.SwitchTo().DefaultContent();
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));
                sGMethods.WebLink_Click(castorobjs.PDMerch_TABS_Certificates, "TABS_Certificates");
                Thread.Sleep(3000);
                driver.SwitchTo().DefaultContent();
                wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt("content"));
                try
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }
                catch (Exception e)
                {

                    driver.SwitchTo().Frame(driver.FindElement(By.XPath("//*[@class='widget-inner headerless']/iframe")));
                }

                sGMethods.SwitchFrame_Name(driver, "tvcTabs0_HMProductDevelopmentKoreaCertificatescontentFrame");
                sGMethods.SwitchFrame_Name(driver, "tableContentFrame");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='toolbar-container']//span[contains(.,\"Show Options\")]")));
                Thread.Sleep(1000);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");
                if (driver.FindElement(By.XPath("//table[@class=\"contentTable\"]/tbody/tr[1]")).Displayed == true)
                {
                    Reporter.ReportEvent("Certificate displayed in the Certificate tab", "Certificate displayed in the Certificate tab", HP.LFT.Report.Status.Passed);
                }

                lswins = driver.WindowHandles.ToList();
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Selenium_Castor_AssignKoreaCertificate script failed", "Selenium_Castor_AssignKoreaCertificate script failed", HP.LFT.Report.Status.Failed, ESSpath);
                List<string> lswins = driver.WindowHandles.ToList();

                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);
                    driver.Close();

                }
                Assert.IsTrue(2 == 3);
            }
        }
       // [Test]
        public void TC24_Selenium_ICCExports_CountryCertificates()
        {
            try
            {
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(2000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_ICC"));
                Thread.Sleep(5000);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                ICCBAMPageObj ICCBamobjs = new ICCBAMPageObj(driver);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Acc2010");
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCATENV"));
                }
                if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCSITENV"));
                }
                if (ConfigUtils.Read("ENV") == "DIT")
                {

                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Int2010");
                }
                /// sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "BopoAcc2010");
                sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Application, "CountryCertificate_Castor");
                Thread.Sleep(2000);
               // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "ProductDevelopment");
                sGMethods.WebEdit_SetValue(ICCBamobjs.WebEdit_ProductID, "Product Name", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));

                sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                //  Thread.Sleep(30000);
                ICCBamobjs.CountryCertificate_Castor_Verification();

                driver.Close();

                // ICCBamobjs.WebEdit_ProductID.GetCssValue("color");
            }
            catch
            {
                Reporter.ReportEvent("Selenium_ICCExportsverificationforHandOver Script failed", "Selenium_ICCExportsverificationforHandOver Script failed", HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }

        public void TC23_2_Selenium_Castor_Merch_MChart()
        {
            try
            {
                //GetConfigurationAssembly.
                Thread.Sleep(1000);
                driver.Manage().Window.Maximize();
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(5000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_Castor"));
                Thread.Sleep(1000);
                //  sGMethods.WaitAlerttoPresent(driver);
                //  sGMethods.AlertAccept(driver);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                Castorpages castorobjs = new Castorpages(driver);
                //  sGMethods.WebButton_Click(castorobjs.WebButton_SEBOPDBUYER,"SEBOPDBUYER");
                castorobjs.CastorLogin(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchUser"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PDMerchPwd"));

                //   driver.FindElement(By.XPath(".//*[@id='divLogin']//input[@value=\"CNHK PD Merch\"]")).Click();

                Thread.Sleep(20000);


                string sparentwindow = sGMethods.GetCurrentWindow(driver);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='GlobalNewTEXT']")));
                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WebEdit_SetValue(castorobjs.PDMerch_SearchBox, "SearchBox Text", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID"));
                driver.FindElement(By.XPath("//a[contains(@href,'isSearchEnterKeyPressed()')]")).Click();
                Thread.Sleep(1000);
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                }
                catch
                {

                }

                Thread.Sleep(2000);
                driver.SwitchTo().DefaultContent();


                //sGMethods.WaitFrametoSwitch(driver, "gsFrame");

                string parentwindow = driver.CurrentWindowHandle;

                foreach (string win in driver.WindowHandles)
                {
                    driver.SwitchTo().Window(win);
                    // System.Diagnostics.Debug.WriteLine(driver.Title);
                }



                sGMethods.SwitchFrame_Name(driver, "structure_browser");
                driver.SwitchTo().Frame(1);
                sGMethods.SwitchFrame_Name(driver, "tableBodyRight");

                //driver.FindElement(By.XPath("(//*[@class='contentTable']//a[@targetlocation='popup'])[1]")).Click();

                string offices = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Office");
                string[] offciesarr = offices.Split(':');
                castorobjs.PDMerch_EditClick_OfficeProduct(offciesarr[1].Trim());

                // driver.SwitchTo().DefaultContent();
                driver.Close();
                driver.SwitchTo().Window(parentwindow);
                // driver.FindElement(By.XPath(".//*[@id='closeTA']")).Click();
                Thread.Sleep(1000);
                sGMethods.GetLatestWindow(driver);
                sGMethods.SwitchFrame_Name(driver, "content");
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//label[contains(.,\"ID\")]/following-sibling::div")));


                sGMethods.WebLink_Click(castorobjs.Tab_Description, "Dscription tab");
                Thread.Sleep(2000);
                string sProductionType = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductType");
                // driver.FindElement(By.XPath(".//*[@id='Appearances']//a/span[contains(.,\"Create Appearance(s)\")]")).Click();
                sGMethods.WebLink_Click(castorobjs.Weblink_FitDescription, "Fit Discription");
                Thread.Sleep(1000);
                sGMethods.WebLink_Click(castorobjs.Weblink_Mchart, "Weblink_Mchart");
                Thread.Sleep(2000);
                sGMethods.GetLatestWindow(driver);
                ////castorobjs.addSupplierDetails(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductionGroup"), xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductionType"));
                castorobjs.addMchartDetails(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season"), sProductionType);

                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='mx_btn-search']")));
                sGMethods.WebButton_Click(castorobjs.PDMerch_Search_AddSupplier, "Search button on Add Mchart window");
                //Thread.Sleep(5000);
                //sGMethods.WebButton_Click(castorobjs.PDMerch_miniClose_AddSupplieroptions, "MiniClose button for Add Supplier Option window");
                Thread.Sleep(50000);
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame("structure_browser");
                driver.SwitchTo().Frame("tableContentFrame");
                driver.SwitchTo().Frame("tableBodyRight");
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//table/tbody/tr/td[1]/input[@type=\"checkbox\"])[1]")));
                Thread.Sleep(5000);
                sGMethods.WebLink_Click(castorobjs.WebCheckBox_AddLabel_SearchResult1, "Add Mchart search Checkbox1");
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame("structure_browser");
                driver.SwitchTo().Frame("tableContentFrame");
                sGMethods.WebButton_Click(castorobjs.WebButton_AddLabel_AddSelected, "AddSelected button");
                Thread.Sleep(1000);
                lswins = driver.WindowHandles.ToList();
                int iswin = 1;
                while ((lswins.Count == 2) && (iswin <= 3000))
                {
                    lswins = driver.WindowHandles.ToList();
                    iswin = iswin + 1;
                    Thread.Sleep(1000);
                }


                Thread.Sleep(2000);
                foreach (string swin in lswins)
                {
                    driver.SwitchTo().Window(swin);

                }
                Thread.Sleep(1000);
                driver.SwitchTo().DefaultContent();
                sGMethods.WaitFrametoSwitch(driver, "content");

                if (driver.FindElement(By.XPath(".//div[@data-id=\"Attachments\"]//div/table//td[contains(.,\"Measurement Chart\")]")).Displayed == true)
                {
                    Reporter.ReportEvent("Selenium_Castor_Merch_Mchart displayed in page", "Selenium_Castor_Merch_Mchart displayed in page", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("Selenium_Castor_Merch_Mchart is not displayed in page", "Selenium_Castor_Merch_Mchart is not displayed in page", HP.LFT.Report.Status.Failed);
                }
                driver.Close();
            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Selenium_Castor_Merch_Mchart Script failed", "Selenium_Castor_Merch_Mchart Script failed", HP.LFT.Report.Status.Failed, ESSpath);
                driver.Close();

            }
        }
    


        public void TC23_2_Selenium_SmokeTest_MchartICCexportsVerification()
        {
            try
            {
                //GetConfigurationAssembly.
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                GeneralMethods sGMethods = new GeneralMethods();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_ICC"));
                Thread.Sleep(1000);
                // IAlert ialt = driver.SwitchTo().Alert();
                //ialt.Accept();

                //  driver.Url = "http://castor-sit.hm.com/enovia/emxLogin.jsp?portal=true&fromLoginPage=true";
                List<string> lswins = driver.WindowHandles.ToList();
                //foreach (string swin in lswins)
                //{
                //    driver.SwitchTo().Window(swin);

                //}

                sGMethods.GetLatestWindow(driver);

                ICCBAMPageObj ICCBamobjs = new ICCBAMPageObj(driver);
                // sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Acc2010");
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCATENV"));
                }
                if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", ConfigUtils.Read("ICCSITENV"));
                }
                if (ConfigUtils.Read("ENV") == "DIT")
                {

                    sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Env, "Environment Dropdown", "Int2010");
                }
                Thread.Sleep(2000);
                sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_Application, "Application DropDown", "MeasurementChart_Castor");
                Thread.Sleep(2000);
                //  sGMethods.SelectDropDownByValue(ICCBamobjs.Dropdown_BusinessObject, "Business Object dropdown", "ProductDevelopment");
               // sGMethods.WebEdit_SetValue(driver.FindElement(By.XPath("//table[@class=\"searchFieldsTable\"]//td[contains(.,\"OrderNumber\")]/following-sibling::td/input")), "OrderNumber", xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber"));

                sGMethods.WebButton_Click(ICCBamobjs.WebButton_LoadBAM);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Apply Filter/Sort\")]")));
                Thread.Sleep(10000);

                ICCBamobjs.Mchart_ExportElementExistinTimeRange();

                driver.Close();

                // ICCBamobjs.WebEdit_ProductID.GetCssValue("color");
            }
            catch
            {

                Reporter.ReportEvent("Selenium_SmokeTest_HMORder_BOPOInternalOrderexportsVerification script failed", "Selenium_SmokeTest_HMORder_BOPOInternalOrderexportsVerification script failed", HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }

        public void TC23_3_Selenium_VerifyCITStatus()
        {
            try
            {
                OFUPageObjects OFUObjects = new OFUPageObjects(driver);
                ICCBAMPageObj ICCBAMObjects = new ICCBAMPageObj(driver);
                GeneralMethods generalMethods = new GeneralMethods();

                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                string season = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");
                string orderNumber = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber");
                string OFU_Url = ConfigUtils.Read("URL_OFU");

                OFU_Url = OFU_Url + orderNumber + "-" + season;

                driver.Navigate().GoToUrl(OFU_Url);

                Thread.Sleep(5000);

                generalMethods.WebLink_Click(OFUObjects.CompositionLink);
                Thread.Sleep(2000);

                generalMethods.GetLatestWindow(driver);

                Assert.AreEqual("Not secured", OFUObjects.GetCompositionStatus());

                OFUObjects.EditComposition.Click();
                Thread.Sleep(2000);

                Assert.True(OFUObjects.VerifyBOM());

                generalMethods.SelectDropDownByValue(OFUObjects.ExclusiveOfDropDown, "Exclusive Of Dropdown", "EXCLUSIVE OF LACE");
                Thread.Sleep(2000);

                generalMethods.WebButton_Click(OFUObjects.SaveButton);
                Thread.Sleep(2000);

                Assert.AreEqual("Pending", OFUObjects.GetCompositionStatus());

                generalMethods.WebButton_Click(OFUObjects.SecureButton);

                Assert.AreEqual("Secured", OFUObjects.GetCompositionStatus().Substring(0, 7));

                //Close the Secure Country Specific Composition Page
                driver.Close();

                generalMethods.GetLatestWindow(driver);
                driver.Navigate().Refresh();
                Assert.True(OFUObjects.VerifyCompositionTabBackgroundColour());

                ICCBAMObjects.LaunchICCWindow();

                Assert.True(ICCBAMObjects.VerifyCITStatusExport(orderNumber));
                Assert.True(ICCBAMObjects.VerifyDetailedCompositionExport(orderNumber));

                driver.Close();
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("Selenium_VerifyCITStatus Script failed", e.Message, HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }

        public void TC23_4_Selenium_VerifyTagsPTAndITPush()
        {
            try
            {
                GeneralMethods generalMethods = new GeneralMethods();
                TagsPageObjects tagsObjects = new TagsPageObjects(driver);

                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                string orderNumber = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber");

                tagsObjects.NavigateAndSearchOrder(orderNumber);
                Assert.True(tagsObjects.PTResend());
                tagsObjects.NavigateAndSearchOrder(orderNumber);
                Assert.True(tagsObjects.ITResend());

                driver.Close();
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("TC23_4_Selenium_VerifyTagsPTAndITPush Script failed", e.Message, HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }

        public void TC29_Selenium_OFULogin()
        {
            try
            {
                OFUPageObjects OFUObjects = new OFUPageObjects(driver);
              
                GeneralMethods generalMethods = new GeneralMethods();

                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                string sEmail = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "SingleSignOn_Email");
               
                string OFU_Url = ConfigUtils.Read("URL_OFU3");
                driver.Navigate().GoToUrl(OFU_Url);

                Thread.Sleep(5000);
                OFUObjects.SingleSignOn(sEmail);
                Thread.Sleep(5000);
                FirefoxOptions ffoption = new FirefoxOptions();
               // ffoption.AddAdditionalCapability()
              //  driver.FindElement(By.XPath(".//*[@id='PRODUCT_LIST']")).Click();
                Thread.Sleep(5000);
                driver.Close();
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("TC2_Selenium_OFULogin Script failed", e.Message, HP.LFT.Report.Status.Failed);
                driver.Close();
            }
        }

        public void TearDown()
        {

            // Clean up after each test
            try
            {
                Thread.Sleep(2000);
                xCellFileHelper.CleanUp();
                Thread.Sleep(5000);
            }
            catch
            {

            }
            try
            {
                driver.Quit();
            }
            catch
            {

            }
          
            Thread.Sleep(5000);
        }


        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
        }
    }
}

