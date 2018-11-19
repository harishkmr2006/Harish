using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using HP.LFT.Report;
using InteropLibrary;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System.Drawing.Imaging;

namespace SITSmokeTests
{
    class RebuildTest
    {
        IWebDriver driver;
        WebDriverWait wait;
        ExcelHelper xCellFileHelper;
        string datafilePath;
        GeneralMethods sGMethods;

        private PrePlanPage pPlanPage;

        //  private string currentWindow;
        private ICCBAMPageObj iccPortal;

        private MOnitorUI intigrationmon;

        //  private string productNumber = "0553821";
        private string productNumber;

        private string orderNumber;

        public RebuildTest()
        {
            var binary = new FirefoxBinary(ConfigUtils.Read("FirefoxPath"));
            // string path = ReadFirefoxProfile();
            FirefoxProfile ffprofile = new FirefoxProfile();
            driver = new FirefoxDriver(binary, ffprofile);
            //  driver = new InternetExplorerDriver(internetExplorerDriverServerDirectory: "\\srv10177\\TEMPSHTIW$\\Desktop\\SIT\\SITSmokeTests\\IEDriverServer.exe"); 
            driver.Manage().Cookies.DeleteAllCookies();
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            // Before each test
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            datafilePath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + ConfigUtils.Read("TestDataPath");
            sGMethods = new GeneralMethods();
            intigrationmon = new MOnitorUI(driver);
            iccPortal = new ICCBAMPageObj(driver);
        }

        [SetUp]
        public void SetUp()
        {
            var binary = new FirefoxBinary(ConfigUtils.Read("FirefoxPath"));
            // string path = ReadFirefoxProfile();
            FirefoxProfile ffprofile = new FirefoxProfile();
            driver = new FirefoxDriver(binary, ffprofile);
            //  driver = new InternetExplorerDriver(internetExplorerDriverServerDirectory: "\\srv10177\\TEMPSHTIW$\\Desktop\\SIT\\SITSmokeTests\\IEDriverServer.exe"); 
            driver.Manage().Cookies.DeleteAllCookies();
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            // Before each test
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            datafilePath = ConfigUtils.Read("TestDataPath");
            sGMethods = new GeneralMethods();
            intigrationmon = new MOnitorUI(driver);
            iccPortal = new ICCBAMPageObj(driver);
        }

        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
            xCellFileHelper.CleanUp();
                        driver.Quit();
            Thread.Sleep(2000);
        }

       // [Test,Order(7)]
        public void TC07_Selenium_SmokeTest_ProductPlan()
        {
            try {
                //string Directive = "02 - H&M+";
                //string season = "7-2018";
                //string ProductClassification = "Blouse - Top";
                //string ProductClassificationSub = "Long sleeve";

                //string Directive = "20 - H&M Man";
                //string season = "7-2018";
                //string ProductClassification = "Coat - Top";
                //string ProductClassificationSub = "Summer";
                //string GraphicalAppearance = "Solid";
                //string articleGraphicalAppearance = "220 - LONG WOOL";
                //string TabName = "5262 - Jacket";
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                string directive = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Section");
                string season = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");
                string productClassification =
                    xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductType");
                // string productClassificationSub = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductClassificationSub");
                string graphicalAppearance =
                    xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "GraphicalAppearance");
                string articleGraphicalAppearance =
                    xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "articleGraphicalAppearance");
                string tabName = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "TabName");
               // string market = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Market");
               // string quantity = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Quantity");
               // string sellPrice = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "SellPrice");
                string articalNo = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ArticalNo");
                // 09-101
                string custCustomGr = "Women";
                string typeofConstr = "Woven";
                string Description = "AutomationTest";

                //try
                //{
                pPlanPage = new PrePlanPage(driver);
                driver.Manage().Window.Maximize();
                Thread.Sleep(1000);
                driver.Navigate().GoToUrl(ConfigUtils.Read("URL_PPlan"));
                Thread.Sleep(1000);
                pPlanPage.waitForPageToLoad();
                pPlanPage.waitForSpinnerToDisappear();
                Thread.Sleep(4000);
                if (!pPlanPage.seasonButtonStatus.Text.Contains(season))
                {
                    sGMethods.WebButton_Click(pPlanPage.get_seasonSelectionDirectiveButton(season));
                }
                Reporter.ReportEvent("Season selected", "Season selected : " + season, HP.LFT.Report.Status.Passed);
                pPlanPage.waitForPageToLoad();
                Assert.IsTrue(pPlanPage.get_ProductNameTableHeaderElement().Displayed);
                if (!pPlanPage.SectionDirectrivesDropDownElementSelectedElement.Text.Contains(directive))
                {
                    sGMethods.SelectDropDownByValue(pPlanPage.SectionDirectrivesDropDownElement,
                        directive);
                }
                Reporter.ReportEvent("Directive selected from dropdown", "Directive selected from dropdown : " + directive,
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
                //driver.FindElement(By.XPath(".//*[@id='center']/div/div[4]/div[1]/div/div[1]/div[5]")).Click();
                //try
                //{
                //    driver.FindElement(By.XPath(".//*[@id='center']/div/div[4]/div[3]/div/div/div[1]")).Click();


                //    List<IWebElement> scrolls = driver.FindElements(By.XPath(".//*[@id='center']/div/div[4]/div[3]/div/div/div")).ToList();
                //    Console.WriteLine(scrolls.Count);
                //    int j = 1;
                //    IJavaScriptExecutor js = driver as IJavaScriptExecutor;
                //    while (j < scrolls.Count)
                //    {
                //        if (j == (scrolls.Count) - 1)
                //        {
                //            Thread.Sleep(250);
                //        }
                //        js.ExecuteScript("arguments[0].scrollIntoView();", scrolls[j]);

                //        scrolls = driver.FindElements(By.XPath(".//*[@id='center']/div/div[4]/div[3]/div/div/div")).ToList();
                //        j = j + 1;
                //    }

                //}
                //catch
                //{


                //}
                //if cont more than one scroll
                Thread.Sleep(5000);
                pPlanPage.get_CreateProductTextBox();
                string productName = "AutoTest" + (DateTime.Now.ToString("yyyyhhmmssffff"));
                pPlanPage.CreateProductNameTextBoxElement.SendKeys(productName);
                pPlanPage.CreateProductNameTextBoxElement.SendKeys(Keys.Enter);
                Reporter.ReportEvent("Product name enter in text box", "Product name enter in text box : " + productName,
                    HP.LFT.Report.Status.Passed);
                Thread.Sleep(5000);
                //  Assert.IsTrue(pPlanPage.get_ProductNameFromLeftColumn(productName).Displayed);
                Reporter.ReportEvent("Product Name display and created ", "Product Name display and created",
                    HP.LFT.Report.Status.Passed);
                Thread.Sleep(2000);
                xCellFileHelper.SetCellValueByRowAndColumn("Selenium_SmokeTest1", "PPlanProduct", productName);


                pPlanPage.selectCore(productName, articalNo);
                pPlanPage.get_ProductNameFromLeftColumn(productName).Click();
                Reporter.ReportEvent("Clicked on product name", "clicked on product Name", HP.LFT.Report.Status.Passed);
              //  Assert.IsTrue(pPlanPage.get_productView_ProductName().Text.Equals(productName));

                // sGMethods.WebButton_Click(pPlanPage.get_productView_SellPrice());
                pPlanPage.get_productView_SellPrice().Click();
                Thread.Sleep(3000);
                pPlanPage.get_value_SellPrice("2").Click();
                Thread.Sleep(2000);
                sGMethods.WebButton_Click(pPlanPage.SaveButtonHeaderElement);
                pPlanPage.waitForSpinnerToDisappear();
                pPlanPage.waitForPageToLoad();
                Thread.Sleep(10000);
                Assert.IsTrue(pPlanPage.get_ProductNameFromLeftColumn(productName).Displayed);
                try
                {
                    Thread.Sleep(60000);
                    pPlanPage.TableRandomOtherColumnElement.Click();
                    pPlanPage.TableRandomOtherColumnElement.SendKeys(Keys.PageDown);
                    Thread.Sleep(5000);
                }
                catch
                {
                    Thread.Sleep(5000);
                }
                pPlanPage.get_RefreshButtonHeaderElement().Click();
               
                pPlanPage.waitForSpinnerToDisappear();
                Assert.IsTrue(pPlanPage.get_ProductNameFromLeftColumn(productName).Displayed);
                System.Diagnostics.Debug.WriteLine(pPlanPage.get_productView_ProductNumber().Text);
                productNumber = pPlanPage.get_productView_ProductNumber().Text;
                string[] ele = productNumber.Split(':');
                productNumber = ele[1].Trim();
                Reporter.ReportEvent("Product Number displayed", "Product Number displayed : " + productNumber,
                    HP.LFT.Report.Status.Passed);
                xCellFileHelper.SetCellValueByRowAndColumn("Selenium_SmokeTest1", "PPlanNumber", productNumber);
                Thread.Sleep(1000);
                pPlanPage.get_WarningLink("Product Classification missing").Click();
                Thread.Sleep(2000);
                pPlanPage.productView_ProductClassification.Click();
                Reporter.ReportEvent("Stated Product Classification", "Stated Product Classification",
                    HP.LFT.Report.Status.Passed);
                Thread.Sleep(3000);
                pPlanPage.get_value_ProductClassification(productClassification).Click();
                //try
                //{
                //    pPlanPage.productView_ProductClassificationcategories.Click();
                //    pPlanPage.get_value_ProductClassificationCategories(productClassificationSub).Click();
                //}
                //catch
                //{
                //}
                Thread.Sleep(1000);
                pPlanPage.productDescription.Click();
                Thread.Sleep(1000);
                pPlanPage.productDescription.SendKeys(Description);
               
                sGMethods.SelectDropDownByValue(pPlanPage.changeTypeOfConstructionIdDropDwon, typeofConstr);
                sGMethods.SelectDropDownByValue(pPlanPage.changeCustomsCustomerGroupIdDropDown, custCustomGr);
                Thread.Sleep(1000);
              //  pPlanPage.get_WarningLink("Please add Articles").Click();
              //Su  Thread.Sleep(1000);
               // pPlanPage.enterArtical(articalNo);
                Thread.Sleep(1000);
                pPlanPage.get_WarningLink("Please select ISWs").Click();
                Thread.Sleep(4000);
                Reporter.ReportEvent("Stated select ISWs", "Stated select ISWs", HP.LFT.Report.Status.Passed);
                pPlanPage.get_PleaseselectISWs(articalNo.Trim())[7].Click();
                Thread.Sleep(2000);
                pPlanPage.get_WarningLink("Article Graphical Appearance missing").Click();
                Thread.Sleep(1000);
                pPlanPage.selectarticleGraphicalAppearance(graphicalAppearance);
                Thread.Sleep(1000);
                Reporter.ReportEvent("Stated Article Graphical Appearance",
                    "Stated Article Graphical Appearance : " + graphicalAppearance, HP.LFT.Report.Status.Passed);
                Thread.Sleep(1000);
                pPlanPage.get_WarningLink("Article Type missing").Click();
                Thread.Sleep(1000);
                pPlanPage.typeSelectarticleGraphicalAppearance(articleGraphicalAppearance);
                Thread.Sleep(1000);
                Reporter.ReportEvent("StatedArticle Type", "Stated Article Type : " + articleGraphicalAppearance,
                    HP.LFT.Report.Status.Passed);
                Thread.Sleep(1000);
                pPlanPage.clickOnVersion();
                //market = market.Trim();
                //quantity = quantity.Trim();
                //string[] splitMarket = market.Split('|');
                //string[] splitQuantity = quantity.Split('|');
                Reporter.ReportEvent("Enter Market and Quantity", "Enter Market and Quantity", HP.LFT.Report.Status.Passed);
                //for (int i = 0; i < pPlanPage.getAllMarket().Count; i++)
                //{
                //    pPlanPage.setmarketquantity(pPlanPage.getAllMarket()[i], quantity, "1");
                //}
                Thread.Sleep(1000);

                pPlanPage.setMarketQuant("600000");

                //pPlanPage.setmarketquantity("NO-02", "999", "1");
                //pPlanPage.setmarketquantity("DK-03", "999", "1");
                sGMethods.WebButton_Click(pPlanPage.SaveButtonHeaderElement);
                Thread.Sleep(1000);
                pPlanPage.waitForSpinnerToDisappear();
                pPlanPage.waitForPageToLoad();
                Thread.Sleep(1000);
                Assert.IsTrue(pPlanPage.get_ProductNameFromLeftColumn(productName).Displayed);
                if (!pPlanPage.SaveButtonHeaderElement.GetAttribute("disabled").Equals("true"))
                {
                    sGMethods.WebButton_Click(pPlanPage.SaveButtonHeaderElement);
                    pPlanPage.waitForSpinnerToDisappear();
                    pPlanPage.waitForPageToLoad();
                }
                Thread.Sleep(5000);
                pPlanPage.verifyArticleNumberDisplays(productName);
                if (!pPlanPage.SaveButtonHeaderElement.GetAttribute("disabled").Equals("true"))
                {
                    sGMethods.WebButton_Click(pPlanPage.SaveButtonHeaderElement);
                    pPlanPage.waitForSpinnerToDisappear();
                    pPlanPage.waitForPageToLoad();
                }
                Reporter.ReportEvent("Article Number displayed", "Article Number displayed", HP.LFT.Report.Status.Passed);
                //*************************** Monitor  ****************************
                try
                {
                    intigrationmon.LaunchWindowIntegrationmonitor(productNumber, ConfigUtils.Read("URL_IMonitor"));
                    intigrationmon.SearchIntegrationmonitor(productNumber, "PlanInformationOnArticleLevel");
                    Reporter.ReportEvent("Search Integration monitor 'PlanInformationOnArticleLevel' Processed",
                        "Search Integration monitor 'PlanInformationOnArticleLevel' Processed",
                        HP.LFT.Report.Status.Passed);
                    intigrationmon.SearchIntegrationmonitor(productNumber, "PlanExternalArticleInfo");
                    Reporter.ReportEvent("Search Integration monitor 'PlanExternalArticleInfo' Processed",
                        "Search Integration monitor 'PlanExternalArticleInfo' Processed", HP.LFT.Report.Status.Passed);
                    intigrationmon.SearchIntegrationmonitor(productNumber, "ProductSellPrice");
                    Reporter.ReportEvent("Search Integration monitor 'ProductSellPrice' Processed",
                        "Search Integration monitor 'ProductSellPrice' Processed", HP.LFT.Report.Status.Passed);
                    intigrationmon.SearchIntegrationmonitor(productNumber, "PlanExternalArticleInfo");
                    Reporter.ReportEvent("Search Integration monitor 'PlanExternalArticleInfo' Processed",
                        "Search Integration monitor 'PlanExternalArticleInfo' Processed", HP.LFT.Report.Status.Passed);
                    intigrationmon.SearchIntegrationmonitor(productNumber, "PlanInformationOnMarketLevel");
                    Reporter.ReportEvent("Search Integration monitor 'PlanInformationOnMarketLevel' Processed",
                        "Search Integration monitor 'PlanInformationOnMarketLevel' Processed", HP.LFT.Report.Status.Passed);
                    intigrationmon.SearchIntegrationmonitor(productNumber, "PlanExternalMarketInfo");
                    Reporter.ReportEvent("Search Integration monitor 'PlanExternalMarketInfo' Processed",
                        "Search Integration monitor 'PlanExternalMarketInfo' Processed", HP.LFT.Report.Status.Passed);
                    intigrationmon.SearchIntegrationmonitor(productNumber, "WeekPlanOnMarketLevel");
                    Reporter.ReportEvent("Search Integration monitor 'WeekPlanOnMarketLevel' Processed",
                        "Search Integration monitor 'WeekPlanOnMarketLevel' Processed", HP.LFT.Report.Status.Passed);
                    intigrationmon.CloseIntegrationmonitorAndBackToPp();
                }
                catch (Exception ex)
                {
                    Reporter.ReportEvent("Fail to verify intigration monitor Section",
                        "Fail to verify intigration monitor Section Exception  : " + ex.StackTrace,
                        HP.LFT.Report.Status.Passed);
                }

                //*************************** ICC  ****************************
                try
                {
                    iccPortal.LaunchICCWindow();
                    
                    /// iccPortal.LaunchWindowICC("Acc2010", ConfigUtils.Read("URL_ICC"));
                    /// 
                    try
                    {
                        iccPortal.SelectValueFromApplicationDropDown(productNumber, "PlanInformationOnArticleLevel_PLES");
                        iccPortal.VerifySearchResult("PlanInformationOnArticleLevel_PLES");
                        iccPortal.VerifyPortInSearchResult("CDW");
                        iccPortal.VerifyPortInSearchResult("Fantomen");
                        Reporter.ReportEvent("ICC portal 'PlanInformationOnArticleLevel_PLES' Processed for CDW and Fantomen",
                            "ICC portal  PlanInformationOnArticleLevel_PLES Processed for CDW and Fantomen",
                            HP.LFT.Report.Status.Passed);
                    }
                    catch
                    {
                        Reporter.ReportEvent("ICC portal 'PlanInformationOnArticleLevel_PLES' Processed for CDW and Fantomen",
                           "ICC portal  PlanInformationOnArticleLevel_PLES Processed for CDW and Fantomen",
                           HP.LFT.Report.Status.Failed);
                    }
                    try
                    {
                        iccPortal.SelectValueFromApplicationDropDown(productNumber, "PlanInformationOnMarketLevel_PLES");
                        iccPortal.VerifySearchResult("PlanInformationOnMarketLevel_PLES");
                        iccPortal.VerifyPortInSearchResult("CDW");
                        iccPortal.VerifyPortInSearchResult("Fantomen");
                        Reporter.ReportEvent("ICC portal 'PlanInformationOnMarketLevel_PLES' Processed for CDW and Fantomen",
                            "ICC portal  PlanInformationOnMarketLevel_PLES Processed for CDW and Fantomen",
                            HP.LFT.Report.Status.Passed);
                    }
                    catch
                    {
                        Reporter.ReportEvent("ICC portal 'PlanInformationOnMarketLevel_PLES' Processed for CDW and Fantomen",
                           "ICC portal  PlanInformationOnMarketLevel_PLES Processed for CDW and Fantomen",
                           HP.LFT.Report.Status.Failed);
                    }
                    try
                    {
                        iccPortal.SelectValueFromApplicationDropDown(productNumber, "PlanInformationOnProductLevel_PLES");
                        iccPortal.VerifySearchResult("PlanInformationOnProductLevel_PLES");
                        iccPortal.VerifyPortInSearchResult("CDW");
                        iccPortal.VerifyPortInSearchResult("Fantomen");
                        Reporter.ReportEvent("ICC portal 'PlanInformationOnProductLevel_PLES' Processed for CDW and Fantomen",
                            "ICC portal  PlanInformationOnProductLevel_PLES Processed for CDW and Fantomen",
                            HP.LFT.Report.Status.Passed);
                    }
                    catch
                    {
                        Reporter.ReportEvent("ICC portal 'PlanInformationOnProductLevel_PLES' Processed for CDW and Fantomen",
                          "ICC portal  PlanInformationOnProductLevel_PLES Processed for CDW and Fantomen",
                          HP.LFT.Report.Status.Failed);
                    }
                    try
                    {
                        iccPortal.SelectValueFromApplicationDropDown(productNumber, "PlanExternalArticleInfo_PLES");
                        iccPortal.VerifySearchResult("PlanExternalArticleInfo_PLES");
                        iccPortal.VerifyPortInSearchResult("MWS");
                        iccPortal.VerifyPortInSearchResult("PRA");
                        Reporter.ReportEvent("ICC portal 'PlanExternalArticleInfo_PLES' Processed for MWS and PRA",
                            "ICC portal 'PlanExternalArticleInfo_PLES' Processed for MWS and PRA", HP.LFT.Report.Status.Passed);
                    }
                    catch
                    {
                        Reporter.ReportEvent("ICC portal 'PlanExternalArticleInfo_PLES' Processed for MWS and PRA",
                           "ICC portal 'PlanExternalArticleInfo_PLES' Processed for MWS and PRA", HP.LFT.Report.Status.Failed);
                    }
                    try
                    {
                        iccPortal.SelectValueFromApplicationDropDown(productNumber, "PlanExternalMarketInfo_PLES");
                        iccPortal.VerifySearchResult("PlanExternalMarketInfo_PLES");
                        iccPortal.VerifyPortInSearchResult("MWS");
                        iccPortal.VerifyPortInSearchResult("PRA");
                        Reporter.ReportEvent("ICC portal 'PlanExternalMarketInfo_PLES' Processed for MWS and PRA",
                            "ICC portal 'PlanExternalMarketInfo_PLES' Processed for MWS and PRA", HP.LFT.Report.Status.Passed);
                    }
                    catch
                    {
                        Reporter.ReportEvent("ICC portal 'PlanExternalMarketInfo_PLES' Processed for MWS and PRA",
                           "ICC portal 'PlanExternalMarketInfo_PLES' Processed for MWS and PRA", HP.LFT.Report.Status.Failed);
                    }
                    try
                    {
                        iccPortal.SelectValueFromApplicationDropDown(productNumber, "PlanExternalProductInfo_PLES");
                        iccPortal.VerifySearchResult("PlanExternalProductInfo_PLES");
                        iccPortal.VerifyPortInSearchResult("MWS");
                        iccPortal.VerifyPortInSearchResult("PRA");
                        Reporter.ReportEvent("ICC portal 'PlanExternalProductInfo_PLES' Processed for MWS and PRA",
                            "ICC portal 'PlanExternalProductInfo_PLES' Processed for MWS and PRA", HP.LFT.Report.Status.Passed);
                    }
                    catch
                    {
                        Reporter.ReportEvent("ICC portal 'PlanExternalProductInfo_PLES' Processed for MWS and PRA",
                            "ICC portal 'PlanExternalProductInfo_PLES' Processed for MWS and PRA", HP.LFT.Report.Status.Failed);
                    }
                   
                    intigrationmon.CloseIntegrationmonitorAndBackToPp();
                }
                catch (Exception ex)
                {
                    intigrationmon.CloseIntegrationmonitorAndBackToPp();
                    Reporter.ReportEvent("Fail to verify ICC portal Section",
                        "Fail to verify ICC portal Section Exception  : " + ex.StackTrace, HP.LFT.Report.Status.Failed);
                }

                //*****************************************************************
                pPlanPage.get_SelectPlanBtn().Click();
                Reporter.ReportEvent("Clicked on Selecte plan button", "Clicked on Selecte plan button",
                    HP.LFT.Report.Status.Passed);
                pPlanPage.get_ISWCheckbox().Click();
                pPlanPage.get_ByAllPmButton().Click();
                Reporter.ReportEvent("Clicked on All PM button", "Clicked on All Plan button", HP.LFT.Report.Status.Passed);
                pPlanPage.waitForPageToLoad();
                pPlanPage.waitForSpinnerToDisappear();
                pPlanPage.get_ReadyToOrderButton().Click();
                pPlanPage.waitForPageToLoad();
                pPlanPage.waitForSpinnerToDisappear();
                for (int i = 0; i < 30; i++)
                {
                    if (pPlanPage.get_SelectPlanBtn().GetAttribute("class").Contains("disabled-popup"))
                    {
                        break;
                    }
                }
                Assert.IsTrue((pPlanPage.get_productView_SellPrice().Displayed));
                System.Diagnostics.Debug.WriteLine(pPlanPage.get_orderNumber().GetAttribute("data-content"));
                orderNumber = pPlanPage.get_OderNumberFromUI();
                Reporter.ReportEvent("Order number created", "Order Number created :" + orderNumber,
                    HP.LFT.Report.Status.Passed);
                xCellFileHelper.SetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber", orderNumber);
                try
                {
                    intigrationmon.LaunchWindowIntegrationmonitor(orderNumber, ConfigUtils.Read("URL_IMonitor"));
                    intigrationmon.SearchIntegrationmonitor(orderNumber, "BoardcardPlan");
                    Reporter.ReportEvent("Search Integration monitor 'BoardcardPlan' Processed",
                        "Search Integration monitor 'BoardcardPlan' Processed", HP.LFT.Report.Status.Passed);
                    intigrationmon.CloseIntegrationmonitorAndBackToPp();
                }
                catch (Exception ex)
                {
                    Reporter.ReportEvent("Fail to verify intigration monitor Section",
                        "Fail to verify intigration monitor Section Exception  : " + ex.StackTrace,
                        HP.LFT.Report.Status.Failed);
                }
                try
                {

                    iccPortal.LaunchICCWindow();

                    iccPortal.SelectValueFromApplicationDropDownForBoardCard(orderNumber, "BoardcardPlan_PLES");
                    try
                    {
                        iccPortal.VerifySearchResult("BoardcardPlan_PLES");
                        iccPortal.VerifyPortInSearchResult("HMOrder");
                        Reporter.ReportEvent("ICC portal 'BoardcardPlan_PLES' Processed for HMOrder",
                     "ICC portal 'BoardcardPlan_PLES' Processed for HMOrder", HP.LFT.Report.Status.Passed);
                    }
                    catch
                    {
                        Reporter.ReportEvent("ICC portal 'BoardcardPlan_PLES' Processed for HMOrder",
                       "ICC portal 'BoardcardPlan_PLES' Processed for HMOrder", HP.LFT.Report.Status.Failed);
                    }
                   
                   
                 
                    intigrationmon.CloseIntegrationmonitorAndBackToPp();
                }
                catch (Exception ex)
                {
                    intigrationmon.CloseIntegrationmonitorAndBackToPp();
                    Reporter.ReportEvent("Fail to verify ICC portal Section",
                        "Fail to verify ICC portal Section Exception  : " + ex.StackTrace, HP.LFT.Report.Status.Failed);
                }
            }catch(Exception exp)
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                Screenshot sc = ((ITakesScreenshot)driver).GetScreenshot();
                sc.SaveAsFile(ESSpath, ImageFormat.Png);
                Reporter.ReportEvent("Product Plan Script Execution failed",
                       "Product Plan Script Execution failed Exception  : " + exp.StackTrace, HP.LFT.Report.Status.Failed, ESSpath);
                System.Diagnostics.Debug.WriteLine("***************** message :" + exp.Message);
                System.Diagnostics.Debug.WriteLine("***************** stack :" + exp.StackTrace);
                Assert.IsTrue(2 == 3);
            }
        }
    }
}
