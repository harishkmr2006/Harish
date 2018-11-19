using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using NUnit.Framework;

namespace SITSmokeTests
{

    class PrePlanPage
    {
        IWebDriver driver;
        private WebDriverWait wait;

        private WebDriverWait extWait;

        GeneralMethods sGMethods;

        public PrePlanPage(IWebDriver drivername)
        {
            PageFactory.InitElements(drivername, this);
            driver = drivername;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            extWait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            sGMethods = new GeneralMethods();
        }

        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'navbar-static-top')]//select")]
        public IWebElement SectionDirectrivesDropDownElement { get; set; }

        [FindsBy(How = How.XPath,
            Using = "//*[contains(@class,'navbar-static-top')]//select//option[@selected='selected']")]
        public IWebElement SectionDirectrivesDropDownElementSelectedElement { get; set; }

        [FindsBy(How = How.XPath, Using = "//pp-season-selection-directive//button[contains(@class,'active')]")]
        public IWebElement seasonButtonStatus { get; set; }


        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'navbar-static-top')]//button[contains(@title,'Save')]")]
        public IWebElement SaveButtonHeaderElement { get; set; }


        [FindsBy(How = How.XPath,
            Using =
                "//*[@id='articleListGrid']//*[@class='grid-canvas']/div[1]//*[contains(@class,'graphical-appearance')]")]
        public IWebElement articleGraphicalAppearance { get; set; }

        [FindsBy(How = How.XPath,
            Using = "//*[@id='articleListGrid']//*[@class='grid-canvas']/div[1]//*[contains(@class,'type')]/input")]
        public IWebElement typerticleGraphicalAppearance { get; set; }

        [FindsBy(How = How.XPath, Using = "//button[contains(text(),'Close Product')]")]
        public IWebElement CloseProdcutButtonElement { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'department-tabs')]//li[@class='active']/a")]
        public IWebElement DepartmentTabActivElement { get; set; }


        [FindsBy(How = How.XPath, Using = ".//*[@id='center']//*[@class='ag-pinned-left-cols-container']//input")]
        public IWebElement CreateProductNameTextBoxElement { get; set; }


        [FindsBy(How = How.XPath, Using = "(.//*[@id='center']//*[@class='ag-body-container']//div[@colid='w4'])[6]")]
        public IWebElement TableRandomOtherColumnElement { get; set; }

        [FindsBy(How = How.XPath,
            Using =
                "//*[@id='product-view']//pp-expand-region[@paneltitle='Product Details']//*[@class='product-classification']//input")]
        public IWebElement productView_ProductClassification { get; set; }

        [FindsBy(How = How.XPath,
            Using =
                "//*[@id='product-view']//pp-expand-region[@paneltitle='Product Details']//input[contains(@class,'product-description')]")]
        public IWebElement productDescription { get; set; }

        [FindsBy(How = How.XPath,
            Using =
                "//*[@id='product-view']//pp-expand-region[@paneltitle='Product Details']//select[contains(@ng-change,'changeCustomsCustomerGroupId')]")]
        public IWebElement changeCustomsCustomerGroupIdDropDown { get; set; }

        [FindsBy(How = How.XPath,
            Using =
                "//*[@id='product-view']//pp-expand-region[@paneltitle='Product Details']//select[contains(@ng-change,'changeTypeOfConstructionId')]")]
        public IWebElement changeTypeOfConstructionIdDropDwon { get; set; }


        [FindsBy(How = How.XPath, Using = "//*[@id='articleListGrid']//*[contains(@class,'article-name-container')]")]
        public IWebElement productView_ArticleSpan { get; set; }

        [FindsBy(How = How.XPath,
            Using =
                "//*[@id='product-view']//pp-expand-region[@paneltitle='Product Details']//*[@class='classification-categories']//input")]
        public IWebElement productView_ProductClassificationcategories { get; set; }


        public IWebElement get_departmentTabElement(string TabName)
        {
            return extWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(
                "//*[contains(@class,'department-tabs')]//li/a[contains(text(),'" + TabName + "')]")));
        }

        public IWebElement get_RefreshButtonHeaderElement()
        {
            return extWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(
                "//*[contains(@class,'navbar-static-top')]//button[contains(@title,'Refresh')]")));
        }

        public IWebElement get_productView_ProductNumber()
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                " //*[@id='product-view']//pp-product-summary//*[@class='product-number']")));
        }

        public IWebElement get_productView_ProductName()
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                "//*[@id='product-view']//pp-product-summary//*[@auto-title='product.name']")));
        }

        public IWebElement get_productView_SellPrice()
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                "//*[@id='product-view']//pp-product-summary//*[@class='sell-price']//input")));
        }

        public IWebElement get_value_SellPrice(string sellPriceIndex)
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                "//ul//li[contains(text(),'None')]/following-sibling::li[" + sellPriceIndex + "]")));
        }


        public IWebElement get_value_ProductClassification(string productClassification)
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                "//ul//*[text()='" + productClassification + "']")));
        }

        public IWebElement get_TableRandomColumnElement(string index)
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                "(.//*[@id='center']//*[@class='ag-body-container']//div[@colid='w" + index + "'])[6]")));
        }


        public IWebElement get_value_ProductClassificationCategories(string productClassification)
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                "//ul//*[text()='" + productClassification + "']")));
        }


        public void enterArtical(string articalNo)
        {
            articalNo = articalNo.Trim();
            driver.FindElement(By.XPath(".//*[@id='articleListGrid']//input")).Click();
            driver.FindElement(By.XPath(".//*[@id='articleListGrid']//input")).Clear();
            driver.FindElement(By.XPath(".//*[@id='articleListGrid']//input")).SendKeys(articalNo);
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//ul//*[contains(text(),'" + articalNo + "')]")).Click();
        }

        public void selectCore(string ProductName,string articalnum)
        {
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='center']//div[contains(.,\"" + ProductName + "\")]/div/span[@class=\"ui-icon ui-icon-triangle-down-hollow\"]")));
            string rownum = driver.FindElement(By.XPath(".//*[@id='center']//span[text()=\"" + ProductName + "\"]/../../../..")).GetAttribute("row");
            int irownum = (Int32.Parse(rownum))+1;
            driver.FindElement(By.XPath(".//*[@id='center']//div[contains(.,\"" + ProductName + "\")]/div/span[@class=\"ui-icon ui-icon-triangle-down-hollow\"]")).Click();
            Thread.Sleep(2000);
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(".//*[@id='center']//div[@row=\"" + irownum.ToString() + "\"]//span[@class=\"cell-fill article-name-container is-new\"]")));
           
            driver.FindElement(By.XPath(".//*[@id='center']//div[@row=\""+ irownum.ToString() + "\"]//span[@class=\"cell-fill article-name-container is-new\"]")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.XPath(".//*[@id='center']//input[@class=\"editor-text ui-autocomplete-input\"]")).SendKeys(articalnum);
            Thread.Sleep(1000);
            driver.FindElement(By.XPath(".//*[@id='gridTabs']//a[text()=\"Details\"]")).Click();
            Thread.Sleep(5000);
            try
            {
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//div[@row=\"" + irownum.ToString() + "\"]/div[@colid=\"storeCategoryAssortment\"]")));
                driver.FindElement(By.XPath(".//div[@row=\"" + irownum.ToString() + "\"]/div[@colid=\"storeCategoryAssortment\"]")).Click();
                driver.FindElement(By.XPath(".//li[@title=\"Core\"]")).Click();
            }
            catch
            {

            }
           
        }
        public void selectarticleGraphicalAppearance(string articalGraphic)
        {
            driver.FindElement(By.XPath(
                    ".//*[@id='articleListGrid']//*[@class='grid-canvas']/div[1]//*[contains(@class,'graphical-appearance')]/*"))
                .Click();
            Thread.Sleep(5000);
            extWait.Until(ExpectedConditions.ElementIsVisible((By.XPath("//ul//a[text()='" + articalGraphic + "']"))))
                .Click();
        }


        public void typeSelectarticleGraphicalAppearance(string articleGraphicalAppearance)
        {
            driver.FindElement(By.XPath(
                    ".//*[@id='articleListGrid']//*[@class='grid-canvas']/div[1]//*[contains(@class,'type')]/*"))
                .Click();
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath("//ul//a[text()='" + articleGraphicalAppearance + "']"))))
                .Click();
            driver.FindElement(By.XPath(
                    ".//*[@id='articleListGrid']//*[@class='grid-canvas']/div[1]//*[contains(@class,'flow-matrix')]/*"))
                .Click();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath(
                    ".//*[@id='articleListGrid']//*[@class='grid-canvas']/div[1]//*[contains(@class,'flow-matrix')]/*"))
                .Click();
            extWait.Until(ExpectedConditions.ElementIsVisible((By.XPath("//ul//a[text()='$']")))).Click();
        }

        public void setmarketquantity1(int index, string quantity, string versionNO)
        {
            driver.FindElement(By.XPath("(//*[@id='center']//div[contains(@class,'market-quantity')]/div/span)[" +
                                        index + "]"))
                .Click();
            Thread.Sleep(2000);
            Actions action = new Actions(driver);
            action
                .MoveToElement(driver.FindElement(By.XPath(
                    "(//*[@id='center']//div[contains(@class,'market-quantity')]/div/span)[" + index + "]")))
                .DoubleClick()
                .Build()
                .Perform();
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath("//*[@id='center']//div[contains(@class,'market-quantity')]//input"))))
                .Click();
            Thread.Sleep(1000);
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath("//*[@id='center']//div[contains(@class,'market-quantity')]//input"))))
                .SendKeys(quantity);
            Thread.Sleep(1000);
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath("//*[@id='center']//div[contains(@class,'market-quantity')]//input"))))
                .SendKeys(Keys.Enter);
            Thread.Sleep(10000);
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("(//*[@id='center']//div[contains(@class,'version-column')]/span)[" + index + "]")))
                .Click();
            Thread.Sleep(1000);
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[@id='center']//div[contains(@class,'version-column')]/input")))
                .Click();
            Thread.Sleep(2000);
            extWait.Until(ExpectedConditions.ElementIsVisible((By.XPath("//ul//a[text()='" + versionNO + "']"))))
                .Click();
            Thread.Sleep(1000);
            extWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(" //button[text()='Apply']"))).Click();
            Thread.Sleep(3000);
            //extWait
            //    .Until(ExpectedConditions.ElementIsVisible(
            //        (By.XPath(
            //            "//*[contains(@class,'ag-total-valid')]//*[(contains(text(),'0') or contains(text(),'9')) and contains(@class,'market-level')]")
            //        )))
            //    .Click();
            extWait
              .Until(ExpectedConditions.ElementToBeClickable(
                  (By.XPath(
                      ".//*[@id='product-view']//button[contains(.,\"Save Product\")]")
                  )));
        }

        public void setmarketquantity(string market, string quantity, string versionNO)
        {
            setmarketquantity1(getIndexFromMarketQuant(market), quantity, versionNO);
        }

        public void clickOnVersion()
        {
            extWait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//pp-version-button/button"))).Click();
            Thread.Sleep(3000);
        }

        public List<string> getAllMarket()
        {
            List<IWebElement> getAllMarket = extWait
                .Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(
                    By.XPath("//*[@id='center']//div[contains(@class,'readonly market-column')]/span")))
                .ToList();
            List<string> marketList = new List<string>();

            foreach(IWebElement ele in getAllMarket)
            {
                marketList.Add(ele.Text);
            }
            return marketList;
        }

        public int getIndexFromMarketQuant(string market)
        {
            int i = 1;
            List<IWebElement> getAllMarket = extWait
                .Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(
                    By.XPath("//*[@id='center']//div[contains(@class,'readonly market-column')]/span")))
                .ToList();
            foreach (IWebElement ele in getAllMarket)
            {
                if (ele.Text.Contains(market))
                {
                    return i;
                }
                i++;
            }
            return 0;
        }

        public void setmarketquantity()
        {
            driver.FindElement(By.XPath("(//*[@id='center']//div[contains(@class,'market-quantity')]/div/span)[4]"))
                .Click();
            Thread.Sleep(2000);
            Actions action = new Actions(driver);
            action
                .MoveToElement(driver.FindElement(By.XPath(
                    "(//*[@id='center']//div[contains(@class,'market-quantity')]/div/span)[4]")))
                .DoubleClick()
                .Build()
                .Perform();
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath("//*[@id='center']//div[contains(@class,'market-quantity')]//input"))))
                .Click();
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath("//*[@id='center']//div[contains(@class,'market-quantity')]//input"))))
                .SendKeys("100");
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath("//*[@id='center']//div[contains(@class,'market-quantity')]//input"))))
                .SendKeys(Keys.Enter);
            Thread.Sleep(3000);
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath(
                        "//*[contains(@class,'ag-total-valid')]//*[contains(text(),'100') and contains(@class,'market-level')]")
                    )))
                .Click();
        }


        public void verifyArticleNumberDisplays(string productName)
        {
            WebDriverWait waiter = new WebDriverWait(driver, TimeSpan.FromSeconds(4));
            IWebElement ele = null;
            int flag = 0;
            int count = 0;
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
            do
            {
                try
                {
                    ele = waiter.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                        ".//*[@id='articleListGrid']//*[@class='grid-canvas']/div[1]//*[contains(@class,'sales-product-number')]/span")));
                    flag = 1;
                    ele.Click();
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("disabled+++++++++++++++++++" +
                                                       SaveButtonHeaderElement.GetAttribute("disabled"));
                    if (!SaveButtonHeaderElement.GetAttribute("disabled").Contains("true"))
                    {
                        System.Diagnostics.Debug.WriteLine("disabled+++++++++++++++++++" +
                                                           SaveButtonHeaderElement.GetAttribute("disabled"));
                        SaveButtonHeaderElement.Click();
                        waitForPageToLoad();
                        waitForSpinnerToDisappear();
                        Thread.Sleep(3000);
                    }
                    get_RefreshButtonHeaderElement().Click();
                   
                    waitForSpinnerToDisappear();
                    Thread.Sleep(3000);
                    get_ProductNameFromLeftColumn(productName);
                    Thread.Sleep(1000);
                }
            } while ((flag == 0) || ((++count) == 20));
            if (flag == 1)
            {
                System.Diagnostics.Debug.WriteLine("Element has been found.!!" + ele.Text);
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Element has not been found.!!");
            }
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(30));
        }


        public List<IWebElement> get_PleaseselectISWs(string articalNo)
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath(
                            "//*[@id='articleListGrid']//*[@class='grid-canvas']//span[contains(.,'" + articalNo +
                            "')]/parent::div/following-sibling::div[contains(@class,'week week')]")))
                .ToList();
        }

        public List<IWebElement> get_ProductNameList()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath(
                            "//*[@id='center']//*[@class='ag-pinned-left-cols-container']//span[@class='product - view - link']")))
                .ToList();
        }


        public IWebElement get_ProductNameFromLeftColumn(string productName)
        {
            return extWait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                ".//*[@id='center']//*[@class='ag-pinned-left-cols-container']//*[text()='" + productName + "']")));
        }

        public IWebElement get_seasonSelectionDirectiveButton(string season)
        {
            return driver.FindElement(By.XPath(" //pp-season-selection-directive//button[@data-season-name='" + season +
                                               "']"));
        }

        public IWebElement get_ProductNameTableHeaderElement()
        {
            return wait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//*[@class='ag-header']//*[text()='Product Name']")));
        }

        public IWebElement get_WarningLink(string warningText)
        {
            return wait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//a[text()='" + warningText + "']")));
        }


        public void waitForSpinnerToDisappear()
        {
            Thread.Sleep(2000);
            try
            {
                extWait.Until(ExpectedConditions.ElementExists(
                    By.XPath("//*[@id='loadingSpinner' and @style='display: block;']")));
            }
            catch
            {
            }
            try
            {
                extWait.Until(
                    ExpectedConditions.ElementExists(
                        By.XPath("//*[@id='loadingSpinner' and @style='display: none;']")));
            }
            catch
            {
            }
        }

        public void get_CreateProductTextBox()
        {
            WebDriverWait waiter = new WebDriverWait(driver, TimeSpan.FromSeconds(1));
            IWebElement ele = null;
            int flag = 0;
            int count = 0;
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
            do
            {
                try
                {
                    ele = waiter.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                        ".//*[@id='center']//*[@class='ag-pinned-left-cols-container']//*[text()='Create product']")));
                    flag = 1;
                    ele.Click();
                }
                catch
                {
                    driver
                        .FindElement(By.XPath(
                            "(.//*[@id='center']//*[@class='ag-body-container']//div[@colid='w25'])[20]"))
                        .SendKeys(Keys.PageDown);
                    Thread.Sleep(1000);
                }
            } while ((flag == 0) || ((++count) == 20));
            if (flag == 1)
            {
                System.Diagnostics.Debug.WriteLine("Element has been found.!!");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Element has not been found.!!");
            }
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
        }

        public void waitForPageToLoad()
        {
            Thread.Sleep(5000);
           // wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*")));
        }

        protected IWebElement waitForElementToExistInDom(By element)
        {
            return wait.Until(ExpectedConditions.ElementExists(element));
        }

        protected IWebElement WaitForElementToDisplaElement(By element)
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(element));
        }

        public List<IWebElement> get_AlertWarning()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//*[@class='market-grid-disabled-overlay']//*[contains(@class,'alert-warning')]")))
                .ToList();
        }

        public IWebElement get_TypeSelectDropDownOptions(string option)
        {
            for (int i = 0; i < 20; i++)
            {
                List<IWebElement> ele =
                    driver.FindElements(By.XPath("//select[@id='ctl00_MainContent_cmbType']/option")).ToList();
                System.Diagnostics.Debug.WriteLine("*****************: " + ele.Count);
                if (ele.Count > 5)
                {
                    break;
                }
                Thread.Sleep(2000);
            }
            {
            }
            return extWait.Until(
                ExpectedConditions.ElementExists(By.XPath("//select[@id='ctl00_MainContent_cmbType']/option[text()='" +
                                                          option + "']")));
        }


        public List<IWebElement> get_AllStatusFromSearchResult()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//*[@id='ctl00_MainContent_GridViewExportLog']/tbody/tr/td[12]")))
                .ToList();
        }


        public IWebElement get_RecordCount()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='ctl00_MainContent_cmbRecordLimit']")));
        }


        public IWebElement get_SelectPlanBtn()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(.,'Select Plan')]")));
        }

        public IWebElement get_ISWCheckbox()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='agHeaderCellLabel']//input[@type='checkbox']")));
        }

        public IWebElement get_ByAllPmButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(.,'Buy all PM')]")));
        }

        public IWebElement get_ReadyToOrderButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//button[contains(.,'Ready to Order')]")));
        }

        public IWebElement get_orderNumber()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath(
                        "//*[@class='infotip' and contains(@data-content,'Locked for changes') and contains(@data-content,'Order No')]")));
        }

        public string Between(string STR, string FirstString, string LastString)
        {
            string FinalString;
            int Pos1 = STR.IndexOf(FirstString) + FirstString.Length;
            int Pos2 = STR.IndexOf(LastString);
            FinalString = STR.Substring(Pos1, Pos2 - Pos1);
            return FinalString;
        }

        public string get_OderNumberFromUI()
        {
            string odNumber = get_orderNumber().GetAttribute("data-content");
            return Between(odNumber, "Order No.", "</div><div").Trim();
        }


        public void setMarketQuant(string quantity)
        {
            driver.FindElement(By.XPath("(.//*[@id='center']//*[contains(.,'Total excl. %PM') and @class='ag-pinned-left-cols-viewport']/following-sibling::div[2]//div[contains(@class,'sum-article-quantity')])[1]")).Click();
            Actions action = new Actions(driver);
            action
                .MoveToElement(driver.FindElement(By.XPath(
                    "(.//*[@id='center']//*[contains(.,'Total excl. %PM') and @class='ag-pinned-left-cols-viewport']/following-sibling::div[2]//div[contains(@class,'sum-article-quantity')])[1]")))
                .DoubleClick()
                .Build()
                .Perform();


            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath("(.//*[@id='center']//*[contains(.,'Total excl. %PM') and @class='ag-pinned-left-cols-viewport']/following-sibling::div[2]//div[contains(@class,'sum-article-quantity')])[1]//input"))))
                .Click();
            Thread.Sleep(1000);
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    (By.XPath("(.//*[@id='center']//*[contains(.,'Total excl. %PM') and @class='ag-pinned-left-cols-viewport']/following-sibling::div[2]//div[contains(@class,'sum-article-quantity')])[1]//input"))))
                .SendKeys(quantity);
            Thread.Sleep(1000);


            extWait
                            .Until(ExpectedConditions.ElementIsVisible(
                                (By.XPath("(.//*[@id='center']//*[contains(.,'Total excl. %PM') and @class='ag-pinned-left-cols-viewport']/following-sibling::div[2]//div[contains(@class,'sum-article-quantity')])[1]//input"))))
                            .SendKeys(Keys.Enter);
            Thread.Sleep(15000);
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("(//*[@id='center']//div[contains(@class,'version-column')]/span)[3]")))
                .Click();
            Thread.Sleep(2000);
            extWait
                .Until(ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[@id='center']//div[contains(@class,'version-column')]/input")))
                .Click();
            Thread.Sleep(5000);
            extWait.Until(ExpectedConditions.ElementIsVisible((By.XPath("//ul//a[text()='1']"))))
                .Click();
            Thread.Sleep(1000);
            extWait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//label[contains(.,'All PMs')]//input"))).Click();
            Thread.Sleep(1000);
            extWait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Apply']"))).Click();
            Thread.Sleep(3000);
        }
    }
}
