using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using NUnit.Framework;

namespace SITSmokeTests
{
    class MOnitorUI
    {

        IWebDriver driver;
        private WebDriverWait wait;
        private WebDriverWait extWait;
        static string currentWindow;
        GeneralMethods sGMethods;

        public MOnitorUI(IWebDriver drivername)
        {
            PageFactory.InitElements(drivername, this);
            driver = drivername;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            extWait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            sGMethods = new GeneralMethods();
        }

        [FindsBy(How = How.XPath, Using = "//*[@id='ctl00_MainContent_LinkButtonKeys']")]
        public IWebElement SearchButtonHeaderElement { get; set; }

        public IWebElement get_ExportInHeaderElement()
        {
            return extWait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(
                "//*[@class='TopTableMenu']//*[text()='Exports']")));
        }
        public IWebElement get_searchKeyTextBox()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='ctl00_MainContent_textBoxKeys']")));
        }
        public List<IWebElement> get_AllProductNumberFromSearchResult()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//*[@id='ctl00_MainContent_GridViewExportLog']/tbody/tr/td[5]")))
                .ToList();
        }

        public List<IWebElement> get_ProductResult(string index)
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//*[@id='ctl00_MainContent_GridViewExportLog']/tbody/tr/td[" + index + "]")))
                .ToList();
        }


        public IWebElement get_TypeSelectDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//select[@id='ctl00_MainContent_cmbType']")));
        }


        public void LaunchWindowIntegrationmonitor(string productNumber,string url)
        {
            LaunchTabAndNavigate();
            currentWindow = get_CurrentWindowHandle();
            SwitchToNewWindow();
            driver.Navigate().GoToUrl(url);
            get_ExportInHeaderElement().Click();
            Assert.IsTrue(get_searchKeyTextBox().Displayed);
            System.Diagnostics.Debug.WriteLine(get_AllProductNumberFromSearchResult().Count);
            int resultcount = get_AllProductNumberFromSearchResult().Count;
            get_searchKeyTextBox().Clear();
            get_searchKeyTextBox().Click();
            get_searchKeyTextBox().SendKeys(productNumber);
            get_TypeSelectDropDown().Click();
            for (int i = 0; i < 10; i++)
            {
                if (get_AllProductNumberFromSearchResult().Count > 12)
                {
                    break;
                }
                Thread.Sleep(2000);
            }
        }

        public void SearchIntegrationmonitor(string productNumber, string dropDownValue)
        {
            SelectValueFromDropDown(dropDownValue);
            System.Diagnostics.Debug.WriteLine(get_ProductResult("14")[0].Text);
            System.Diagnostics.Debug.WriteLine(get_ProductResult("6")[0].Text);
            System.Diagnostics.Debug.WriteLine(get_ProductResult("14")[0].Text);
            System.Diagnostics.Debug.WriteLine(get_ProductResult("12")[0].Text);
            Assert.IsTrue(get_ProductResult("6")[0].Text.Contains(dropDownValue));
            Assert.IsTrue(get_ProductResult("14")[0].Text.Contains(productNumber));
            Assert.IsTrue(get_ProductResult("12")[0].Text.Contains("Processed"));
        }

        public void SelectValueFromDropDown(string dropDownValue)
        {
            get_TypeSelectDropDown().Click();
            get_TypeSelectDropDown().SendKeys(dropDownValue);
            SearchButtonHeaderElement.Click();
            Thread.Sleep(3000);
        }

        public void LaunchTabAndNavigate()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("window.open('','_blank');");
        }

        public string get_CurrentWindowHandle()
        {
            return driver.CurrentWindowHandle;
        }

        public void CloseIntegrationmonitorAndBackToPp()
        {
            driver.Close();
            driver.SwitchTo().Window(currentWindow);
        }

        public void SwitchToNewWindow()
        {
            List<string> allWindows = driver.WindowHandles.ToList();

            foreach (var win in allWindows)
            {
                driver.SwitchTo().Window(win);
            }
        }
    }
}
