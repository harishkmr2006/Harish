using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using HP.LFT.Report;
using System.Drawing;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using OpenQA.Selenium.Interactions;

namespace SITSmokeTests
{
    public class OFUPageObjects
    {
        IWebDriver driver;
        WebDriverWait wait;

        public OFUPageObjects(IWebDriver drivername)
        {
            PageFactory.InitElements(drivername, this);
            driver = drivername;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
        }

        [FindsBy(How = How.XPath, Using = "html/body/div[7]/div[2]/table/tbody/tr/td[3]/a")]
        public IWebElement CompositionLink { get; set; }

        [FindsBy(How = How.XPath, Using = "html/body/div[7]/div[2]/table/tbody/tr/td[3]")]
        public IWebElement CompositionTab { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".icon-pencil")]
        public IWebElement EditComposition { get; set; }

        [FindsBy(How = How.CssSelector, Using = ".placement")]
        public IWebElement Placement { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='OrderedDevelopmentOptions_0__ExclusiveOfId']")]
        public IWebElement ExclusiveOfDropDown { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='saveButton']")]
        public IWebElement SaveButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='secureButton']")]
        public IWebElement SecureButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//div[@class=\"phholder\"]/div")]
        public IWebElement SingleSignon_Email { get; set; }

        [FindsBy(How = How.XPath, Using = ".//input[@value=\"Next\"]")]
        public IWebElement SingleSignon_Nextbutton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//div[contains(text(),\"Work or school account\")]")]
        public IWebElement SingleSignon_WorkAccount { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='ORDER_LIST']")]
        public IWebElement OFU_OrderList { get; set; }


        public void SingleSignOn(string eMail)
        {
            GeneralMethods sGMethods = new GeneralMethods();
            try
            {
                Actions saction = new Actions(driver);
                saction.SendKeys(SingleSignon_Email, Keys.Enter).SendKeys(SingleSignon_Email, eMail).Build().Perform();
              //  sGMethods.WebEdit_SetValue(SingleSignon_Email, "Single sign eMail", eMail);
                Thread.Sleep(2000);
                sGMethods.WebButton_Click(SingleSignon_Nextbutton, "SingleSignon_Nextbutton");
                Thread.Sleep(2000);
                try
                {
                    sGMethods.WebButton_Click(SingleSignon_WorkAccount, "SingleSignon_WorkAccount");
                }
                catch
                {

                }
               // Thread.Sleep(2000);
               // driver.Url = "https://ofu-sit.azurewebsites.net/orders";
              //  driver.Navigate().GoToUrl("https://ofu-sit.azurewebsites.net/orders");
               // wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='ORDER_LIST']")));
                Thread.Sleep(20000);
                string surl = driver.Url;
                Thread.Sleep(2000);
                wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='ORDER_LIST']")));
            }
            catch
            {

            }
        }

        public string GetCompositionStatus()
        {
            return driver.FindElement(By.XPath(".//*[@id='order-information-table']/tbody/tr[4]/td[4]")).Text.Trim();
        }

        public bool VerifyBOM()
        {
            IWebElement table = driver.FindElement(By.XPath(".//div[@class='detailedcomposition']/table/tbody"));
            string Position = "Shell";
            string Placement = "BAG";
            string materialComposition = "100% BUFFALO LEATHER";
            int positionIndex = 4;
            int placementIndex = 5;
            int materialCompositionIndex = 6;
            string cellData;
            bool BOMMatches = true;

            List<IWebElement> rows = table.FindElements(By.TagName("tr")).ToList();

            for (int row = 1; row <= rows.Count; row++)
            {
                cellData = driver.FindElement(By.XPath(".//div[@class='detailedcomposition']/table/tbody/tr[" + row + "]/td[" + positionIndex + "]")).Text.Trim();
                if (Position != cellData)
                {
                    BOMMatches = false;
                    Reporter.ReportEvent("Selenium_VerifyCITStatus Script failed", "Positions do not match on row " + row + ". Expected: " + Position + ", Actual: " + cellData, Status.Failed);
                    break;
                }

                cellData = driver.FindElement(By.XPath(".//div[@class='detailedcomposition']/table/tbody/tr[" + row + "]/td[" + placementIndex + "]/span/span[@class='placement']")).Text.Trim();
                if (Placement != cellData)
                {
                    BOMMatches = false;
                    Reporter.ReportEvent("Selenium_VerifyCITStatus Script failed", "Placements do not match on row " + row + ". Expected: " + Placement + ", Actual: " + cellData, Status.Failed);
                    break;
                }

                cellData = driver.FindElement(By.XPath(".//div[@class='detailedcomposition']/table/tbody/tr[" + row + "]/td[" + materialCompositionIndex + "]")).Text.Trim();
                if (materialComposition != cellData)
                {
                    BOMMatches = false;
                    Reporter.ReportEvent("Selenium_VerifyCITStatus Script failed", "Material Compositions do not match on row " + row + ". Expected: " + materialComposition + ", Actual: " + cellData, Status.Failed);
                    break;
                }
            }

            return BOMMatches;
        }

        public bool VerifyCompositionTabBackgroundColour()
        {
            string expectedBackgroundColor = "#468847";
            string color = driver.FindElement(By.XPath("html/body/div[7]/div[2]/table/tbody/tr/td[3]")).GetCssValue("background-color").ToString();

            string[] colors = color.Replace("rgba(", "").Replace(")", "").Split(',');

            int red = Convert.ToInt32(colors[0]);
            int green = Convert.ToInt32(colors[1].Trim());
            int blue = Convert.ToInt32(colors[2].Trim());
            int alpha = Convert.ToInt32(colors[3].Trim());

            Color c = Color.FromArgb(alpha, red, green, blue);

            string hex = string.Format("#{0}", c.Name.Substring(c.Name.Length - 6));

            if (expectedBackgroundColor != hex)
            {
                Reporter.ReportEvent("Selenium_VerifyCITStatus Script failed", "Background Colors of Composition Tab does not match. Expected: " + expectedBackgroundColor + ", Actual: " + hex, Status.Failed);
                return false;
            }
            else
                Reporter.ReportEvent("Selenium_VerifyCITStatus Script passed", "Background Colors of Composition Tab is Green");

            return true;
        }
    }
}