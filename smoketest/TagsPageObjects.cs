using HP.LFT.Report;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SITSmokeTests
{
    public class TagsPageObjects
    {
        IWebDriver driver;
        private WebDriverWait wait;
        private WebDriverWait extWait;
        GeneralMethods generalMethods;

        public TagsPageObjects(IWebDriver drivername)
        {
            PageFactory.InitElements(drivername, this);

            driver = drivername;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            extWait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            generalMethods = new GeneralMethods();
        }

        [FindsBy(How = How.XPath, Using = ".//*[@id='ctl00_display']/a")]
        public IWebElement DisplayOrders { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='txtSelectedOrder']")]
        public IWebElement OrderNumberTextBox { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='ctl00_phMain_SearchButton']")]
        public IWebElement SearchButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='lbPlanningMarketToD']/option[1]")]
        public IWebElement All_PMAndTOD { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='PTOReSend']")]
        public IWebElement PTResendButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='ITReSend']")]
        public IWebElement ITResendButton { get; set; }

        [FindsBy(How=How.XPath, Using = ".//*[@id='ctl00_btnBackTop']")]
        public IWebElement BackLink { get; set; }


        public bool PTResend()
        {
            generalMethods.WebButton_Click(PTResendButton);
            ClickDetailsButton();
            return VerifyDataInTag();
        }

        public bool ITResend()
        {
            generalMethods.WebButton_Click(ITResendButton);
            ClickDetailsButton();
            return VerifyDataInTag();
        }

        private bool VerifyDataInTag()
        {
            IWebElement planningMarketRow;
            List<IWebElement> currencyRows;

            List<IWebElement> priceTagOrderLines = driver.FindElements(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div")).ToList();

            for (int priceTagOrderLine = 4; priceTagOrderLine < priceTagOrderLines.Count; priceTagOrderLine++)
            {
                int articleRowNo = 0;
                Dictionary<string, string> prices = new Dictionary<string, string>();
                planningMarketRow = driver.FindElement(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div[" + priceTagOrderLine + "]/div[1]"));
                string planningMarketName = planningMarketRow.Text.Split('\r')[0];

                //Expand the Planning Market
                driver.FindElement(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div[" + priceTagOrderLine + "]/div[1]/div[1]/div/div[@class='icon iconClose']")).Click();

                currencyRows = driver.FindElements(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div[" + priceTagOrderLine + "]/div[@ng-repeat='price in item.Prices']")).ToList();

                //Prices
                for (int currencyRow = 0; currencyRow < currencyRows.Count; currencyRow++)
                {
                    int row = currencyRow + 2;
                    string price = driver.FindElement(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div[" + priceTagOrderLine + "]/div[" + row + "]/div/div[1]")).Text;
                    string[] temp = price.Split(' ');

                    prices.Add(temp[0].Trim(), price.Substring(temp[0].Length).Trim());
                    articleRowNo = row;
                }

                //Adjust the Article Row if there are no prices for the PM
                articleRowNo = (articleRowNo == 0) ? articleRowNo + 2 : ++articleRowNo;

                //Expand the article
                driver.FindElement(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div[" + priceTagOrderLine + "]/div[" + articleRowNo + "]/div[1]/div[1]/div/div[@class='icon iconClose']")).Click();

                string articleName = driver.FindElement(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div[" + priceTagOrderLine + "]/div[" + articleRowNo + "]/div[1]/div[1]/span")).Text;

                List<IWebElement> sizes = driver.FindElements(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div[" + priceTagOrderLine + "]/div[" + articleRowNo + "]/div[@class='ng-scope']")).ToList();

                //Sizes
                for (int sizeRow = 0; sizeRow < sizes.Count; sizeRow++)
                {
                    int t = sizeRow + 1;
                    string sizeDetails = driver.FindElement(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div[" + priceTagOrderLine + "]/div[" + articleRowNo + "]/div[@class='ng-scope'][" + t + "]/div/div")).Text;
                    t++;
                    IWebElement preveiwButton = driver.FindElement(By.XPath(".//*[@id='aspnetForm']/div[5]/div[2]/div/div[" + priceTagOrderLine + "]/div[" + articleRowNo + "]/div[" + t + "]/div/div[4]/input"));
                    preveiwButton.Click();

                    if (!VerifyPricesAndSizeInTag(planningMarketName, articleName, sizeDetails, prices))
                        return false;

                    driver.FindElement(By.XPath("html/body/div[3]/div/div/a")).Click();
                    Thread.Sleep(2000);
                }
            }
            return true;
        }

        private bool VerifyPricesAndSizeInTag(string planningMarketName, string articleName, string sizeDetails, Dictionary<string, string> prices)
        {
            string[] sizeRegionAndSizes = sizeDetails.Split(',');
            int i = 0;
            
            List<IWebElement> sizeRows = driver.FindElements(By.XPath(".//*[@id='tabs-0']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[@class='sizeRow']")).ToList();

            foreach (IWebElement sizeRow in sizeRows)
            {
                //sizeNameAndSize will have the text like "EUR 34". So split it and verify the sizeRegion and size
                string[] temp = sizeRegionAndSizes[i].Trim().Split(' ');
                int row = i + 2;
                string sizeRegion = driver.FindElement(By.XPath(".//*[@id='tabs-0']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[" + row + "]/td/span[1]")).Text.Trim();
                List<IWebElement> sizeNames = driver.FindElements(By.XPath(".//*[@id='tabs-0']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[" + row + "]/td/span[2]/span[@class='sizeName highlight']")).ToList();
                

                string size;

                if (sizeNames.Count > 0)
                    size = driver.FindElement(By.XPath(".//*[@id='tabs-0']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[" + row + "]/td/span[2]/span[@class='sizeName highlight']")).Text.Trim();
                else
                    size = driver.FindElement(By.XPath(".//*[@id='tabs-0']/table/tbody/tr/td/table[2]/tbody/tr/td[1]/table/tbody/tr[" + row + "]/td/span[2]/span[@class='sizeName']")).Text.Trim();

                if ((temp[0].Trim() != sizeRegion) || (temp[1].Trim() != size))
                {
                    Reporter.ReportEvent("Selenium_PricesAndSizeInTag Script", "PM: " + planningMarketName + ", Article: " + articleName + ", Size: " + temp[0].Trim() + " " + temp[1].Trim() + " not found in tag. Expected: " + temp[0].Trim() + " " + temp[1].Trim() + ". Acutal: " + sizeRegion + " " + size);
                    return false;
                }
                else
                    Reporter.ReportEvent("Selenium_PricesAndSizeInTag Script", "PM: " + planningMarketName + ", Article: " + articleName + ", Size: " + temp[0].Trim() + " " + temp[1].Trim() + " found in tag.");
                i++;
            }

            //Prices
            List<IWebElement> priceRows = driver.FindElements(By.XPath(".//*[@id='tabs-0']/table/tbody/tr/td/table[1]/tbody/tr")).ToList();

            foreach(IWebElement priceRow in priceRows)
            {
                string priceCurrency = priceRow.FindElement(By.ClassName("priceCurrency")).Text;
                string price = priceRow.FindElement(By.ClassName("pricePrice")).Text;
                if (price != prices[priceCurrency])
                {
                    Reporter.ReportEvent("Selenium_PricesAndSizeInTag Script", "PM: " + planningMarketName + ", Article: " + articleName + ", Price: " + priceCurrency + " " + prices[priceCurrency] + " not found in tag. Expected: " + priceCurrency + " " + prices[priceCurrency] + ". Actual: " + priceCurrency + " " + price);
                    return false;
                }
                else
                    Reporter.ReportEvent("Selenium_PricesAndSizeInTag Script", "PM: " + planningMarketName + ", Article: " + articleName + ", Price: " + priceCurrency +  " " + prices[priceCurrency] + " found in tag.");
            }

            return true;
        }

        private void ClickDetailsButton()
        {
            List<IWebElement> rows = driver.FindElements(By.XPath(".//*[@id='ctl00_phMain_grdPTO']/tbody/tr")).ToList();
            int detailsButtonIndex = rows.Count;

            driver.FindElement(By.XPath(".//*[@id='ctl00_phMain_grdPTO_ctl" + detailsButtonIndex + "_btnDetails']")).Click();
        }

        public void NavigateAndSearchOrder(string orderNumber)
        {
            string Tags_Url = ConfigUtils.Read("URL_Tags");

            driver.Navigate().GoToUrl(Tags_Url);
            Thread.Sleep(5000);

            generalMethods.WebLink_Click(DisplayOrders);
            Thread.Sleep(2000);

            OrderNumberTextBox.SendKeys(orderNumber);
            Thread.Sleep(2000);
            Actions builder = new Actions(driver);
            builder.SendKeys(Keys.Tab).Perform();

            generalMethods.WebButton_Click(SearchButton);
            Thread.Sleep(3000);
        }
    }
}