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
using OpenQA.Selenium.Firefox;

namespace SITSmokeTests
{

    class OFUPage
    {
        IWebDriver driver;
        private WebDriverWait wait;

        private WebDriverWait extWait;

        GeneralMethods sGMethods;

        public OFUPage(IWebDriver drivername)
        {
            PageFactory.InitElements(drivername, this);
            driver = drivername;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            extWait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            sGMethods = new GeneralMethods();
        }

        [FindsBy(How = How.XPath, Using = "//div[contains(@class,'placeholderContainer')]/input")]
        public IWebElement EnterEmail { get; set; }

        [FindsBy(How=How.XPath, Using = "//div[contains(@class,'col-xs-24')]/input")]
        public IWebElement ClickNextButton { get; set; }

        [FindsBy(How=How.XPath,Using = ".//div[@class='input-group mb-2 mr-sm-2 mb-sm-0 w-100 quick-search']/input")]
        public IWebElement SearchOrder { get; set; }

        

        public IWebElement get_value_ProductClassificationCategories(string productClassification)
        {
            return wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(
                "//ul//*[text()='" + productClassification + "']")));
        }


        public void entermail(string mail)
        {
            Thread.Sleep(200);
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("document.getElementById('i0116').value='" + mail + "';");
            driver.FindElement(By.XPath("//div[contains(@class,'placeholderContainer')]/input")).Click();
            driver.FindElement(By.XPath("//div[contains(@class,'placeholderContainer')]/input")).SendKeys(" ");
            driver.FindElement(By.XPath("//div[contains(@class,'placeholderContainer')]/input")).SendKeys("\b");



        }


        public void ordersearch(string ordernnumber)
        {

            Thread.Sleep(3000);
          //  Console.WriteLine(driver.s)
            var binary = new FirefoxBinary(ConfigUtils.Read("FirefoxPath"));
            // string path = ReadFirefoxProfile();
            FirefoxProfile ffprofile = new FirefoxProfile();
            //  ffprofile.SetPreference("javascript.enabled","false");
            IWebDriver driver1 = new FirefoxDriver(binary, ffprofile);
            //  driver = new InternetExplorerDriver(internetExplorerDriverServerDirectory: "\\srv10177\\TEMPSHTIW$\\Desktop\\SIT\\SITSmokeTests\\IEDriverServer.exe"); 
            //driver.Manage().Cookies.DeleteAllCookies();
            driver1.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
            driver1.Navigate().GoToUrl("https://ofu-sit.azurewebsites.net/orders");

            Thread.Sleep(3000);
            Console.WriteLine(driver1.PageSource);

            driver1.FindElement(By.XPath(".//div[@class='input-group mb-2 mr-sm-2 mb-sm-0 w-100 quick-search']/input")).SendKeys(ordernnumber);

            IWebElement element = driver.FindElement(By.XPath("//td[contains(.,'999945 - 1510')]"));
            String contents = (String)((IJavaScriptExecutor)driver).ExecuteScript("return      arguments[0].innerHTML;", element);
            
           

            //extWait
            //Until(ExpectedConditions.ElementIsVisible(By.XPath(".//td[2]/small/a/span/span[contains(@text(),'"+ordernnumber+"')]")));
            //var element = driver.FindElement(By.XPath("//a[contains(.,'999945') and contains(.,'1510')]"));
            //  var element = driver.FindElement(By.XPath(".//div[@class='input-group mb-2 mr-sm-2 mb-sm-0 w-100 quick-search']/input"));
            IJavaScriptExecutor executor = (IJavaScriptExecutor)driver;
            try
            {
               var s3 = executor.ExecuteScript("document.readyState");
                executor.ExecuteScript("document.getElementById('search-result-dropdown').getElementsByTagName('tr')[0].getElementsByTagName('a')[0].href;");

            }
            catch { }
            try
            {
                Console.WriteLine("%%%%%%%%%%%%%%%%%%%%" + executor.ExecuteScript("document.getElementById('search-result-dropdown').getElementsByTagName('tr')[0].getElementsByTagName('a')[0].href;"));

            }catch { }
            try
            {
                var loca = executor.ExecuteScript("document.getElementById('search-result-dropdown').getElementsByTagName('tr')[0].getElementsByTagName('a')[0].href;");

            }
            
            catch { }
            //Actions saction = new Actions(driver);
            //saction.SendKeys(element,Keys.Enter).Build().Perform();
            //  .SendKeys(SingleSignon_Email, eMail).Build().Perform();

            //driver.FindElement(By.XPath(".//div[@class='input-group mb-2 mr-sm-2 mb-sm-0 w-100 quick-search']/input")).Click();
            //driver.FindElement(By.XPath(".//div[@class='input-group mb-2 mr-sm-2 mb-sm-0 w-100 quick-search']/input")).SendKeys(Keys.Enter);

            //   driver.FindElements(By.CssSelector("ng-scope"))[1].Click();


            Thread.Sleep(3000);


        }


        public IWebElement get_OFUPageLoad(String name)
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("*//a[text()='"+name+"']")));
        }

    }
}
