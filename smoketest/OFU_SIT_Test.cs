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
    class OFU_SIT_Test
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

        public OFU_SIT_Test()
        {
            var binary = new FirefoxBinary(ConfigUtils.Read("FirefoxPath"));
            // string path = ReadFirefoxProfile();
            FirefoxProfile ffprofile = new FirefoxProfile();
          //  ffprofile.SetPreference("javascript.enabled","false");
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

       //  [Test]
        public void TC29_Selenium_SmokeTest_OFU()
        {

            xCellFileHelper = new ExcelHelper(datafilePath, 1);
            string directive = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "SingleSignOn_Email");
           string OFU_Url = ConfigUtils.Read("URL_OFU3");
           driver.Navigate().GoToUrl(OFU_Url);
           driver.Manage().Window.Maximize();
           Console.WriteLine("OFU");
            OFUPage OFUTest = new OFUPage(driver);
            OFUTest.EnterEmail.Click();
            OFUTest.EnterEmail.SendKeys(directive);
            OFUTest.ClickNextButton.Click();
            OFUTest.get_OFUPageLoad("Order Follow Up");
            Thread.Sleep(500);
            OFUTest.ordersearch("999945");
           // OFUTest.SearchOrder.SendKeys(Keys.Enter);
            Console.WriteLine("Clicked");

        }
    }
}
