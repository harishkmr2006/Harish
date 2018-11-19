using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Data;


namespace BOPO.NUnit.ParallelTests
{
    public class TestBase
    {
        public IWebDriver driver { get; set; }
        public IWebDriver IEdriver { get; set; }
        public static IWebDriver DriverIE;
        public static IWebDriver DriverFF;
        public static IWebDriver DriverCleardown;
        public String sFilename;
        public String sIEFilename;
        public static String sbrowsertype;
        public static DataTable stable;
        public static List<Result> TestRESULT = new List<Result>();
        public static string sTestName = "";
        // public static String sFilenameIE = @"C:\Project\DataIE.xlsx";
        /* public static IEnumerable<String> BrowserTypetoExecute()
         {
             String[] browsers = { "firefox", "ie" };
             foreach (String browser in browsers)
             {
                 yield return browser;
             }
         }*/
        [SetUp]
        public void Browserconfig()
        {
            System.Environment.SetEnvironmentVariable("Browser1", ConfigUtils.Read("Browser1"));
            System.Environment.SetEnvironmentVariable("Browser2", ConfigUtils.Read("Browser2"));

        }

        //public IWebDriver SetUp(string sbrowsername, IWebDriver driver)
        //{

        //    if (sbrowsername.Equals("ie"))
        //    {
        //        // System.Environment.SetEnvironmentVariable("webdriver.ie.driver", "C:\\Selenium\\IEDriverServer.exe");

        //        // DesiredCapabilities capabilities = new DesiredCapabilities();
        //        // capabilities = DesiredCapabilities.InternetExplorer();
        //        // capabilities.SetCapability(CapabilityType.BrowserName, "internet explorer");
        //        //  capabilities.SetCapability(CapabilityType.Version, "11");
        //        //  capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
        //        // IEdriver = new InternetExplorerDriver();
        //        if (driver != null)
        //        {
        //            Thread.Sleep(20000);
        //        }
        //        driver = new InternetExplorerDriver();
        //        driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(25));
        //        return driver;
        //    }

        //    else if (sbrowsername.Equals("chrome"))
        //    {
        //        driver = new ChromeDriver();
        //        return driver;


        //    }
        //    else
        //    {
        //        // sFilename = ConfigUtils.Read("FFDATA");
        //        sIEFilename = ConfigUtils.Read("IEDATA");
        //        DesiredCapabilities capabilities = new DesiredCapabilities();
        //        capabilities = DesiredCapabilities.Firefox();
        //        // capabilities.SetCapability(CapabilityType.BrowserName, "firefox");
        //        // capabilities.SetCapability(CapabilityType.Version, "38");

        //        capabilities.SetCapability("profile", new FirefoxProfile());
        //        capabilities.SetCapability("binary", "C:\\Program Files (x86)\\Mozilla Firefox ESR v38.7.0\\firefox.exe");
        //        capabilities.SetCapability("FirefoxBinary", "C:\\Program Files (x86)\\Mozilla Firefox ESR v38.7.0\\firefox.exe");
        //        capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
        //        Console.WriteLine("Node value is");
        //        Console.WriteLine(ConfigUtils.Read("remotenode1"));
        //        // new FirefoxDriver(new FirefoxBinary("C:\\Program Files (x86)\\Mozilla Firefox ESR v38.7.0\\firefox.exe"), new FirefoxProfile(), TimeSpan.FromMinutes(10));
        //        // driver = new RemoteWebDriver(new Uri("http://9.193.83.57:5555/wd/hub"), capabilities);

        //        driver = new RemoteWebDriver(new Uri(ConfigUtils.Read("remotenode1")), capabilities);

        //        // driver = new FirefoxDriver();
        //        // sFilename = "C:\\Project\\Data.xlsx";

        //        driver.Manage().Window.Maximize();
        //        driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(25));

        //        //  driver.Navigate().GoToUrl(ConfigUtils.Read("URL"));
        //        // Console.WriteLine("Opened URL");
        //        return driver;

        //    }


        //}

        [TearDown]
        public void Clearup()
        {
            try
            {

                if (sbrowsertype.Equals("ie"))
                {
                    System.Environment.SetEnvironmentVariable("TestCaseName", System.Environment.GetEnvironmentVariable("TestCaseName2"));
                }
                else
                {
                    System.Environment.SetEnvironmentVariable("TestCaseName", System.Environment.GetEnvironmentVariable("TestCaseName1"));
                }
                // Reporter swriteresult = new Reporter();
                // swriteresult.WriteResults(sTestName);
                if (sbrowsertype.Equals("ie"))
                {
                    // IEdriver.Close();

                }
                else
                {
                    //driver.Close();
                }
                // IEdriver.Close();
                //driver.Close();
                DriverCleardown.Close();


            }
            catch
            {
                // System.Environment.SetEnvironmentVariable("TestCaseName", "CurrentFailedTest");
                // Reporter swriteresult = new Reporter();
                // swriteresult.WriteResults(sTestName);
            }
        }

    }
}
