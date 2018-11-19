using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using System;

namespace BOPO.NUnit.ParallelTests
{
    public class EnvironmentSetUp
    {
        public EnvironmentSetUp()
        {
            try
            {
                string sprojectpath = this.GetType().Assembly.Location;
                string[] swrokingdirectorypath = sprojectpath.Split('\\');
                string sdatafiledirectory = "";
                foreach (string sitem in swrokingdirectorypath)
                {
                    if (sitem == "BOPO.NUnit.ParallelTests")
                    {
                        break;
                    }
                    sdatafiledirectory = sdatafiledirectory + sitem + "\\";
                }
                System.Environment.SetEnvironmentVariable("ProjectWorkingDirectory", sdatafiledirectory);
            }

            catch (Exception ex)
            {
                throw ex;
            }


        }

        public IWebDriver Setup_Driver(string BrowserName, IWebDriver Driver, string ExecutionType)
        {
            try
            {
                if (BrowserName.Equals("ie"))
                {
                    System.Environment.SetEnvironmentVariable("webdriver.ie.driver", System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "IEDriverServer.exe");
                    if (ExecutionType == "Remote")
                    {
                        Driver = GetIERemoteDriver();
                        Driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(25));
                        return Driver;
                    }
                    else
                    {
                        Driver = GetIELocalDriver();
                        Driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
                        return Driver;
                    }
                }

                else if (BrowserName.Equals("chrome"))
                {
                    Driver = new ChromeDriver();
                    return Driver;
                }
                else
                {
                    if (ExecutionType == "Remote")
                    {
                        Driver = GetFFRemoteDriver();
                        Driver.Manage().Window.Maximize();
                        Driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));
                        return Driver;
                    }

                    else
                    {
                        Driver = GetFFLocalDriver();
                        Driver.Manage().Cookies.DeleteAllCookies();
                        Driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
                        return Driver;
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private static FirefoxProfile CreateFirefoxProfile()
        {
            try
            {
                var FFprofile = new FirefoxProfile();
                FFprofile.AddExtension("Tools/modify_headers-0.7.1.1-fx.xpi");
                FFprofile.SetPreference("general.useragent.override", "UA-STRING");
                FFprofile.SetPreference("extensions.modify_headers.currentVersion", "0.7.1.1-signed");
                FFprofile.SetPreference("modifyheaders.headers.count", 1);
                FFprofile.SetPreference("modifyheaders.headers.action0", "Add");
                FFprofile.SetPreference("modifyheaders.headers.name0", "SampleHeader");
                FFprofile.SetPreference("modifyheaders.headers.value0", "test1234");
                FFprofile.SetPreference("modifyheaders.headers.enabled0", true);
                FFprofile.SetPreference("modifyheaders.config.active", true);
                FFprofile.SetPreference("modifyheaders.config.alwaysOn", true);
                FFprofile.SetPreference("modifyheaders.config.start", true);

                return FFprofile;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private static IWebDriver GetFFLocalDriver()
        {
            try
            {
                var capabilities = new DesiredCapabilities();
                var profile = CreateFirefoxProfile();
                capabilities.SetCapability("username", "TEMPHAHUC");
                capabilities.SetCapability("accessKey", "Citrix1234");
                capabilities.SetCapability(FirefoxDriver.BinaryCapabilityName, "C:\\Program Files (x86)\\Mozilla Firefox ESR v45.7.0\\firefox.exe");
                capabilities.SetCapability("network.automatic-ntlm-auth.trusted-uris", Environment.UserDomainName);
                capabilities.SetCapability("signon.autologin.proxy", true);
                var Localdriver = new FirefoxDriver(capabilities);
                return Localdriver;
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static IWebDriver GetFFRemoteDriver()
        {
            try
            {
                var capabilities = new DesiredCapabilities("Firefox", "45.0.2", new Platform(PlatformType.Windows));
                var profile = CreateFirefoxProfile();
                capabilities = DesiredCapabilities.Firefox();
                capabilities.SetCapability(FirefoxDriver.BinaryCapabilityName, "C:\\Program Files (x86)\\Mozilla Firefox ESR v45.7.0\\firefox.exe");
                capabilities.SetCapability("network.automatic-ntlm-auth.trusted-uris", Environment.UserDomainName);
                capabilities.SetCapability("signon.autologin.proxy", true);
                // capabilities.SetCapability(FirefoxDriver.ProfileCapabilityName, profile);
                var RemoteFFDriver = new RemoteWebDriver(new Uri("http://10.64.246.48:4444/wd/hub"), capabilities);

                return RemoteFFDriver;
            }

            catch (Exception ex)
            {
                throw ex;
            }

        }

        private static IWebDriver GetIELocalDriver()
        {
            try
            {
                var IEoptons = new InternetExplorerOptions();
                IEoptons.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                var LocalIEDriver = new InternetExplorerDriver(IEoptons);

                return LocalIEDriver;
            }

            catch (Exception ex)
            {
                throw ex;
            }


        }

        private static IWebDriver GetIERemoteDriver()
        {
            try
            {
                var capabilities = new DesiredCapabilities();
                capabilities = DesiredCapabilities.InternetExplorer();
                capabilities.SetCapability(CapabilityType.BrowserName, "internet explorer");
                capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Any));
                capabilities.SetCapability("ignoreProtectedModeSettings", true);
                var RemoteIEDriver = new RemoteWebDriver(new Uri("http://10.60.210.54:5555/wd/hub"), capabilities);

                return RemoteIEDriver;
            }

            catch (Exception ex)
            {
                throw ex;
            }

        }



    }
}
