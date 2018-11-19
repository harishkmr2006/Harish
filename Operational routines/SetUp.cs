using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using System;
using System.IO;
using System.Linq;

namespace BOPO.NUnit.ParallelTests
{
    public class SetUp
    {
        public SetUp()
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

        public IWebDriver Setup_Driver(string sbrowsername, IWebDriver driver, string env)
        {
            if (sbrowsername.Equals("ie"))
            {
                System.Environment.SetEnvironmentVariable("webdriver.ie.driver", System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "IEDriverServer.exe");
                if (env == "Remote")
                {


                    DesiredCapabilities capabilities = new DesiredCapabilities();
                    capabilities = DesiredCapabilities.InternetExplorer();
                    capabilities.SetCapability(CapabilityType.BrowserName, "internet explorer");
                    // capabilities.SetCapability(CapabilityType.Version, "11.0.9600.18282");
                    capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Any));
                    capabilities.SetCapability("ignoreProtectedModeSettings", true);
                    // IEdriver = new InternetExplorerDriver();
                    // if (driver != null)
                    //  {
                    //     Thread.Sleep(20000);
                    //  }
                    // driver = new RemoteWebDriver(new Uri("http://localhost:4444/wd/hub"), capabilities);
                    driver = new RemoteWebDriver(new Uri("http://10.60.208.235:5555/wd/hub"), capabilities);
                    //driver = new InternetExplorerDriver();
                    driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(25));
                    return driver;
                }
                else
                {
                    InternetExplorerOptions IEoptons = new InternetExplorerOptions();
                    IEoptons.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                    driver = new InternetExplorerDriver(IEoptons);
                    driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
                    return driver;
                }
            }

            else if (sbrowsername.Equals("chrome"))
            {
                driver = new ChromeDriver();
                return driver;


            }
            else
            {
                if (env == "Remote")
                {
                    // sFilename = ConfigUtils.Read("FFDATA");
                    // sIEFilename = ConfigUtils.Read("IEDATA");
                    DesiredCapabilities capabilities = new DesiredCapabilities("Firefox", "45.0.2", new Platform(PlatformType.Windows));
                    capabilities = DesiredCapabilities.Firefox();
                    //DesiredCapabilities capabilities = new DesiredCapabilities();
                    // capabilities = DesiredCapabilities.Firefox();
                    // capabilities.SetCapability(CapabilityType.BrowserName, "firefox");
                    // capabilities.SetCapability(CapabilityType.Version, "38");

                    // capabilities.SetCapability("profile", new FirefoxProfile());
                    // capabilities.SetCapability("binary", "C:\\Program Files (x86)\\Mozilla Firefox ESR v45.2.0\\firefox.exe");
                    // capabilities.SetCapability("FirefoxBinary", new FirefoxBinary("C:\\Program Files (x86)\\Mozilla Firefox ESR v45.2.0\\firefox.exe"));
                    // capabilities.SetCapability("FirefoxProfile", new FirefoxProfile());
                    // capabilities.SetCapability(CapabilityType.Platform, new Platform(PlatformType.Windows));
                    // Console.WriteLine("Node value is");
                    // Console.WriteLine(ConfigUtils.Read("remotenode1"));
                    // new FirefoxDriver(new FirefoxBinary("C:\\Program Files (x86)\\Mozilla Firefox ESR v38.7.0\\firefox.exe"), new FirefoxProfile(), TimeSpan.FromMinutes(10));
                    driver = new RemoteWebDriver(new Uri("http://10.64.246.48:4444/wd/hub"), capabilities);

                    //driver = new RemoteWebDriver(new Uri("http://localhost:4444/wd/hub"), capabilities);

                    // driver = new FirefoxDriver();
                    // sFilename = "C:\\Project\\Data.xlsx";

                    driver.Manage().Window.Maximize();
                    driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(10));

                    //  driver.Navigate().GoToUrl(ConfigUtils.Read("URL"));
                    // Console.WriteLine("Opened URL");
                    return driver;
                }
                else

                {
                    try
                    {


                        DesiredCapabilities capabilities = new DesiredCapabilities();

                        capabilities.SetCapability("username", "TEMPSRMAL");
                        capabilities.SetCapability("accessKey", "Citrix1234");
                        capabilities.SetCapability(FirefoxDriver.BinaryCapabilityName, "C:\\Program Files (x86)\\Mozilla Firefox ESR v45.5.0\\firefox.exe");
                        capabilities.SetCapability("network.automatic-ntlm-auth.trusted-uris", Environment.UserDomainName);
                        capabilities.SetCapability("signon.autologin.proxy", true);


                        string spath = "";
                        bool bfoundprofile = false;

                        Console.WriteLine("firefox profile manager profiles path");
                        FirefoxProfileManager profileManager = new FirefoxProfileManager();

                        foreach (string ffpro in profileManager.ExistingProfiles)
                        {

                            if (ffpro.Contains("default"))
                            {
                                spath = ffpro;
                                bfoundprofile = true;
                                break;
                            }
                        }
                        if (bfoundprofile == false)
                        {
                            spath = ReadFirefoxProfile();
                        }

                        try
                        {

                            /* string browserversion = ConfigUtils.Read("Browserversion");
                             string pathToCurrentUserProfiles = Environment.ExpandEnvironmentVariables("%APPDATA%") + @"\Mozilla\Firefox\Profiles"; // Path to profile
                             string[] pathsToProfiles = Directory.GetDirectories(pathToCurrentUserProfiles, "*.default*", SearchOption.TopDirectoryOnly);
                             if (pathsToProfiles.Length != 0)
                             {
                                 FirefoxProfile profile = new FirefoxProfile(pathsToProfiles[0]);
                                 profile.SetPreference("browser.tabs.loadInBackground", false); // set preferences you need
                                                                                                // profile.SetPreference("network.http.phishy-userpass-length", 255);
                                                                                                // profile.SetPreference("network.automatic-ntlm-auth.trusteduris", "hm.com");
                                                                                                // profile.SetPreference("signon.autologin.proxy", true);
                                                                                                // profile.SetPreference("applicationCacheEnabled", false);
                                                                                                //Driver = new FirefoxDriver(new FirefoxBinary(), profile, TimeSpan.FromSeconds(10));
                                 driver = new FirefoxDriver(new FirefoxBinary("C:\\Program Files (x86)\\Mozilla Firefox ESR v"+browserversion+"\\firefox.exe"), profile, TimeSpan.FromMinutes(10));
                                 // driver = new FirefoxDriver(new FirefoxBinary(), profile, TimeSpan.FromMinutes(10));
                                 //driver = new FirefoxDriver(profile);
                             }
                             else
                             {
                                 driver = new FirefoxDriver(new FirefoxBinary("C:\\Program Files (x86)\\Mozilla Firefox ESR v" + browserversion + "\\firefox.exe"), new FirefoxProfile(spath), TimeSpan.FromMinutes(10));
                             }
                             //  System.setProperty("webdriver.firefox.bin", "Path to binary");

                             */
                            var binary = new FirefoxBinary("C:\\Program Files (x86)\\Mozilla Firefox ESR v45.5.0\\firefox.exe");
                            string path = ReadFirefoxProfile();
                            FirefoxProfile ffprofile = new FirefoxProfile(path);
                            driver = new FirefoxDriver(binary, ffprofile);


                            // System.setProperty
                            //driver = new FirefoxDriver(new FirefoxBinary("C:\\Program Files (x86)\\Mozilla Firefox ESR v45.3\\firefox.exe"), new FirefoxProfile(), TimeSpan.FromMinutes(10));
                            //driver = new FirefoxDriver();
                            driver.Manage().Cookies.DeleteAllCookies();
                            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(15));
                            return driver;





                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            return driver;
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        return driver;
                    }
                }

            }


        }
        public static string ReadFirefoxProfile()
        {
            string apppath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);

            string mozilla = System.IO.Path.Combine(apppath, "Mozilla");

            bool exist = System.IO.Directory.Exists(mozilla);

            if (exist)
            {

                string firefox = System.IO.Path.Combine(mozilla, "firefox");

                if (System.IO.Directory.Exists(firefox))
                {
                    string prof_file = System.IO.Path.Combine(firefox, "profiles.ini");

                    bool file_exist = System.IO.File.Exists(prof_file);

                    if (file_exist)
                    {
                        StreamReader rdr = new StreamReader(prof_file);

                        string resp = rdr.ReadToEnd();

                        string[] lines = resp.Split(new string[] { "\r\n" }, StringSplitOptions.None);

                        string location = lines.First(x => x.Contains("Path=")).Split(new string[] { "=" }, StringSplitOptions.None)[1];

                        string prof_dir = System.IO.Path.Combine(firefox, location);

                        return prof_dir;
                    }
                }
            }
            return "";
        }

    }
}
