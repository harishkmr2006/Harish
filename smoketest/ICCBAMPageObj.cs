using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.Verifications;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System.Threading;


using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.PageObjects;
using HP.LFT.Report;

namespace SITSmokeTests
{
    class ICCBAMPageObj
    {

        IWebDriver driver;
        private WebDriverWait wait;
        private WebDriverWait extWait;
        static string currentWindow;
        GeneralMethods sGMethods;
        public ICCBAMPageObj(IWebDriver drivername)
        {

            PageFactory.InitElements(drivername, this);

            driver = drivername;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            extWait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            sGMethods = new GeneralMethods();
        }

    

        [FindsBy(How = How.XPath, Using = ".//*[@id='selectEnv']")]
        public IWebElement Dropdown_Env { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='selectIO']")]
        public IWebElement Dropdown_Application { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='selectBO']")]
        public IWebElement Dropdown_BusinessObject { get; set; }

        [FindsBy(How = How.XPath, Using = "//table[@class=\"searchFieldsTable\"]//td[contains(.,\"ProductDevelopmentId\")]/following-sibling::td/input")]
        public IWebElement WebEdit_ProductID { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='main']//button[contains(.,\"Load BAM\")]")]
        public IWebElement WebButton_LoadBAM { get; set; }

        public string Time_ProductDevelopment_Sendports()
        {
            string srtext = "";
            try
            {
                srtext = driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.HMOrder.SendPort\")]/following-sibling::td[1])[1]")).Text.Trim();
            }
            catch
            {
                if ((srtext == "") || (srtext.Equals(null)))
                {
                    srtext = driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.RM.Primary.SendPort\")]/following-sibling::td[1])[1]")).Text.Trim();
                }
            }



            return srtext;
        }
        public string Verify_ProductDevelopment_Sendports()
        {
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.HMOrder.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort", "ProductDevelopment.HMOrder.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort", "ProductDevelopment.HMOrder.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }

               // System.Diagnostics.Debug.WriteLine(driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.HMOrder.SendPort\")]")).GetCssValue("color"));

                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.HMOrder.SendPort\")]")).GetCssValue("Color")== "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort Color", "ProductDevelopment.HMOrder.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort Color", "ProductDevelopment.HMOrder.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort", "exception raised"+e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.RM.Primary.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort", "ProductDevelopment.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort", "ProductDevelopment.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.RM.Primary.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort Color", "ProductDevelopment.RM.Primary.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort Color", "ProductDevelopment.RM.Primary.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.Bistro.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort", "ProductDevelopment.Bistro.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort", "ProductDevelopment.Bistro.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.Bistro.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort Color", "ProductDevelopment.Bistro.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort Color", "ProductDevelopment.Bistro.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.FashionSample.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort", "ProductDevelopment.FashionSample.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort", "ProductDevelopment.FashionSample.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.FashionSample.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort Color", "ProductDevelopment.FashionSample.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort Color", "ProductDevelopment.FashionSample.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.Salsa.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort", "ProductDevelopment.Salsa.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort", "ProductDevelopment.Salsa.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.Salsa.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort Color", "ProductDevelopment.Salsa.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort Color", "ProductDevelopment.Salsa.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.SPIN.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort", "ProductDevelopment.SPIN.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort", "ProductDevelopment.SPIN.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.SPIN.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort Color", "ProductDevelopment.SPIN.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort Color", "ProductDevelopment.SPIN.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.PLES.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.PLES.SendPort", "ProductDevelopment.PLES.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.PLES.SendPort", "ProductDevelopment.PLES.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.PLES.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.PLES.SendPort Color", "ProductDevelopment.PLES.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.PLES.SendPort Color", "ProductDevelopment.PLES.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.PLES.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            Thread.Sleep(2000);
            string srtext = "";
            try
            {
                srtext = driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.HMOrder.SendPort\")]/following-sibling::td[1])[1]")).Text.Trim();
            }
            catch
            {
                if ((srtext == "")||(srtext.Equals(null)))
                {
                    srtext = driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.RM.Primary.SendPort\")]/following-sibling::td[1])[1]")).Text.Trim();
                }
            }
             
           

            return srtext;
        }


        public Boolean Castor_ExportElementExistinTimeRange(string sexport, string timestamp)
        {
            DateTime ddate = DateTime.Parse(timestamp);

           
            DateTime ddate1 = ddate.Add(TimeSpan.FromMinutes(-1));
            DateTime ddate2 = ddate.Add(TimeSpan.FromMinutes(1));
            DateTime ddate3 = ddate.Add(TimeSpan.FromMinutes(-2));
            DateTime ddate4 = ddate.Add(TimeSpan.FromMinutes(2));
            DateTime ddate5 = ddate.Add(TimeSpan.FromMinutes(3));

            Boolean brecordfound = false;
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr[contains(.,\"" + ddate.ToString().Substring(0, 16)+"\")]/td[contains(.,\""+sexport+"\")]")).Displayed)
                {
                    brecordfound = true;
                    return true;
                }
            }
            catch
            {

            }
          
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr[contains(.,\"" + ddate2.ToString().Substring(0, 16) + "\")]/td[contains(.,\"" + sexport + "\")]")).Displayed)
                {
                    brecordfound = true;
                    return true;
                }
            }
            catch
            {

            }
           
           

            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr[contains(.,\"" + ddate4.ToString().Substring(0, 16) + "\")]/td[contains(.,\"" + sexport + "\")]")).Displayed)
                {
                    brecordfound = true;
                    return true;
                }
            }
            catch
            {

            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr[contains(.,\"" + ddate1.ToString().Substring(0, 16) + "\")]/td[contains(.,\"" + sexport + "\")]")).Displayed)
                {
                    brecordfound = true;
                    return true;
                }
            }
            catch
            {

            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr[contains(.,\"" + ddate3.ToString().Substring(0, 16) + "\")]/td[contains(.,\"" + sexport + "\")]")).Displayed)
                {
                    brecordfound = true;
                    return true;
                }
            }
            catch
            {

            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr[contains(.,\"" + ddate5.ToString().Substring(0, 16) + "\")]/td[contains(.,\"" + sexport + "\")]")).Displayed)
                {
                    brecordfound = true;
                    return true;
                }
            }
            catch
            {

            }
            return false;



        }
        public void Verify_ProductDevelopment_SendportswithTimestamp(string timestamp)
        {
           


            try
            {
                if (Castor_ExportElementExistinTimeRange("AssignedProductDevelopment.RM.Primary.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.RM.Primary.SendPort", "AssignedProductDevelopment.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.RM.Primary.SendPort", "AssignedProductDevelopment.RM.Primary.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"AssignedProductDevelopment.RM.Primary.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.RM.Primary.SendPort Color", "AssignedProductDevelopment.RM.Primary.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.RM.Primary.SendPort Color", "AssignedProductDevelopment.RM.Primary.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("AssignedProductDevelopment.RM.Primary.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (Castor_ExportElementExistinTimeRange("AssignedProductDevelopment.GARP.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.GARP.SendPort", "AssignedProductDevelopment.GARP.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.GARP.SendPort", "AssignedProductDevelopment.GARP.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"AssignedProductDevelopment.GARP.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.GARP.SendPort Color", "AssignedProductDevelopment.GARP.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.GARP.SendPort Color", "AssignedProductDevelopment.GARP.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("AssignedProductDevelopment.GARP.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (Castor_ExportElementExistinTimeRange("ProductDevelopment.FashionSample.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort", "ProductDevelopment.FashionSample.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort", "ProductDevelopment.FashionSample.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.FashionSample.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort Color", "ProductDevelopment.FashionSample.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort Color", "ProductDevelopment.FashionSample.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (Castor_ExportElementExistinTimeRange("AssignedProductDevelopment.Salsa.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.Salsa.SendPort", "AssignedProductDevelopment.Salsa.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.Salsa.SendPort", "AssignedProductDevelopment.Salsa.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"AssignedProductDevelopment.Salsa.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.Salsa.SendPort Color", "AssignedProductDevelopment.Salsa.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.Salsa.SendPort Color", "AssignedProductDevelopment.Salsa.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("AssignedProductDevelopment.Salsa.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (Castor_ExportElementExistinTimeRange("AssignedProductDevelopment.SPIN.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.SPIN.SendPort", "AssignedProductDevelopment.SPIN.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.SPIN.SendPort", "AssignedProductDevelopment.SPIN.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"AssignedProductDevelopment.SPIN.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.SPIN.SendPort Color", "AssignedProductDevelopment.SPIN.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("AssignedProductDevelopment.SPIN.SendPort Color", "AssignedProductDevelopment.SPIN.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("AssignedProductDevelopment.SPIN.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

          

            
        }


        public void Verify_DevelopmentOptionGroup_SendportswithTimestamp(string timestamp)
        {



            try
            {
                if (Castor_ExportElementExistinTimeRange("DevelopmentOptionGroup.Bistro.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent("DevelopmentOptionGroup.Bistro.SendPort", "DevelopmentOptionGroup.Bistro.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.Bistro.SendPort", " DevelopmentOptionGroup.Bistro.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"DevelopmentOptionGroup.Bistro.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.Bistro.SendPort Color", " DevelopmentOptionGroup.Bistro.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.Bistro.SendPort Color", " DevelopmentOptionGroup.Bistro.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("DevelopmentOptionGroup.Bistro.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (Castor_ExportElementExistinTimeRange("DevelopmentOptionGroup.RM.Primary.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.RM.Primary.SendPort", " DevelopmentOptionGroup.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.RM.Primary.SendPort", " DevelopmentOptionGroup.RM.Primary.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"DevelopmentOptionGroup.RM.Primary.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.RM.Primary.SendPort Color", " DevelopmentOptionGroup.RM.Primary.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.RM.Primary.SendPort Color", " DevelopmentOptionGroup.RM.Primary.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent(" DevelopmentOptionGroup.RM.Primary.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (Castor_ExportElementExistinTimeRange("DevelopmentOptionGroup.Disco_2_0.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.Disco_2_0.SendPort", " DevelopmentOptionGroup.Disco_2_0.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.Disco_2_0.SendPort", " DevelopmentOptionGroup.Disco_2_0.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"DevelopmentOptionGroup.Disco_2_0.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.Disco_2_0.SendPort Color", " DevelopmentOptionGroup.Disco_2_0.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent(" DevelopmentOptionGroup.Disco_2_0.SendPort Color", " DevelopmentOptionGroup.Disco_2_0.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("DevelopmentOptionGroup.Disco_2_0.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (Castor_ExportElementExistinTimeRange("DevelopmentOptionGroup.Salsa.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent("DevelopmentOptionGroup.Salsa.SendPort", "DevelopmentOptionGroup.Salsa.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("DevelopmentOptionGroup.Salsa.SendPort", "DevelopmentOptionGroup.Salsa.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"DevelopmentOptionGroup.Salsa.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("DevelopmentOptionGroup.Salsa.SendPort Color", "DevelopmentOptionGroup.Salsa.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("DevelopmentOptionGroup.Salsa.SendPort Color", "DevelopmentOptionGroup.Salsa.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("DevelopmentOptionGroup.Salsa.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (Castor_ExportElementExistinTimeRange("DevelopmentOptionGroup.SPIN.SendPort", timestamp) == true)
                {
                    Reporter.ReportEvent("DevelopmentOptionGroup.SPIN.SendPort", "DevelopmentOptionGroup.SPIN.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("DevelopmentOptionGroup.SPIN.SendPort", "DevelopmentOptionGroup.SPIN.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"DevelopmentOptionGroup.SPIN.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("DevelopmentOptionGroup.SPIN.SendPort Color", "DevelopmentOptionGroup.SPIN.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("DevelopmentOptionGroup.SPIN.SendPort Color", "DevelopmentOptionGroup.SPIN.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("DevelopmentOptionGroup.SPIN.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            




        }


        public void Verify_ProductDevelopment_Sendports_3()
        {
            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.HMOrder.SendPort\")])[3]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort", "ProductDevelopment.HMOrder.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort", "ProductDevelopment.HMOrder.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }

                // System.Diagnostics.Debug.WriteLine(driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.HMOrder.SendPort\")]")).GetCssValue("color"));

                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.HMOrder.SendPort\")])[3]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort Color", "ProductDevelopment.HMOrder.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort Color", "ProductDevelopment.HMOrder.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.HMOrder.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.RM.Primary.SendPort\")])[3]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort", "ProductDevelopment.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort", "ProductDevelopment.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.RM.Primary.SendPort\")])[3]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort Color", "ProductDevelopment.RM.Primary.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort Color", "ProductDevelopment.RM.Primary.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.RM.Primary.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.Bistro.SendPort\")])[3]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort", "ProductDevelopment.Bistro.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort", "ProductDevelopment.Bistro.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.Bistro.SendPort\")])[3]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort Color", "ProductDevelopment.Bistro.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort Color", "ProductDevelopment.Bistro.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.Bistro.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.FashionSample.SendPort\")])[3]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort", "ProductDevelopment.FashionSample.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort", "ProductDevelopment.FashionSample.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.FashionSample.SendPort\")])[3]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort Color", "ProductDevelopment.FashionSample.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort Color", "ProductDevelopment.FashionSample.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.FashionSample.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.Salsa.SendPort\")])[3]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort", "ProductDevelopment.Salsa.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort", "ProductDevelopment.Salsa.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.Salsa.SendPort\")])[3]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort Color", "ProductDevelopment.Salsa.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort Color", "ProductDevelopment.Salsa.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.Salsa.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.SPIN.SendPort\")])[3]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort", "ProductDevelopment.SPIN.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort", "ProductDevelopment.SPIN.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.SPIN.SendPort\")])[3]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort Color", "ProductDevelopment.SPIN.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort Color", "ProductDevelopment.SPIN.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.SPIN.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }



            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.PLES.SendPort\")])[3]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.PLES.SendPort", "ProductDevelopment.PLES.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.PLES.SendPort", "ProductDevelopment.PLES.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.PLES.SendPort\")])[3]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.PLES.SendPort Color", "ProductDevelopment.PLES.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.PLES.SendPort Color", "ProductDevelopment.PLES.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.PLES.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.SIM.SendPort\")])[3]")).Displayed)
                {
                    Reporter.ReportEvent("ProductDevelopment.SIM.SendPort", "ProductDevelopment.SIM.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.SIM.SendPort", "ProductDevelopment.SIM.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"ProductDevelopment.SIM.SendPort\")])[3]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("ProductDevelopment.SIM.SendPort Color", "ProductDevelopment.SIM.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("ProductDevelopment.SIM.SendPort Color", "ProductDevelopment.SIM.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("ProductDevelopment.SIM.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
        }

        public void HMOrder_BOPOInternalOrderExports_Verification()
        {
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//div/table/tbody/tr/td[contains(.,\"BOPOInternalOrder.HMOrder.ReceivePort\")]")).Displayed)
                {
                    Reporter.ReportEvent("BOPOInternalOrder.HMOrder.ReceivePort", "BOPOInternalOrder.HMOrder.ReceivePort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.HMOrder.ReceivePort", "BOPOInternalOrder.HMOrder.ReceivePort is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch
            {

            }
           


            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.Castor.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Castor.SendPort", "BOPOInternalOrder.Castor.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Castor.SendPort", "BOPOInternalOrder.Castor.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.Castor.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Castor.SendPort Color", "BOPOInternalOrder.Castor.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Castor.SendPort Color", "BOPOInternalOrder.Castor.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("BOPOInternalOrder.Castor.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.Bistro.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Bistro.SendPort", "BOPOInternalOrder.Bistro.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Bistro.SendPort", "BOPOInternalOrder.Bistro.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.Bistro.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Bistro.SendPort Color", "BOPOInternalOrder.Bistro.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Bistro.SendPort Color", "BOPOInternalOrder.Bistro.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("BOPOInternalOrder.Bistro.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            //try
            //{
            //    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.OFS.SendPort\")]")).Displayed)
            //    {
            //        Reporter.ReportEvent("BOPOInternalOrder.OFS.SendPort", "BOPOInternalOrder.OFS.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
            //    }
            //    else
            //    {
            //        Reporter.ReportEvent("BOPOInternalOrder.OFS.SendPort", "BOPOInternalOrder.OFS.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
            //    }
            //    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.OFS.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
            //    {
            //        Reporter.ReportEvent("BOPOInternalOrder.OFS.SendPort Color", "BOPOInternalOrder.OFS.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
            //    }
            //    else
            //    {
            //        Reporter.ReportEvent("BOPOInternalOrder.OFS.SendPort Color", "BOPOInternalOrder.OFS.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
            //    }
            //}
            //catch (Exception e)
            //{
            //    Reporter.ReportEvent("BOPOInternalOrder.OFS.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            //}
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.Disco.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Disco.SendPort", "BOPOInternalOrder.Disco.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Disco.SendPort", "BOPOInternalOrder.Disco.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.Disco.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Disco.SendPort Color", "BOPOInternalOrder.Disco.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Disco.SendPort Color", "BOPOInternalOrder.Disco.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("BOPOInternalOrder.Disco.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.PLES.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("BOPOInternalOrder.PLES.SendPort", "BOPOInternalOrder.PLES.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.PLES.SendPort", "BOPOInternalOrder.PLES.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.PLES.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("BOPOInternalOrder.PLES.SendPort Color", "BOPOInternalOrder.PLES.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.PLES.SendPort Color", "BOPOInternalOrder.PLES.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("BOPOInternalOrder.PLES.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.Disco_2_0.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Disco_2_0.SendPort", "BOPOInternalOrder.Disco_2_0.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Disco_2_0.SendPort", "BOPOInternalOrder.Disco_2_0.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.Disco_2_0.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Disco_2_0.SendPort Color", "BOPOInternalOrder.Disco_2_0.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.Disco_2_0.SendPort Color", "BOPOInternalOrder.Disco_2_0.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("BOPOInternalOrder.Disco_2_0.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            //try
            //{
            //    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.GARP.SendPort\")]")).Displayed)
            //    {
            //        Reporter.ReportEvent("BOPOInternalOrder.GARP.SendPort", "BOPOInternalOrder.GARP.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
            //    }
            //    else
            //    {
            //        Reporter.ReportEvent("BOPOInternalOrder.GARP.SendPort", "BOPOInternalOrder.GARP.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
            //    }
            //    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.GARP.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
            //    {
            //        Reporter.ReportEvent("BOPOInternalOrder.GARP.SendPort Color", "BOPOInternalOrder.GARP.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
            //    }
            //    else
            //    {
            //        Reporter.ReportEvent("BOPOInternalOrder.GARP.SendPort Color", "BOPOInternalOrder.GARP.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
            //    }
            //}
            //catch (Exception e)
            //{
            //    Reporter.ReportEvent("BOPOInternalOrder.GARP.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            //}
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.SALSA.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("BOPOInternalOrder.SALSA.SendPort", "BOPOInternalOrder.SALSA.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.SALSA.SendPort", "BOPOInternalOrder.SALSA.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.SALSA.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("BOPOInternalOrder.SALSA.SendPort Color", "BOPOInternalOrder.SALSA.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.SALSA.SendPort Color", "BOPOInternalOrder.SALSA.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("BOPOInternalOrder.SALSA.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.BoPoSearch.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("BOPOInternalOrder.BoPoSearch.SendPort", "BOPOInternalOrder.BoPoSearch.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.BoPoSearch.SendPort", "BOPOInternalOrder.BoPoSearch.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"BOPOInternalOrder.BoPoSearch.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("BOPOInternalOrder.BoPoSearch.SendPort Color", "BOPOInternalOrder.BoPoSearch.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("BOPOInternalOrder.BoPoSearch.SendPort Color", "BOPOInternalOrder.BoPoSearch.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("BOPOInternalOrder.BoPoSearch.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
        }
        public void HMOrder_GenericOrderExports_Verification()
        {
            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//div/table/tbody/tr/td[contains(.,\"GenericOrder.HMOrder.ReceivePort\")])[2]")).Displayed)
                {
                    Reporter.ReportEvent("GenericOrder.HMOrder.ReceivePort", "GenericOrder.HMOrder.ReceivePort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("GenericOrder.HMOrder.ReceivePort", "GenericOrder.HMOrder.ReceivePort is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }

                //try
                //{
                //    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.BOPOSearch.SendPort\")]")).Displayed)
                //    {
                //        Reporter.ReportEvent("GenericOrder.BOPOSearch.SendPort", "GenericOrder.BOPOSearch.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                //    }
                //    else
                //    {
                //        Reporter.ReportEvent("GenericOrder.BOPOSearch.SendPort", "GenericOrder.BOPOSearch.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                //    }
                //    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.BOPOSearch.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                //    {
                //        Reporter.ReportEvent("GenericOrder.BOPOSearch.SendPort Color", "GenericOrder.BOPOSearch.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                //    }
                //    else
                //    {
                //        Reporter.ReportEvent("GenericOrder.BOPOSearch.SendPort Color", "GenericOrder.BOPOSearch.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                //    }
                //}
                //catch (Exception e)
                //{
                //    Reporter.ReportEvent("GenericOrder.BOPOSearch.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                //}
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.ITS.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.ITS.SendPort", "GenericOrder.ITS.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.ITS.SendPort", "GenericOrder.ITS.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.ITS.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.ITS.SendPort Color", "GenericOrder.ITS.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.ITS.SendPort Color", "GenericOrder.ITS.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.ITS.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.PRA.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.PRA.SendPort", "GenericOrder.PRA.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.PRA.SendPort", "GenericOrder.PRA.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.PRA.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.PRA.SendPort Color", "GenericOrder.PRA.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.PRA.SendPort Color", "GenericOrder.PRA.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.PRA.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.RM.Primary.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.RM.Primary.SendPort", "GenericOrder.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.RM.Primary.SendPort", "GenericOrder.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.RM.Primary.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.RM.Primary.SendPort Color", "GenericOrder.RM.Primary.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.RM.Primary.SendPort Color", "GenericOrder.RM.Primary.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.RM.Primary.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.SAPECC.OrderPool.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.SAPECC.OrderPool.SendPort", "GenericOrder.SAPECC.OrderPool.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.SAPECC.OrderPool.SendPort", "GenericOrder.SAPECC.OrderPool.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.SAPECC.OrderPool.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.SAPECC.OrderPool.SendPort Color", "GenericOrder.SAPECC.OrderPool.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.SAPECC.OrderPool.SendPort Color", "GenericOrder.SAPECC.OrderPool.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.SAPECC.OrderPool.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.SelfBilling.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.SelfBilling.SendPort", "GenericOrder.SelfBilling.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.SelfBilling.SendPort", "GenericOrder.SelfBilling.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.SelfBilling.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.SelfBilling.SendPort Color", "GenericOrder.SelfBilling.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.SelfBilling.SendPort Color", "GenericOrder.SelfBilling.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.SelfBilling.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.GOEP.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.GOEP.SendPort", "GenericOrder.GOEP.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.GOEP.SendPort", "GenericOrder.GOEP.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.GOEP.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.GOEP.SendPort Color", "GenericOrder.GOEP.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.GOEP.SendPort Color", "GenericOrder.GOEP.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.GOEP.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.SPC.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.SPC.SendPort", "GenericOrder.SPC.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.SPC.SendPort", "GenericOrder.SPC.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.SPC.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.SPC.SendPort Color", "GenericOrder.SPC.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.SPC.SendPort Color", "GenericOrder.SPC.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.SPC.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.CDW.Prod.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.CDW.Prod.SendPort", "GenericOrder.CDW.Prod.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.CDW.Prod.SendPort", "GenericOrder.CDW.Prod.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.CDW.Prod.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.CDW.Prod.SendPort Color", "GenericOrder.CDW.Prod.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.CDW.Prod.SendPort Color", "GenericOrder.CDW.Prod.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.CDW.Prod.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.CDW.PreProd.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.CDW.PreProd.SendPort", "GenericOrder.CDW.PreProd.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.CDW.PreProd.SendPort", "GenericOrder.CDW.PreProd.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.CDW.PreProd.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.CDW.PreProd.SendPort Color", "GenericOrder.CDW.PreProd.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.CDW.PreProd.SendPort Color", "GenericOrder.CDW.PreProd.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.CDW.PreProd.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.CCP.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.CCP.SendPort", "GenericOrder.CCP.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.CCP.SendPort", "GenericOrder.CCP.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.CCP.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.CCP.SendPort Color", "GenericOrder.CCP.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.CCP.SendPort Color", "GenericOrder.CCP.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.CCP.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.TaxService.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.TaxService.SendPort", "GenericOrder.TaxService.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.TaxService.SendPort", "GenericOrder.TaxService.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.TaxService.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.TaxService.SendPort Color", "GenericOrder.TaxService.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.TaxService.SendPort Color", "GenericOrder.TaxService.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.TaxService.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
                try
                {
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.TAGS.SendPort\")]")).Displayed)
                    {
                        Reporter.ReportEvent("GenericOrder.TAGS.SendPort", "GenericOrder.TAGS.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.TAGS.SendPort", "GenericOrder.TAGS.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                    if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"GenericOrder.TAGS.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                    {
                        Reporter.ReportEvent("GenericOrder.TAGS.SendPort Color", "GenericOrder.TAGS.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                    }
                    else
                    {
                        Reporter.ReportEvent("GenericOrder.TAGS.SendPort Color", "GenericOrder.TAGS.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                    }
                }
                catch (Exception e)
                {
                    Reporter.ReportEvent("GenericOrder.TAGS.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("GenericOrder.HMOrder.ReceivePort 2", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
        }
        public void HMOrder_OrderedProductSpecification_Verification()
        {
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//div/table/tbody/tr/td[contains(.,\"OrderedProductSpecification.Castor.ReceivePort\")]")).Displayed)
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Castor.ReceivePort", "OrderedProductSpecification.Castor.ReceivePort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Castor.ReceivePort", "OrderedProductSpecification.Castor.ReceivePort is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch
            {

            }



            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.Bistro.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Bistro.SendPort", "OrderedProductSpecification.Bistro.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Bistro.SendPort", "OrderedProductSpecification.Bistro.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.Bistro.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Bistro.SendPort Color", "OrderedProductSpecification.Bistro.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Bistro.SendPort Color", "OrderedProductSpecification.Bistro.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("OrderedProductSpecification.Bistro.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.HMOrder.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("OrderedProductSpecification.HMOrder.SendPort", "OrderedProductSpecification.HMOrder.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.HMOrder.SendPort", "OrderedProductSpecification.HMOrder.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.HMOrder.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("OrderedProductSpecification.HMOrder.SendPort Color", "OrderedProductSpecification.HMOrder.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.HMOrder.SendPort Color", "OrderedProductSpecification.HMOrder.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("OrderedProductSpecification.HMOrder.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.Disco.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Disco.SendPort", "OrderedProductSpecification.Disco.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Disco.SendPort", "OrderedProductSpecification.Disco.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.Disco.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Disco.SendPort Color", "OrderedProductSpecification.Disco.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Disco.SendPort Color", "OrderedProductSpecification.Disco.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("OrderedProductSpecification.Disco.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.TAGS.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("OrderedProductSpecification.TAGS.SendPort", "OrderedProductSpecification.TAGS.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.TAGS.SendPort", "OrderedProductSpecification.TAGS.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.TAGS.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("OrderedProductSpecification.TAGS.SendPort Color", "OrderedProductSpecification.TAGS.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.TAGS.SendPort Color", "OrderedProductSpecification.TAGS.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("OrderedProductSpecification.TAGS.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.RM.Primary.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("OrderedProductSpecification.RM.Primary.SendPort", "OrderedProductSpecification.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.RM.Primary.SendPort", "OrderedProductSpecification.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.RM.Primary.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("OrderedProductSpecification.RM.Primary.SendPort Color", "OrderedProductSpecification.RM.Primary.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.RM.Primary.SendPort Color", "OrderedProductSpecification.RM.Primary.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("OrderedProductSpecification.RM.Primary.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.SALSA.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("OrderedProductSpecification.SALSA.SendPort", "OrderedProductSpecification.SALSA.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.SALSA.SendPort", "OrderedProductSpecification.SALSA.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.SALSA.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("OrderedProductSpecification.SALSA.SendPort Color", "OrderedProductSpecification.SALSA.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.SALSA.SendPort Color", "OrderedProductSpecification.SALSA.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("OrderedProductSpecification.SALSA.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.Disco_2_0.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Disco_2_0.SendPort", "OrderedProductSpecification.Disco_2_0.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Disco_2_0.SendPort", "OrderedProductSpecification.Disco_2_0.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"OrderedProductSpecification.Disco_2_0.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Disco_2_0.SendPort Color", "OrderedProductSpecification.Disco_2_0.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("OrderedProductSpecification.Disco_2_0.SendPort Color", "OrderedProductSpecification.Disco_2_0.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("OrderedProductSpecification.Disco_2_0.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
           
        }

        public void CapacityBooking_Castor_Verification()
        {
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//div/table/tbody/tr/td[contains(.,\"CapacityBooking.Bistro.ReceivePort\")]")).Displayed)
                {
                    Reporter.ReportEvent("CapacityBooking.Bistro.ReceivePort", "CapacityBooking.Bistro.ReceivePort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CapacityBooking.Bistro.ReceivePort", "CapacityBooking.Bistro.ReceivePort is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch
            {

            }

            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.Castor.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("CapacityBooking.Castor.SendPort", "CapacityBooking.Castor.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CapacityBooking.Castor.SendPort", "CapacityBooking.Castor.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.Castor.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("CapacityBooking.Castor.SendPort Color", "CapacityBooking.Castor.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CapacityBooking.Castor.SendPort Color", "CapacityBooking.Castor.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("CapacityBooking.Castor.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.SALSA.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("CapacityBooking.SALSA.SendPort", "CapacityBooking.SALSA.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CapacityBooking.SALSA.SendPort", "CapacityBooking.SALSA.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.SALSA.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("CapacityBooking.SALSA.SendPort Color", "CapacityBooking.SALSA.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CapacityBooking.SALSA.SendPort Color", "CapacityBooking.SALSA.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("CapacityBooking.SALSA.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.BoPoSearch.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("CapacityBooking.BoPoSearch.SendPort", "CapacityBooking.BoPoSearch.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CapacityBooking.BoPoSearch.SendPort", "CapacityBooking.BoPoSearch.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.BoPoSearch.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("CapacityBooking.BoPoSearch.SendPort Color", "CapacityBooking.BoPoSearch.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CapacityBooking.BoPoSearch.SendPort Color", "CapacityBooking.BoPoSearch.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("CapacityBooking.BoPoSearch.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.RM.Primary.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent(" CapacityBooking.RM.Primary.SendPort", " CapacityBooking.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent(" CapacityBooking.RM.Primary.SendPort", " CapacityBooking.RM.Primary.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.RM.Primary.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent(" CapacityBooking.RM.Primary.SendPort Color", " CapacityBooking.RM.Primary.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent(" CapacityBooking.RM.Primary.SendPort Color", " CapacityBooking.RM.Primary.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("CapacityBooking.RM.Primary.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);

            }
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.RM.StandBy.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("CapacityBooking.RM.StandBy.SendPort", "CapacityBooking.RM.StandBy.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CapacityBooking.RM.StandBy.SendPort", "CapacityBooking.RM.StandBy.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr/td[contains(.,\"CapacityBooking.RM.StandBy.SendPort\")]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("CapacityBooking.RM.StandBy.SendPort Color", "CapacityBooking.RM.StandBy.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CapacityBooking.RM.StandBy.SendPort Color", "CapacityBooking.RM.StandBy.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("CapacityBooking.RM.StandBy.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }

        }
     
        public void CountryCertificate_Castor_Verification()
        {
            try
            {
                if (driver.FindElement(By.XPath("(.//scrollable-table//div/table/tbody/tr/td[contains(.,\"CountryCertificate.Castor.ReceivePort\")])[1]")).Displayed)
                {
                    Reporter.ReportEvent("CountryCertificate.Castor.ReceivePort", "CountryCertificate.Castor.ReceivePort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CountryCertificate.Castor.ReceivePort", "CountryCertificate.Castor.ReceivePort is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch
            {

            }

            try
            {
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"CountryCertificate.Tags.SendPort\")])[1]")).Displayed)
                {
                    Reporter.ReportEvent("CountryCertificate.Tags.SendPort", "CountryCertificate.Tags.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CountryCertificate.Tags.SendPort", "CountryCertificate.Tags.SendPort displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("(//scrollable-table//table/tbody/tr/td[contains(.,\"CountryCertificate.Tags.SendPort\")])[1]")).GetCssValue("Color") == "rgb(105, 105, 105)")
                {
                    Reporter.ReportEvent("CountryCertificate.Tags.SendPort Color", "CountryCertificate.Tags.SendPort Grey color displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("CountryCertificate.Tags.SendPort Color", "CountryCertificate.Tags.SendPort Grey color is not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("CountryCertificate.Tags.SendPort", "exception raised" + e.Message, HP.LFT.Report.Status.Failed);
            }
        }

        ///==================================================================
        ///preplan code merge





        public IWebElement get_ApplicationDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='selectIO']")));
        }

        public IWebElement get_LoadBAMButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//button[text()='Load BAM']")));
        }

        public IWebElement get_AscendingResultButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath(" //button[@title='Click for Ascending Order']")));
        }

        public List<IWebElement> get_ResultFromICC()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//div[@class='scrollArea']/table//tbody//tr")))
                .ToList();
        }

        public List<IWebElement> get_SendPort()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//div[@class='scrollArea']/table//tbody//tr//td[contains(text(),'SendPort')]")))
                .ToList();
        }


        public IWebElement get_BoardCardTxtBox()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(text(),'BoardcardNumber')]/following-sibling::td/input")));
        }

        public IWebElement get_ProductIdTxtBox()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(text(),'ProductId')]/following-sibling::td/input")));
        }


        public IWebElement get_EnvironmentDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='selectEnv']")));
        }

        public void LaunchICCWindow()
        {
            LaunchTabAndNavigate();
            currentWindow = get_CurrentWindowHandle();
            SwitchToNewWindow();
            string url = ConfigUtils.Read("URL_ICC");
            driver.Navigate().GoToUrl(url);
            SelectEnvironment();
        }

        private void SelectEnvironment()
        {
            if (get_EnvironmentDropDown().Displayed)
            {
                if (ConfigUtils.Read("ENV") == "AT")
                {
                    sGMethods.SelectDropDownByValue(get_EnvironmentDropDown(), ConfigUtils.Read("ICCATEnv"));
                }
                else if (ConfigUtils.Read("ENV") == "SIT")
                {
                    sGMethods.SelectDropDownByValue(get_EnvironmentDropDown(), ConfigUtils.Read("ICCSITENV"));
                }
                else if (ConfigUtils.Read("ENV") == "DIT")
                {
                    sGMethods.SelectDropDownByValue(get_EnvironmentDropDown(), ConfigUtils.Read("ICCDITEnv"));
                } 
            }
        }

        public void SelectValueFromApplicationDropDown(string applicationValue)
        {
            if (get_ApplicationDropDown().Displayed)
            {
                sGMethods.SelectDropDownByValue(get_ApplicationDropDown(), applicationValue);
            }
            get_LoadBAMButton().Click();

            try
            {
                Assert.IsTrue(get_AscendingResultButton().Displayed);
            }
            catch
            {
                get_LoadBAMButton().Click();
                Assert.IsTrue(get_AscendingResultButton().Displayed);
            }
        }

        public void SelectValueFromApplicationDropDown(string productNumber, string applicationValue)
        {
            if (get_ApplicationDropDown().Displayed)
            {
                sGMethods.SelectDropDownByValue(get_ApplicationDropDown(), applicationValue);
            }
            get_ProductIdTxtBox().Clear();
            get_ProductIdTxtBox().SendKeys(productNumber);
            get_LoadBAMButton().Click();
            Assert.IsTrue(get_AscendingResultButton().Displayed);
        }

        public void SelectValueFromApplicationDropDownForBoardCard(string productNumber, string applicationValue)
        {
            if (get_ApplicationDropDown().Displayed)
            {
                sGMethods.SelectDropDownByValue(get_ApplicationDropDown(), applicationValue);
            }
            get_BoardCardTxtBox().Clear();
            get_BoardCardTxtBox().SendKeys(productNumber);
            get_LoadBAMButton().Click();
            Assert.IsTrue(get_AscendingResultButton().Displayed);
        }

        public void VerifySearchResult(string applicationValue)
        {
            Assert.IsTrue(get_ResultFromICC().Count > 0);
            string[] stringList = applicationValue.Split('_');
            Assert.IsTrue(get_ResultFromICC()[0].Text.Contains(stringList[0]));
        }

        public bool VerifyPortInSearchResult(string portName)
        {
            foreach (var port in get_SendPort())
            {
                if (port.Text.Contains(portName) && !(port.GetAttribute("class").Contains("sndError")))
                {
                    return true;
                }
            }
            return false;
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

        public void Mchart_ExportElementExistinTimeRange()
        {
            DateTime ddate = DateTime.Today;
            

         
            try
            {
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr[contains(.,\"" + ddate.ToShortDateString()+ "\")]/td[contains(.,\"MeasurementChart.OFS.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("MeasurementChart.OFS.SendPort", "MeasurementChart.OFS.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("MeasurementChart.OFS.SendPort", "MeasurementChart.OFS.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
                if (driver.FindElement(By.XPath("//scrollable-table//table/tbody/tr[contains(.,\"" + ddate.ToShortDateString() + "\")]/td[contains(.,\"MeasurementChart.ShowRoom.SendPort\")]")).Displayed)
                {
                    Reporter.ReportEvent("MeasurementChart.ShowRoom.SendPort", "MeasurementChart.ShowRoom.SendPort displayed in ICC", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("MeasurementChart.ShowRoom.SendPort", "MeasurementChart.ShowRoom.SendPort not displayed in ICC", HP.LFT.Report.Status.Failed);
                }
            }
            catch
            {
                Reporter.ReportEvent("Mchart_ExportElementExistinTimeRange script failed", "Mchart_ExportElementExistinTimeRange script failed", HP.LFT.Report.Status.Failed);
            }
        }

        public bool VerifyCITStatusExport(string orderNumber)
        {
            string sendPortName = "CITStatus.Castor.SendPort";
            return VerifyOFUExports("CITStatus_SCSC", orderNumber, sendPortName);
            
        }

        public bool VerifyDetailedCompositionExport(string orderNumber)
        {
            string sendPortName = "DetailedComposition.Tags.SendPort";
            return VerifyOFUExports("DetailedComposition_SCSC", orderNumber, sendPortName);
        }

        private bool VerifyOFUExports(string application, string orderNumber, string sendPortName)
        {
            //(.//scrollable-table/div/div[2]/table/tbody/tr/td[contains(., "CITStatus.Castor.SendPort")])[1]

            /*
             * Open ICC - Done
             * Select Enivronment - Done
             * Choose Application - Done
             * Input Order Number - Done
             * Input Season - Done
             * Input SendPort Name - Done
             * Click Apply Filter/Sort Button - Done
             * Verify if the row is displayed - Done
             * */
            
            sGMethods.SelectDropDownByValue(Dropdown_Application, "Application DropDown", application);
            sGMethods.WebEdit_SetValue(driver.FindElement(By.XPath("//table[@class=\"searchFieldsTable\"]//td[contains(.,\"OrderNumber\")]/following-sibling::td/input")), "OrderNumber", orderNumber);
            //sGMethods.WebEdit_SetValue(driver.FindElement(By.XPath("//table[@class=\"searchFieldsTable\"]//td[contains(.,\"SeasonNumber\")]/following-sibling::td/input")), "SeasonNumber", season);
            sGMethods.WebEdit_SetValue(driver.FindElement(By.XPath(".//*[@id='sndPortName']")), "sndPortName", sendPortName);

            sGMethods.WebButton_Click(WebButton_LoadBAM);
            Thread.Sleep(1000);
            wait.Until(ExpectedConditions.ElementExists(By.XPath(".//*[@id='main']//fieldset/table/tbody//button[contains(.,\"Load BAM\")]")));

            try
            {
                if (driver.FindElement(By.XPath("(.//scrollable-table//div/table/tbody/tr/td[contains(.,\"" + sendPortName + "\")])[1]")).Displayed)
                {
                    Reporter.ReportEvent("VerifyOFUExports", sendPortName + " is displayed in ICC", Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("VerifyOFUExports", sendPortName + " is not displayed in ICC", Status.Failed);
                    return false;
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("Selenium_VerifyOFUExports Script failed", e.Message, Status.Failed);
            }
            return true;
        }
    }
}
