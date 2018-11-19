using BOPO.NUnit.ParallelTests.Helpers;
using BOPO.NUnit.ParallelTests.POM.TestRunup;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace BOPO.NUnit.ParallelTests.Tests.TestRunup
{
    public class TC04_ProductPlan_IEBrowsers : TestBase
    {

      IWebDriver driver;

        public Boolean TC04_ProductPlan_IEBrowser()
        {
            bool finalResult = false;
            Initial_Setup("TC04_ProductPlan_IEBrowser");
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "ProductPlan");
            try {

                POM_ProductPlanPage pPlanPage = new POM_ProductPlanPage(BrowserName, Driver, logStack);

               this.driver = Driver;

               
                string directive = "20 - H&M Man";
            string directiveOther = "18 - Ladies Trend";
            string season = "7-2018";
            string tabName = "5262 - Jacket";
            string changeseason = "6-2017";
            
            //try
            //{
            //driver.Manage().Window.Maximize();
            //driver.Navigate().GoToUrl("http://productplan.sit.hm.com");
         //   pPlanPage.waitForPageToLoad();
           // pPlanPage.waitForSpinnerToDisappear();
            Thread.Sleep(4000);
            if (!pPlanPage.seasonButtonStatus.Text.Contains(season))
            {
                pPlanPage.get_seasonSelectionDirectiveButton(season).Click();
            }
           ReportCustom.Report(logStack, driver, "Season selected : " + season, "Season selected : " + season);
           // ReportCustom.ReportEvent("Season selected", "Season selected : " + season, HP.LFT.Report.Status.Passed);
            pPlanPage.waitForPageToLoad();
            Assert.IsTrue(pPlanPage.get_ProductNameTableHeaderElement().Displayed);
            if (!pPlanPage.SectionDirectrivesDropDownElementSelectedElement.Text.Contains(directive))
            {
                    
                SelectElement ele = new SelectElement(pPlanPage.SectionDirectrivesDropDownElement);
                ele.SelectByText(directive);
            }
           ReportCustom.Report(logStack, driver, "Directive selected from dropdown : " + directive, "Directive selected from dropdown : " + directive);
         //   ReportCustom.ReportEvent("Directive selected from dropdown", "Directive selected from dropdown : " + directive,
           //     HP.LFT.Report.Status.Passed);
            if (pPlanPage.get_departmentTabElement(tabName).Displayed)
            {
                if (!pPlanPage.DepartmentTabActivElement.Text.Contains(tabName))
                {
                    pPlanPage.get_departmentTabElement(tabName).Click();
                    pPlanPage.waitForPageToLoad();
                    pPlanPage.waitForSpinnerToDisappear();
                }
            }
           ReportCustom.Report(logStack, driver, "Tab selected : " + tabName, "Tab selected : " + tabName);
           // ReportCustom.ReportEvent("Tab selected", "Tab selected : " + tabName, HP.LFT.Report.Status.Passed);
            Thread.Sleep(4000);
                // pPlanPage.get_TotalBuyStatusIconExpand().Click();
                try
                {
                    driver.FindElement(By.XPath("//button[text()='Full']")).Click();
                    pPlanPage.waitForPageToLoad();
                    pPlanPage.waitForSpinnerToDisappear();
                }
                catch
                {
                }

                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), pPlanPage.versionInfo.GetAttribute("textContent"));
                VersionConrol.addApplication("ProductPlan", pPlanPage.versionInfo.GetAttribute("textContent"));
                var str1 = pPlanPage.get_TotalBuyStatusIconExpand().GetAttribute("title");
                var str2= "Show Total Buying Status";
                if (str1==str2)
                {
                    pPlanPage.get_TotalBuyStatusIconExpand().Click();
                }
                String s=pPlanPage.get_TotalBuyStatusTableLastColumnPercentKeys()[1].Text;
            Assert.IsTrue(pPlanPage.get_TotalBuyStatusTableLastColumnPercentKeys().Count > 5);
            Assert.IsTrue(pPlanPage.get_ProductNameTableHeaderElement().Displayed);
            //pPlanPage.waitForSpinnerToDisappear();
            //if cont more than one scroll
            pPlanPage.get_CreateProductTextBox();
            string productName = "AutoTest" + (DateTime.Now.ToString("yyyyhhmmssffff"));
            pPlanPage.CreateProductNameTextBoxElement.SendKeys(productName);
            pPlanPage.CreateProductNameTextBoxElement.SendKeys(Keys.Enter);
           ReportCustom.Report(logStack, driver, "Product name enter in text box : " + productName, "Product name enter in text box : " + productName);
          //  ReportCustom.ReportEvent("Product name enter in text box", "Product name enter in text box : " + productName,
            //    HP.LFT.Report.Status.Passed);
            Thread.Sleep(2000);
            Assert.IsTrue(pPlanPage.get_ProductNameFromLeftColumn(productName).Displayed);
           ReportCustom.Report(logStack, driver, "Product Name display and created", "Product Name display and created");
          //  ReportCustom.ReportEvent("Product Name display and created ", "Product Name display and created",
            //    HP.LFT.Report.Status.Passed);
            Thread.Sleep(2000);
            pPlanPage.get_productExpandButton(productName).Click();
            pPlanPage.get_ArticleCodeArea(productName).Click();
            pPlanPage.get_ArticleCodeAreaTextBox().Clear();
            pPlanPage.get_ArticleCodeAreaTextBox().Click();
            pPlanPage.get_ArticleCodeAreaTextBox().SendKeys("09-101");
            driver.FindElement(By.XPath("//ul//*[contains(text(),'09-101')]")).Click();
            Thread.Sleep(2000);
            pPlanPage.rightClickOnProductName(productName);
            pPlanPage.get_Save().Click();
            pPlanPage.waitForSpinnerToDisappear();
            pPlanPage.waitForPageToLoad();
            pPlanPage.rightClickOnProductName(productName);
            pPlanPage.get_Delete().Click();
            pPlanPage.get_DeleteConfirmation().Click();
            pPlanPage.SaveButtonHeaderElement.Click();
            pPlanPage.waitForSpinnerToDisappear();
            pPlanPage.waitForPageToLoad();
            pPlanPage.get_UnSelectedTabs()[pPlanPage.get_UnSelectedTabs().Count - 1].Click();
            pPlanPage.waitForSpinnerToDisappear();
            pPlanPage.waitForPageToLoad();
         //   Thread.Sleep(4000);
            //pPlanPage.get_TotalBuyStatusIconExpand().Click();
            //pPlanPage.get_TotalBuyStatusTableLastColumnPercentKeys()[1].Click();
            //Assert.IsTrue(pPlanPage.get_TotalBuyStatusTableLastColumnPercentKeys().Count > 5);
            if (!pPlanPage.seasonButtonStatus.Text.Contains(season))
            {
                pPlanPage.get_seasonSelectionDirectiveButton(season).Click();
            }
           ReportCustom.Report(logStack, driver, "Season selected : " + season, "Season selected : " + season);
          //  ReportCustom.ReportEvent("Season selected", "Season selected : " + season, HP.LFT.Report.Status.Passed);
            pPlanPage.waitForPageToLoad();
            Assert.IsTrue(pPlanPage.get_ProductNameTableHeaderElement().Displayed);
            if (!pPlanPage.SectionDirectrivesDropDownElementSelectedElement.Text.Contains(directiveOther))
            {
                SelectElement ele = new SelectElement(pPlanPage.SectionDirectrivesDropDownElement);
                ele.SelectByText(directiveOther);
            }
            pPlanPage.waitForSpinnerToDisappear();
            pPlanPage.waitForPageToLoad();
            pPlanPage.get_UnSelectedTabs()[pPlanPage.get_UnSelectedTabs().Count - 1].Click();
            //pPlanPage.waitForSpinnerToDisappear();
            pPlanPage.waitForPageToLoad();
            Thread.Sleep(4000);
            pPlanPage.get_seasonSelectionDirectiveButton(changeseason).Click();
                ReportCustom.Report(logStack, driver, "Season selected : " + changeseason, "Season selected : " + changeseason);
               pPlanPage.waitForPageToLoad();
                Thread.Sleep(5000);
                Assert.IsTrue(pPlanPage.get_ProductNameTableHeaderElement().Displayed);

                //   pPlanPage.get_TotalBuyStatusIconExpand().Click();
                // pPlanPage.get_TotalBuyStatusTableLastColumnPercentKeys()[1].Click();
                //Assert.IsTrue(pPlanPage.get_TotalBuyStatusTableLastColumnPercentKeys().Count > 5);
                //------------------------------ REPORTING-----------------------------------------------------------
                System.Environment.SetEnvironmentVariable("Endtime", DateTime.Now.ToString("yyyy:MM:dd_HH:mm:ss"));
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.EndTime.ToString(), VersionConrol.getTimeStamp());
                TestBase.sTestName = stestcaseid;
                ReportCustom swriteresult = new ReportCustom();
                swriteresult.WriteResults(stestcaseid, logStack);
                TestBase.sbrowsertype = System.Environment.GetEnvironmentVariable("Browser1");
                TestBase.DriverCleardown = Driver;

                //----------------------VIDEO RECORDING END--------------------------------
                //stopVideoRecording();
                //-------------------------------------------------------------------------
                finalResult = true;
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.TestResult.ToString(), "Pass");
                return finalResult;

            }
            catch (Exception e)
            {
                System.Environment.SetEnvironmentVariable("Endtime", DateTime.Now.ToString("yyyy:MM:dd_HH:mm:ss"));
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.EndTime.ToString(), VersionConrol.getTimeStamp());
                TestBase.sbrowsertype = System.Environment.GetEnvironmentVariable("Browser2");
                TestBase.sTestName = stestcaseid;
                ReportCustom swriteresult = new ReportCustom();
                swriteresult.WriteResults(stestcaseid, logStack, finalResult);
                Console.WriteLine("StackTrace from Test Class :" + e.StackTrace);
                Console.WriteLine("Message from Test Class" + e.Message);

                System.Diagnostics.Debug.WriteLine
                    ("StackTrace from Test Class :" + e.StackTrace);
                System.Diagnostics.Debug.WriteLine("Message from Test Class" + e.Message);

                //stopVideoRecording();
                TestBase.DriverCleardown = Driver;
                VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.TestResult.ToString(), "Fail");
                return finalResult;


            }

        }
    }
}
