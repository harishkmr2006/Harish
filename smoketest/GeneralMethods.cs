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
using OpenQA.Selenium.Interactions;

using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.PageObjects;
using HP.LFT.Report;

namespace SITSmokeTests
{
   public class GeneralMethods
    {
       public void WebButton_Click(IWebElement ele)
        {
            try
            {
                if (ele.Enabled)
                {
                    string text = ele.Text;
                    ele.Click();
                    Reporter.ReportEvent("WebButton Click", "Verify WebButton" + text + " enabled and Clicked-", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("WebButton Click", "Verify WebButton " + ele.Text + " disabled ", HP.LFT.Report.Status.Failed);
                }
            }catch(Exception e)
            {
                Reporter.ReportEvent("WebButton Click", "Verify WebButton" + ele.Text + " enabled and Clicked thrown an exception- "+ e.Message, HP.LFT.Report.Status.Failed);
            }
           
        }
        public void WebButton_Click(IWebElement ele, string WebButtonName)
        {
            try
            {
                if (ele.Enabled)
                {
                    ele.Click();
                    Reporter.ReportEvent("WebButton Click", "Verify WebButton " + WebButtonName + " enabled and Clicked", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("WebButton Click", "Verify WebButton " + WebButtonName + " enabled ", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("WebButton Click", "Verify WebButton " + WebButtonName + " enabled and Clicked thrown an exception- " + e.Message, HP.LFT.Report.Status.Failed);
            }

        }

        public void WebLink_Click(IWebElement ele)
        {
            try
            {
                if (ele.Enabled)
                {
                    ele.Click();
                    Reporter.ReportEvent("WebLink Click", "Verify WebLink" + ele.Text + " enabled and Clicked-", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("WebLink Click", "Verify WebLink " + ele.Text + " enabled ", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("WebLink Click", "Verify WebLink" + ele.Text + " enabled and Clicked thrown an exception- " + e.Message, HP.LFT.Report.Status.Failed);
            }

        }
        public void WebLink_Click(IWebElement ele,string sWebLinkValue)
        {
            try
            {
                if (ele.Enabled)
                {
                    ele.Click();
                    Reporter.ReportEvent("WebLink Click", "Verify WebLink" + sWebLinkValue + " enabled and Clicked-", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("WebLink Click", "Verify WebLink " + sWebLinkValue + " enabled ", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("WebLink Click", "Verify WebLink" + sWebLinkValue + " enabled and Clicked thrown an exception- " + e.Message, HP.LFT.Report.Status.Failed);
            }

        }
        public void WebEdit_SetValue(IWebElement ele, string sValue)
        {
            try
            {
                if (ele.Enabled)
                {
                   // ele.Click();
                    ele.SendKeys(sValue);
                    Reporter.ReportEvent("WebEdit Set Value", "Verify WebEdit" + ele.Text + " enabled and Value Set", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("WebEdit Set Value", "Verify WebEdit " + ele.Text + " enabled ", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("WebEdit Set Value", "Verify WebEdit" + ele.Text + " enabled and Set Value thrown an exception- " + e.Message, HP.LFT.Report.Status.Failed);
            }

        }
       
        public void WebEdit_SetValue(IWebElement ele,string WebEditName, string sValue)
        {
            try
            {
                
                if (ele.Enabled)
                {
                    ele.Click();
                    ele.SendKeys(sValue);
                    Reporter.ReportEvent("WebEdit Set Value", "Verify WebEdit" + WebEditName + " enabled and Value Set is "+ sValue, HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("WebEdit Set Value", "Verify WebEdit " + WebEditName + " enabled to set value  "+ sValue, HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("WebEdit Set Value", "Verify WebEdit" + WebEditName + " enabled and Set Value "+ sValue+" thrown an exception- " + e.Message, HP.LFT.Report.Status.Failed);
            }

        }

        public void WebEdit_SendKeysValue(IWebDriver idriver, IWebElement ele, string WebEditName, string sValue)
        {
            try
            {
                Actions act = new Actions(idriver);
                act.SendKeys(ele, OpenQA.Selenium.Keys.Tab);
                if (ele.Enabled)
                {
                    act.SendKeys(ele, sValue);
                    Reporter.ReportEvent("WebEdit Set Value", "Verify WebEdit" + WebEditName + " enabled and Value Set", HP.LFT.Report.Status.Passed);
                }
                else
                {
                    Reporter.ReportEvent("WebEdit Set Value", "Verify WebEdit " + WebEditName + " enabled ", HP.LFT.Report.Status.Failed);
                }
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("WebEdit Set Value", "Verify WebEdit" + WebEditName + " enabled and Set Value thrown an exception- " + e.Message, HP.LFT.Report.Status.Failed);
            }

        }

        public  void SelectDropDownByValue(IWebElement ele, string value)
        {
            try
            {
                new SelectElement(ele).SelectByText(value);
                Thread.Sleep(1000);
                string sseltext = new SelectElement(ele).AllSelectedOptions.SingleOrDefault().Text;

                if (sseltext.Trim() == value.Trim())
                {
                    // Console.WriteLine("Value Selected successfully from the dropdown field --" + value);
                    Reporter.ReportEvent("Drop Down Selection", "Drop down selection for Value "+ value, HP.LFT.Report.Status.Passed);
                }
                else
                {
                    // Console.Error.WriteLine("Value not Selected successfully from the dropdown field --" + value);
                    Reporter.ReportEvent("Drop Down Selection", "Drop down selection for Value " + value, HP.LFT.Report.Status.Failed);
                }
            }
            catch
            {
                Reporter.ReportEvent("Drop Down Selection", "Drop down selection for Value " + value, HP.LFT.Report.Status.Failed);

            }
        }

        public void SelectDropDownByValue(IWebElement ele, string DropDownName, string value)
        {
            try
            {
                new SelectElement(ele).SelectByText(value);
                Thread.Sleep(1000);
                string sseltext = new SelectElement(ele).AllSelectedOptions.SingleOrDefault().Text;

                if (sseltext.Trim() == value.Trim())
                {
                    // Console.WriteLine("Value Selected successfully from the dropdown field --" + value);
                    Reporter.ReportEvent("Drop Down Selection", DropDownName+" Drop down selection for Value " + value, HP.LFT.Report.Status.Passed);
                }
                else
                {
                    // Console.Error.WriteLine("Value not Selected successfully from the dropdown field --" + value);
                    Reporter.ReportEvent("Drop Down Selection", DropDownName+" Drop down selection for Value " + value, HP.LFT.Report.Status.Failed);
                }
            }
            catch
            {
                Reporter.ReportEvent("Drop Down Selection", DropDownName+" Drop down selection for Value " + value, HP.LFT.Report.Status.Failed);

            }
        }
        public Boolean selectByVisibleText(IWebElement element, String testdata)
        {
            bool executedActionStatus = false;
            try
            {

                SelectElement select = new SelectElement(element);
                select.SelectByText(testdata.Trim());
                executedActionStatus = true;
                Reporter.ReportEvent("Drop Down Selection", "Drop down selection for Value " + testdata, HP.LFT.Report.Status.Passed);
            }
            catch (Exception ex)
            {

                Reporter.ReportEvent("Drop Down Selection", "Drop down selection for Value " + testdata, HP.LFT.Report.Status.Failed);
            }
            return executedActionStatus;
        }
        public string GetTextFromDDL(IWebElement element)
        {
            return new SelectElement(element).AllSelectedOptions.SingleOrDefault().Text;
        }
        public void ClickRadioBtnwithValue(IList<IWebElement> element, string sname)
        {

            try
            {
                IList<IWebElement> oRadiobtns = element;

              
                int Size = oRadiobtns.Count;

              
                Boolean bfound = false;
                for (int i = 0; i < Size; i++)
                {
               
                    String Value = oRadiobtns.ElementAt(i).GetAttribute("value");

                   
                    if (Value.Equals(sname))
                    {
                        Assert.That(Value == sname);
                        bfound = true;
                        oRadiobtns.ElementAt(i).Click();
                       
                     

                        if (oRadiobtns.ElementAt(i).Selected.Equals(true))
                        {
                            Reporter.ReportEvent("Radio Button Selection", "Radio Button Selection for " + sname, HP.LFT.Report.Status.Passed);
                        }
                        else
                        {
                            Reporter.ReportEvent("Radio Button Selection", "Radio Button Selection for " + sname, HP.LFT.Report.Status.Failed);
                        }
                        break;
                    }

                }
                if (bfound == false)
                {
                    Reporter.ReportEvent("Radio Button Selection", "Radio Button Selection for " + sname, HP.LFT.Report.Status.Failed);
                }
            }
            catch
            {
                Reporter.ReportEvent("Radio Button Selection", "Radio Button Selection for " + sname, HP.LFT.Report.Status.Failed);
            }
        }

        public static void CheckCheckboxwithName(IList<IWebElement> element, string sname)
        {
            try
            {
                IList<IWebElement> oCheckBox = element;

            
                int Size = oCheckBox.Count;
                Boolean bfound = false;
               
                for (int i = 0; i < Size; i++)
                {
                  
                    String Value = oCheckBox.ElementAt(i).GetAttribute("name");

                  
                    if (Value.Equals(sname))
                    {
                        Assert.That(Value == sname);
                        bfound = true;
                        oCheckBox.ElementAt(i).Click();
                       // Reporter.ReportEvent("Check box Selection", "Check box Selection for " + sname, HP.LFT.Report.Status.Passed);
                        if (oCheckBox.ElementAt(i).Selected.Equals(true))
                        {
                            Reporter.ReportEvent("Check box Selection", "Check box Selection for " + sname, HP.LFT.Report.Status.Passed);
                        }
                        else
                        {
                            Reporter.ReportEvent("Check box Selection", "Check box Selection for " + sname, HP.LFT.Report.Status.Failed);
                        }
                       
                        break;
                    }
                }
                if (bfound == false)
                {
                    Reporter.ReportEvent("Check box Selection", "Check box Selection for " + sname, HP.LFT.Report.Status.Failed);
                }
            }
            catch
            {
                Reporter.ReportEvent("Check box Selection", "Check box Selection for " + sname, HP.LFT.Report.Status.Failed);
            }
        }
        public void AlertAccept(IWebDriver idriver)
        {
            IAlert alert = idriver.SwitchTo().Alert();
            alert.Accept();
        }
        public void GetLatestWindow(IWebDriver idriver)
        {
            List<string> lswins = idriver.WindowHandles.ToList();
            foreach (string swin in lswins)
            {
                idriver.SwitchTo().Window(swin);

            }
        }
        public void GoToDefaultWindow(IWebDriver idriver)
        {
            idriver.SwitchTo().DefaultContent();
        }
        public void SwitchFrame_Name(IWebDriver idriver, string FrameName)
        {
            idriver.SwitchTo().Frame(FrameName);
        }

        public void SwitchFrame_WebElelemnt(IWebDriver idriver, IWebElement FrameElelmentName)
        {
            idriver.SwitchTo().Frame(FrameElelmentName);
        }

        public string GetCurrentWindow(IWebDriver idriver)
        {
            return idriver.CurrentWindowHandle;
        }

        public void WaitElementToExists(IWebDriver idriver, string xpath)
        {
            var wait = new WebDriverWait(idriver, TimeSpan.FromSeconds(1800));

            wait.Until(ExpectedConditions.ElementExists(By.XPath(xpath)));
        }

        public void WaitFrametoSwitch(IWebDriver idriver, string Framename)
        {
            var wait = new WebDriverWait(idriver, TimeSpan.FromSeconds(1800));
            wait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(Framename));
        }

        public void WaitAlerttoPresent(IWebDriver idriver)
        {
            var wait = new WebDriverWait(idriver, TimeSpan.FromSeconds(1800));
            wait.Until(ExpectedConditions.AlertIsPresent());
        }

        public void TAB_Click(IWebElement ele, IWebDriver idriver)
        {
            Actions act = new Actions(idriver);
            act.SendKeys(ele, OpenQA.Selenium.Keys.Tab);
        }


    }

}
