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
    class SIT_Library_UI
    {
        IWebDriver driver;
        private WebDriverWait wait;
        private WebDriverWait extWait;
        static string currentWindow;
        GeneralMethods sGMethods;

        public SIT_Library_UI(IWebDriver drivername)
        {
            PageFactory.InitElements(drivername, this);
            driver = drivername;
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            extWait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            sGMethods = new GeneralMethods();
        }

        [FindsBy(How = How.XPath, Using = ".//*[@id='selectEnv']")]
        public IWebElement Dropdown_Env { get; set; }

        public IWebElement get_LibraryFromHeader()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//li[text()='Library']/following-sibling::li[1]//img")));
        }

        public IWebElement get_LicenseCompanies()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//label[text()='License Companies']")));
        }

        public IWebElement get_TransportPacking()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//label[text()='Transport Packing']")));
        }

        public IWebElement get_ConsumerPackaging()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//label[text()='Consumer Packaging']")));
        }

        public IWebElement get_Colours()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//label[text()='Colours']")));
        }

        public IWebElement get_Materials()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//label[text()='Materials']")));
        }


        public void traverseContentFrame()
        {
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("content");
        }

        public void LaunchLicenseCompanies()
        {
            try
            {
                get_LibraryFromHeader().Click();
                get_LicenseCompanies().Click();
            }
            catch (Exception ex)
            {
                get_LibraryFromHeader().Click();
                get_LicenseCompanies().Click();
            }
        }

        public void LaunchConsumerPackaging()
        {
            try
            {
                get_LibraryFromHeader().Click();
                get_ConsumerPackaging().Click();
            }
            catch (Exception ex)
            {
                get_LibraryFromHeader().Click();
                get_ConsumerPackaging().Click();
            }
        }


        public void LaunchColours()
        {
            try
            {
                get_LibraryFromHeader().Click();
                get_Colours().Click();
            }
            catch (Exception ex)
            {
                get_LibraryFromHeader().Click();
                get_Colours().Click();
            }
        }

        public void Materials()
        {
            try
            {
                get_LibraryFromHeader().Click();
                get_Materials().Click();
            }
            catch (Exception ex)
            {
                get_LibraryFromHeader().Click();
                get_Materials().Click();
            }
        }

        public void LaunchTransportPacking()
        {
            try
            {
                get_LibraryFromHeader().Click();
                get_TransportPacking().Click();
            }
            catch (Exception ex)
            {
                get_LibraryFromHeader().Click();
                get_TransportPacking().Click();
            }
        }

        public void traverseToContentBodyFrame()
        {
            IWebElement ele = extWait.Until(
                ExpectedConditions.ElementExists(By.XPath("//div[contains(@style,'visibility: visible')]//iframe")));
            driver.SwitchTo().Frame(ele);
            extWait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.Name("tableContentFrame")));
            extWait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.Id("tableBodyRight")));
        }

        public void traverseToContentBodyFrameTableSettings()
        {
            IWebElement ele = extWait.Until(
                ExpectedConditions.ElementExists(By.XPath("//div[contains(@style,'visibility: visible')]//iframe")));
            driver.SwitchTo().Frame(ele);
            extWait.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.Name("tableContentFrame")));
        }

        public void tabFrame()
        {
            IWebElement ele = extWait.Until(
                ExpectedConditions.ElementExists(By.XPath("//div[contains(@style,'visibility: visible')]//iframe")));
            driver.SwitchTo().Frame(ele);
        }

        public List<IWebElement> get_LicenseCompaniesTableCompaniesName()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//table//tr//td[1]//a[contains(@href,'java')]")))
                .ToList();
        }

        public List<IWebElement> get_materialCompositionType()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//*[@class='materialCompositionType']")))
                .ToList();
        }



        public List<IWebElement> get_LicenseCompaniesTableRowRecord()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//table//tr//td[text()='Inactive']/parent::tr")))
                .ToList();
        }


        public List<IWebElement> get_LicenseCompaniesTableRowRecordName()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//table//tr//td[text()='Inactive']/preceding-sibling::td[3]")))
                .ToList();
        }


        public List<IWebElement> get_LicenseCompaniesTableRowRecordCheckbox()
        {
            return extWait.Until(
                    ExpectedConditions.PresenceOfAllElementsLocatedBy(
                        By.XPath("//table//tr//input[@type='checkbox']")))
                .ToList();
        }

        public IWebElement get_Drop()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//li[contains(@class,'enabled')]//*[text()='Drop']")));
        }

        public IWebElement get_Activate()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//li[contains(@class,'enabled')]//*[text()='Activate']")));
        }

        public IWebElement get_Checkbox(int index)
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//table//tr[" + index + "]//input[@type='checkbox']")));
        }

        public IWebElement get_StatusByCompanyName(String companyName)
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//table//td[contains(.,'" + companyName + "')]/following-sibling::td[3]")));
        }

        public IWebElement get_RowByCompanyName(String companyName)
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//table//td[contains(.,'" + companyName + "')]/parent::tr")));
        }

        public IWebElement get_ActionLinkForDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//span[text()='Actions']")));
        }

        public IWebElement get_CreateTransportPackingForDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//span[text()='Create Transport Packing']")));
        }

        public IWebElement get_CreateConsumerPackagingForDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//span[text()='Create Consumer Packaging']")));
        }

        public IWebElement get_CreateCreateFancyForDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//span[text()='Create Denim']")));
        }


        public IWebElement get_CreateTransportPackingText()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(text(),'Create Transport Packing')]")));
        }

        public IWebElement get_CreateConsumerPackagingText()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(text(),'Create Consumer Packaging')]")));
        }

        public IWebElement get_MaterialText()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//td//*[contains(text(),'Create')]")));
        }


        public IWebElement get_showTypeSelector()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//input[contains(@onclick,'showTypeSelector')]")));
        }

        public IWebElement get_FabricTypeDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath(".//*[@id='FabricTypeId']")));
        }

        public IWebElement get_TransportPackingExpand()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(@onclick,'Transport Packing')]")));
        }

        public IWebElement get_PackagingExpand()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[contains(@onclick,'Consumer Packaging')]")));
        }


        public IWebElement get_PolybagRadioButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//input[contains(@value,'Poly bag')]")));
        }

        public IWebElement get_BoxRadioButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//input[contains(@value,'Box')]")));
        }

        public void switchToFrame()
        {
            try
            {
                driver.SwitchTo().Frame("detailsDisplay");
            }
            catch
            {
                driver.SwitchTo().Frame("detailsDisplay");
            }
        }


        public IWebElement get_SelectButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Select']")));
        }

        public IWebElement get_txtDescription()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//textarea[@name='txtDescription']")));
        }

        public IWebElement get_txtDescriptionTextBox()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath(".//*[@id='DescriptionId']")));
        }

        public IWebElement get_OfficeIdDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath(".//*[@id='OfficeId']")));
        }

        public IWebElement get_Concept()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//*[@id='InternalBrand_html']//option[2]")));
        }

        public IWebElement get_SeasonDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath("//select[@name='txtSeasonCreated']")));
        }

        public IWebElement get_SeasonIdDropDown()
        {
            return extWait.Until(
                ExpectedConditions.ElementIsVisible(
                    By.XPath(".//*[@id='SeasonId']")));
        }

        public void selectSeason(string value)
        {
            SelectElement ele = new SelectElement(get_SeasonDropDown());
            ele.SelectByValue(value);
        }

        public IWebElement get_DoneButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Done']")));
        }

        public IWebElement get_DoneButtonPopUp()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("(//a[text()='Done'])[1]")));
        }



        public IWebElement get_FindButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Find']")));
        }

        public IWebElement get_MaterialTab(string tabName)
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//li/div/a[text()='" + tabName + "']")));
        }

        public IWebElement get_SubmitButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//button[text()='Submit']")));
        }

        public IWebElement get_EditButton()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@title='Edit']/img")));
        }

        public IWebElement get_DoneButtonCastor()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//a[text() = 'Done']")));
        }

        public IWebElement get_stateNameHighlight(string Name)
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@class='stateNameHighlight' and text()='" + Name + "']")));
        }

        public IWebElement get_StatePromote()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@id='AEFLifecyclePromote']/img")));
        }


        public IWebElement get_Initiatedlink()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[text()='Initiated']")));
        }

        public IWebElement get_btnSupplier()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@name='btnSupplier']")));
        }

        public IWebElement get_btnFiberContent()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@name='btnFiberContent']")));
        }

        public IWebElement get_btnPurchaseCost()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("//*[@id='PurchaseCost']")));
        }

        public IWebElement get_btnCurrencyId()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath(".//*[@id='CurrencyId']")));
        }

        public IWebElement get_btnPurchaseUOMId()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath(".//*[@id='PurchaseUOMId']")));
        }




        public IWebElement get_btnSupplierCheckBox()
        {
            return extWait.Until(
                ExpectedConditions.ElementToBeClickable(
                    By.XPath("(//table//input)[3]")));
        }


        public string get_libraryWindow()
        {
            driver.SwitchTo().DefaultContent();
            return driver.CurrentWindowHandle;
        }

        public void get_NewLaunchedWindow()
        {
            foreach (var eleHandle in driver.WindowHandles)
            {
                driver.SwitchTo().Window(eleHandle);
            }
        }

    }
}
