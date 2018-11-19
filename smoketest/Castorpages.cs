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
    public class Castorpages
    {
        IWebDriver driver;
        public int MINWAITTIME = 10;
        public int AVGWAITTIME = 30;
        public int MAXWAITTIME = 90;
        public Castorpages(IWebDriver drivername)
        {
            
            PageFactory.InitElements(drivername, this);

            driver = drivername;

        }

        [FindsBy(How = How.XPath, Using = ".//*[@id='divLogin']//input[@value=\"SEBO PD Buyer\"]")]
        public IWebElement WebButton_SEBOPDBUYER { get; set; }

        public string frame_Content = "content";
        public string frame_tvcTabs0_ProductDevelopmentTabcontentFrame = "tvcTabs0_HMProductDevelopmentBuyerStartPageTabcontentFrame";

        [FindsBy(How = How.XPath, Using = ".//*[@id='tvcTabs0']//a[@title=\"Product Developments\"]")]
        public IWebElement Tab_ProductDevelopments { get; set; }

        public string frame_tvcTabs0_ProductDevelopmentcontentFrame = "tvcTabs0_HMProductDevelopmentBuyerStartPagecontentFrame";

        [FindsBy(How = How.XPath, Using = ".//*[@id='tvcTabs0']//a[@title=\"Product Developments\"]")]
        public IWebElement ProductDevelopments { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[@class='content']/iframe[@class='frame']")]
        public IWebElement CreateProduct_FrameElelemnt { get; set; }

        // IWebElement Frame_Createproduct_Elelment = driver.FindElement(By.XPath("//*[@class='content']/iframe[@class='frame']"));
        public string sCreatePRODXPATH = ".//*[@id='tvcTabs0']//a[@title=\"Product Developments\"]";

        [FindsBy(How = How.XPath, Using = ".//span[text()='Create Product Development']")]
        public IWebElement Link_CreateProductDevelopment { get; set; }
        public string xpath_ProductDevelopmentName = ".//input[@id=\"txtProductName\"]";

        [FindsBy(How = How.XPath, Using = ".//input[@id=\"txtProductName\"]")]
        public IWebElement WebEdit_ProductDevelopmentName { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='owningOffice']")]
        public IWebElement DropDown_OwningOffice { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='txtCompanyId']")]
        public IWebElement DropDown_Brand { get; set; }

        [FindsBy(How = How.XPath, Using = ".//a[contains(.,\"Create\")]")]
        public IWebElement WebButton_Create { get; set; }

        [FindsBy(How = How.XPath, Using = ".//ul[@class=\"he-tb-toolbar vertical he-tb-slim\"]//i[@class=\"fa fa-pencil\"]")]
        public IWebElement EditButton_ProductDevelopment { get; set; }

        [FindsBy(How = How.XPath, Using = ".//label[contains(text(),\"Season\")]/..//div[@class=\"selectize-input items not-full has-options\"]")]
        public IWebElement DropDown_Season { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='6-2017']")]
        public IWebElement Options_Season { get; set; }

        [FindsBy(How = How.XPath, Using = ".//label[contains(text(),\"Department\")]/..//div[@class=\"selectize-input items not-full has-options\"]")]
        public IWebElement DropDown_Department { get; set; }
        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='1201']")]
        public IWebElement Options_Department { get; set; }

        [FindsBy(How = How.XPath, Using = ".//label[contains(text(),\"Concept\")]/..//div[contains(@class,\"selectize-input items\")]")]
        public IWebElement DropDown_Concept { get; set; }
        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='Everyday Collection']")]
        public IWebElement Options_Concept { get; set; }

        [FindsBy(How = How.XPath, Using = ".//label[contains(text(),\"Product Type\")]/..//div[contains(@class,\"selectize-input items not-full\")]")]
        public IWebElement DropDown_ProductType { get; set; }
        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='Other jacket - Top']")]
        public IWebElement Options_ProductType { get; set; }

        [FindsBy(How = How.XPath, Using = ".//label[contains(text(),\"Sales Mode\")]/..//div[@class=\"selectize-input items full has-options has-items\"]")]
        public IWebElement DropDown_SalesMode { get; set; }
        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='Single']")]
        public IWebElement Options_SalesMode { get; set; }

        [FindsBy(How = How.XPath, Using = ".//label[contains(text(),\"Licensed\")]/..//div[@class=\"selectize-input items full has-options has-items\"]")]
        public IWebElement DropDown_Licensed { get; set; }
        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='No']")]
        public IWebElement Options_Licensed { get; set; }

        [FindsBy(How = How.XPath, Using = ".//label[contains(text(),\"Currency\")]/..//div[contains(@class,\"selectize-input items not-full has-options\")]")]
        public IWebElement DropDown_Currency { get; set; }
        [FindsBy(How = How.XPath, Using = "//*[contains(@class,'selectize-dropdown') and contains(@style,'block')]//*[text()='SEK']")]
        public IWebElement Options_Currency { get; set; }

        [FindsBy(How = How.XPath, Using = ".//label[contains(text(),\"Target Price\")]/../input[contains(@name,\"_targetPrice\")]")]
        public IWebElement WebEdit_TargetPrice { get; set; }
        [FindsBy(How = How.XPath, Using = ".//button[contains(text(),'Save')]")]
        public IWebElement WebButton_Save { get; set; }

        // [FindsBy(How = How.XPath, Using = ".//*[@id='tab_tvc-page-hm-productdevelopment-original-ProductDevelopment-xml-ts0-tDescription']")]
        [FindsBy(How = How.XPath, Using = ".//a[contains(text(),\"Description\")]")]
        public IWebElement Tab_Description { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='Appearances']//a/span[contains(.,\"Create Appearance(s)\")]")]
        public IWebElement Link_CreateAppearance { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='colorName']")]
        public IWebElement WebEdit_Colour { get; set; }
        [FindsBy(How = How.XPath, Using = ".//table/tbody/tr//td[contains(.,\"09-101 \")]")]
        public IWebElement Options_Colour { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='txtAppearanceName']")]
        public IWebElement WebEdit_Appearance { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='.tvc.btn.continue.a']/a")]
        public IWebElement WebButton_Appearance_Create { get; set; }

        //  [FindsBy(How = How.XPath, Using = ".//*[@id='tab_tvc-page-hm-productdevelopment-original-ProductDevelopment-xml-ts0-tLabelAndPackaging']")]
        [FindsBy(How = How.XPath, Using = ".//a[contains(text(),\"Label and Packaging\")]")]
        public IWebElement WebLink_tab_LabelAndPackaging { get; set; }
        // [FindsBy(How = How.XPath, Using = ".//div[@id=\"tabContent_tvc-page-hm-productdevelopment-original-ProductDevelopment-xml-ts0-tLabelAndPackaging\"]//ul[@class=\"he-tb-toolbar\"]//span[contains(.,\"Actions\")]")]
        [FindsBy(How = How.XPath, Using = ".//div[@data-title=\"Label and Packaging\"]//ul[@class=\"he-tb-toolbar\"]//span[contains(.,\"Actions\")]")]
        public IWebElement WebLink_Label_Actions { get; set; }
        // [FindsBy(How = How.XPath, Using = "//div[@id=\"tabContent_tvc-page-hm-productdevelopment-original-ProductDevelopment-xml-ts0-tLabelAndPackaging\"]//ul[@class=\"he-tb-toolbar\"]//ul[@class=\"he-tb-menu\"]//span[text()=\"Add Label\"]")]
        [FindsBy(How = How.XPath, Using = "//div[@data-title=\"Label and Packaging\"]//ul[@class=\"he-tb-toolbar\"]//ul[@class=\"he-tb-menu\"]//span[text()=\"Add Label\"]")]
        public IWebElement WebLink_AddLabel { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='txtTextSearch']")]
        public IWebElement WebEdit_AddLabel_Search { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='mx_btn-search']")]
        public IWebElement WebButton_AddLabel_Search { get; set; }

        [FindsBy(How = How.XPath, Using = "(//table[@id=\"tbl\"]/tbody/tr/td[1]/input[@type=\"checkbox\"])[1]")]
        public IWebElement WebCheckBox_AddLabel_SearchResult1 { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']//span[contains(.,\"Add Selected\")]")]
        public IWebElement WebButton_AddLabel_AddSelected { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='tab_tvc-page-hm-productdevelopment-original-ProductDevelopment-xml-ts0-tOfficeOverview']")]
        public IWebElement WebLink_tab_Offices { get; set; }

        // [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']//li[Contains(.,\"Add Office\")]")]
        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul[1]/li[1]")]
        public IWebElement WebLink_Label_AddOffice { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[@id=\"contentTable\"]/tbody/tr/td/div[contains(.,\"Name: CNHK\")]/../../td/input[@type=\"checkbox\"]")]
        public IWebElement WebCheck_Office_CNHK { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul/li[contains(.,\"Add Selected\")]/img")]
        public IWebElement WebLink_Office_AddSelected { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table/thead/tr/th/input[@type=\"checkbox\"]")]
        public IWebElement WebCheck_Office_Publish { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul[1]/li[contains(.,\"Publish\")]/img")]
        public IWebElement WebLink_Office_Publish { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='to_0']/td/div[@ng-model=\"office.to\"]/div/input")]
        public IWebElement WebEdit_Publish_TOField { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='ui-select-choices-row-0-0']")]
        public IWebElement Weboption_Publish_Searchoption1 { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='saveButton']")]
        public IWebElement Webbutton_Publish_Publish { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='selectEnv']")]
        public IWebElement Dropdown_Env_ICC { get; set; }


        //[FindsBy(How = How.XPath, Using = "//div[@id=\"GTBsearchDiv\"]//input[@id=\"GlobalNewTEXT\"]")]
        //public IWebElement PDMerch_SearchBox { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='GlobalNewTEXT']")]
        public IWebElement PDMerch_SearchBox { get; set; }

        [FindsBy(How = How.XPath, Using = ".//div[contains(.,\"Weight in gr for 1 piece/pack\")]/input")]
        public IWebElement PDMerch_PDEdit_Weight { get; set; }

        [FindsBy(How = How.XPath, Using = ".//a[text()=\"Supplier and Request\"]")]
        public IWebElement PDMerch_WebLink_tab_SupplierandRequest { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul/li[contains(.,\"Add Supplier\")]/img")]
        public IWebElement PDMerch_WebLink_AddSupplier { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='mx_btn-search']")]
        public IWebElement PDMerch_Search_AddSupplier { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='full_search']//button[@class=\"cancel\"]")]
        public IWebElement PDMerch_miniClose_AddSupplieroptions { get; set; }

        [FindsBy(How = How.XPath, Using = ".//a[text()=\"Options\"]")]
        public IWebElement PDMerch_WebLink_tab_options { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[@id=\"tbl\"]/tbody/tr/td/input[@type=\"checkbox\"]")]
        public IWebElement PDMerch_WebLink_optionGroupCheckBox { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[@id=\"tbl\"]/tbody/tr/td/input[@type=\"checkbox\"]")]
        public IWebElement PDMerch_WebLink_optionSuuplierCheckBox { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='TopPanelOuterDiv']/div/img[@title=\"Enter edit mode\"]")]
        public IWebElement PDMerch_WebButton_OptionGroupEditbutton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='productionGroup']")]
        public IWebElement PDMerch_WebList_OptionGroupEditPage_ProductGroup { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='productionType']")]
        public IWebElement PDMerch_WebList_OptionGroupEditPage_productionType { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='productionAppearance']")]
        public IWebElement PDMerch_WebList_OptionGroupEditPage_productionApperance { get; set; }

        [FindsBy(How = How.XPath, Using = ".//select[@label=\"Country of Production\"]")]
        public IWebElement PDMerch_WebList_OptionGroupEditPage_COP { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul[1]/li[contains(.,\"Action\")]/img")]
        public IWebElement PDMerch_Action_OptionGroupEditPage_BOM { get; set; }

        [FindsBy(How = How.XPath, Using = "html/body/div//ul/li[contains(.,\"Add Leather...\")]/div/img")]
        public IWebElement PDMerch_ActionOption_OptionGroupEditPage_BOM { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='mx_btn-search']")]
        public IWebElement PDMerch_LeatherSearch_OptionGroupEditPage_BOM { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='position']")]
        public IWebElement PDMerch_WebList_Position_OptionGroupEditPage_BOM { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='.tvc.btn.apply.a']/a")]
        public IWebElement PDMerch_WebButton_PositionAdd_OptionGroupEditPage_BOM { get; set; }

        [FindsBy(How = How.XPath, Using = ".//input[@name=\"tvcSelectAll\"]")]
        public IWebElement PDMerch_WebCheck_SelectAll_OptionGroupEditPage_BOM { get; set; }

        [FindsBy(How = How.XPath, Using = "html/body/div//ul/li[contains(.,\"Edit\")]/div/img")]
        public IWebElement PDMerch_ActionEditOption_OptionGroupEditPage_BOM { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='placement']")]
        public IWebElement PDMerch_EditOption_Placement_OptionGroupEditPage_BOM { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='countryOfOrigin']")]
        public IWebElement PDMerch_EditOption_countryOfOrigin_OptionGroupEditPage_BOM { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[@class=\"content-tableEmpty\"]/tbody/tr/td[contains(.,\"No Request Sent\")]/input[@type=\"checkbox\"]")]
        public IWebElement PDMerch_OptionGroup_SupplierCheckbox { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul[1]/li[contains(.,\"Action\")]/i")]
        public IWebElement PDMerch_Action_OptiontoAdd { get; set; }

        [FindsBy(How = How.XPath, Using = "html/body/div/div/ul/li[contains(.,\"Create Option (in Option Group)\")]/div/img")]
        public IWebElement PDMerch_ActionOption_OptiontoAdd { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='TopPanelOuterDiv']/div/img[@title=\"Enter edit mode\"]")]
        public IWebElement PDMerch_ActionOption_EditButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//div[@id=\"TopPanelOuterDiv\"]/table/tbody/tr/td/div[contains(.,\"Country of Production\")]/div//select")]
        public IWebElement PDMerch_ActionOption_CountryOfProduction { get; set; }
        [FindsBy(How = How.XPath, Using = ".//div[@id=\"TopPanelOuterDiv\"]/table/tbody/tr/td/div[contains(.,\"Country of Delivery\")]/div//select")]
        public IWebElement PDMerch_ActionOption_CountryOfDelivery { get; set; }
        [FindsBy(How = How.XPath, Using = ".//div[@id=\"TopPanelOuterDiv\"]/table/tbody/tr/td/div[contains(.,\"Country of Origin\")]/div//select")]
        public IWebElement PDMerch_ActionOption_CountryOfOrigin { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul/li[contains(.,\"Add Cost\")]/img")]
        public IWebElement PDMerch_ActionOption_AddCostButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='opdTd']/div[@class=\"opdDiv notDelivery\"]/img")]
        public IWebElement PDMerch_ActionOption_AddCost_OrderPlacementDate { get; set; }

        [FindsBy(How = How.XPath, Using = ".//div[@id='ui-datepicker-div']/table/tbody/tr/td/a[contains(@class,\"ui-state-default ui-state-highlight\")]")]
        public IWebElement PDMerch_ActionOption_AddCost_OrderPlacementDateSelect { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='aggregatedCostDivId']/table/tbody/tr[2]/td[2]/input")]
        public IWebElement PDMerch_ActionOption_AddCost_MeterialCost { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='.tvc.btn.continue.a']/a")]
        public IWebElement PDMerch_ActionOption_AddCost_SaveCost { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='TopPanelControl_3']/img")]
        public IWebElement PDMerch_ActionOption_OptionDetails_SetStatus { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='TopPanelControl_2']/img")]
        public IWebElement PDMerch_ActionOption_OptionDetails_Lock { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[contains(@class,\"content-table  lowPriceTable\")]//input[@type=\"checkbox\"]")]
        public IWebElement PDMerch_ActionOption_SupplertoHandOver { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul/li[contains(.,\"Handover Option\")]/img")]
        public IWebElement PDMerch_ActionOption_HandOverLink { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='saveButton']")]
        public IWebElement PDMerch_ActionOption_MailHandOverButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[@id=\"tbr\"]//table//tbody/tr/td[contains(.,\"Handover\")]")]
        public IWebElement PDMerch_ActionOption_HandOvertextToVerify { get; set; }
        [FindsBy(How = How.XPath, Using = ".//table[@id=\"tbr\"]//table//tbody/tr/td[contains(.,\"Received\")]")]
        public IWebElement PDBuyer_ActionOption_ReceivedtextToVerify { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul/li[contains(.,\"Proceed\")]/i")]
        public IWebElement PDBuyer_ActionOption_ProceedLink { get; set; }
        [FindsBy(How = How.XPath, Using = "(html/body/div//div[contains(.,\"Proceed\")]/img)[1]")]
        public IWebElement PDBuyer_ActionOption_ProceedLink_ProccedOption { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='saveButton']")]
        public IWebElement PDBuyer_ActionOption_ProceedLink_RTOButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[@id=\"tbr\"]//table//tbody/tr/td[contains(.,\"Proceed\")]")]
        public IWebElement PDBuyer_ActionOption_ProceedtextToVerify { get; set; }


        [FindsBy(How = How.XPath, Using = "//form/fieldset//input[@name=\"username\"]")]
        public IWebElement Login_WebEdit_UserName { get; set; }
        [FindsBy(How = How.XPath, Using = "//form/fieldset//input[@name=\"password\"]")]
        public IWebElement Login_WebEdit_Password { get; set; }
        [FindsBy(How = How.XPath, Using = "//form//input[@value=\"Log in\"]")]
        public IWebElement Login_WebButton_Login { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='owningOffice']")]
        public IWebElement PDBuyer_Createproduct_OwningOffice { get; set; }

        [FindsBy(How = How.XPath, Using = ".//a[contains(.,\"Capacity Booking\")]")]
        public IWebElement PDMerch_TABS_CapacityBooking { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']//span[contains(.,\"Create New Booking\")]")]
        public IWebElement PDMerch_TABS_CreateNewBooking { get; set; }

        [FindsBy(How = How.XPath, Using = ".//a[contains(.,\"Create Capacity Booking\")]")]
        public IWebElement PDMerch_Button_CreateCapacityBooking { get; set; }

        [FindsBy(How = How.XPath, Using = "//input[@value=\"Save\"]")]
        public IWebElement PDMerch_Button_CapacityBooking_Save { get; set; }

        [FindsBy(How = How.XPath, Using = "//div[contains(@data-bind,\"click: addBooking\")]/i")]
        public IWebElement PDMerch_Button_CapacityBooking_AddBooking { get; set; }
        [FindsBy(How = How.XPath, Using = "//form/table/tbody//td//span[contains(@data-bind,\"text: SelectedSupplier()\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_Supplierspan { get; set; }

        [FindsBy(How = How.XPath, Using = "//form/table/tbody//select[contains(@data-bind,\"options: Suppliers\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_SupplierSelection { get; set; }

        [FindsBy(How = How.XPath, Using = ".//input[contains(@data-bind,\"datepicker: OrderPlacementDate\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_OPD { get; set; }

        [FindsBy(How = How.XPath, Using = ".//input[contains(@data-bind,\"datepicker: TimeOfDelivery\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_TOD { get; set; }

        [FindsBy(How = How.XPath, Using = ".//input[contains(@data-bind,\"datepicker: InShopWeek\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_ISW { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='ui-datepicker-div']/table/tbody/tr/td[contains(@class,\"ui-datepicker-today\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_TodayDate { get; set; }

        [FindsBy(How = How.XPath, Using = ".//input[contains(@data-bind,\"value: NumberOfPacks\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_QTY { get; set; }

        [FindsBy(How = How.XPath, Using = ".//Select[contains(@data-bind,\"options: Store\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_Channelselect { get; set; }
        [FindsBy(How = How.XPath, Using = ".//Select[contains(@data-bind,\"options: Base\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_BaseTrailelect { get; set; }
        [FindsBy(How = How.XPath, Using = "//Select[contains(@data-bind,\"options: availableCountriesOfProduction\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_COP { get; set; }
        [FindsBy(How = How.XPath, Using = "//input[contains(@data-bind,\"click: save\")]")]
        public IWebElement PDMerch_Button_CapacityBooking_SaveButton { get; set; }
        [FindsBy(How = How.XPath, Using = ".//a[contains(.,\"Samples\")]")]
        public IWebElement PDMerch_TABS_Samples { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul/li[contains(.,\"Order Sample\")]/img")]
        public IWebElement WebLink_Samples_OrderSample { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='orderContext']")]
        public IWebElement WebRadio_CreateSampleOrder_Order { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='contentDiv']/data-ng-include//select[@ng-model=\"orderId\"]")]
        public IWebElement WebList_CreateSampleOrder_OrderList { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='sampleType_0']/select")]
        public IWebElement WebList_CreateSampleOrder_SelectSampletype { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='ltCondition_0_0']/select")]
        public IWebElement WebList_CreateSampleOrder_LTCondition { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='productSize_0_0_0']")]
        public IWebElement Webedit_CreateSampleOrder_Size { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='productQuantity_0_0_0_0']")]
        public IWebElement Webedit_CreateSampleOrder_Qty { get; set; }
        [FindsBy(How = How.XPath, Using = ".//div[@id=\"toBeReceived\"]/input")]
        public IWebElement WedDate_CreateSampleOrder_toBeReceivedDate { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='ui-datepicker-div']/table/tbody//td[contains(@class,\"ui-datepicker-days-cell-over\")]")]
        public IWebElement WedDate_CreateSampleOrder_ReceivedDateSelection { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='sendButton']")]
        public IWebElement WebList_CreateSampleOrder_SendButton { get; set; }

        // [FindsBy(How = How.XPath, Using = ".//*[@id='s2id_autogen1']")]
        [FindsBy(How = How.XPath, Using = ".//input[@data-placeholder=\"To...\"]/..//input[@type=\"text\"]")]
        public IWebElement WebEdit_CreateSampleOrder_SendToField { get; set; }
        [FindsBy(How = How.XPath, Using = ".//li[contains(.,\"Create New Contact\")]/div")]
        public IWebElement WebList_CreateSampleOrder_CreateNewcontactOption { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='hmTypeaheadAddContactEmailAddress']")]
        public IWebElement WebEdit_CreateSampleOrder_CreateNewcontact_Emailaddress { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='hmTypeaheadAddContactFirstName']")]
        public IWebElement WebEdit_CreateSampleOrder_CreateNewcontact_Firstname { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='hmTypeaheadAddContactLastName']")]
        public IWebElement WebEdit_CreateSampleOrder_CreateNewcontact_Lastname { get; set; }
        [FindsBy(How = How.XPath, Using = "//input[@value=\"Add\"]")]
        public IWebElement WebButton_CreateSampleOrder_CreateNewcontact_AddButton { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='sendButtonId']")]
        public IWebElement WebButton_CreateSampleOrder_CreateNewcontact_SendButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='tinymce']")]
        public IWebElement WebButton_CreateSampleOrder_CreateNewcontact_Message { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']/ul[@class=\"uip-toolbar-row\"]/li/span[text()=\"Sample\"]")]
        public IWebElement WebButton_ReceivedSample_Sample { get; set; }

        [FindsBy(How = How.XPath, Using = ".//div[@class=\"content-wrapper\"]/span[contains(.,\"Register Sample\")]")]
        public IWebElement WebButton_ReceivedSample_RegisterSample { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='contextSupplier']//input[@value=\"Order\"]")]
        public IWebElement WebRadio_RegisterSample_OrderRadio { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='orderid']")]
        public IWebElement WebList_RegisterSample_OrderidList { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='checkAllLeft']")]
        public IWebElement WebCheck_RegisterSample_AllCheckboxes { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='moveToRight']")]
        public IWebElement WebLink_RegisterSample_MoveRighArrow { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='.tvc.btn.continue.a']/a")]
        public IWebElement WebLink_RegisterSample_RegisterButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[@class=\"contentTable\"]//td[contains(.,\"Counter Sample\")]")]
        public IWebElement WebLink_RegisterSample_Countersamplerecord { get; set; }
        [FindsBy(How = How.XPath, Using = ".//table[@class=\"contentTable\"]//td[contains(.,\"Production Sample\")]")]
        public IWebElement WebLink_RegisterSample_Productionsamplerecord { get; set; }

        [FindsBy(How = How.XPath, Using = ".//input[@name=\"tvcSelectAll\"]")]
        public IWebElement WebLink_RegisterSample_SelectAllCheckbox { get; set; }

        [FindsBy(How = How.XPath, Using = ".//div[@class=\"content-wrapper\"]/span[contains(.,\"Hand Over To Lab\")]")]
        public IWebElement WebButton_ReceivedSample_HandOverToLab { get; set; }
        [FindsBy(How = How.XPath, Using = "(.//td[@tip-title=\"Handed over to Lab by\"])[1]")]
        public IWebElement Tooltip_ReceivedSample_HandOverToLabrecord { get; set; }

        //  [FindsBy(How = How.XPath, Using = ".//*[@id='divToolbar']//table/tbody/tr/td[contains(.,\"Quality Assurance\")]")]
        [FindsBy(How = How.XPath, Using = ".//*[@id='divToolbar']/div//li[contains(.,\"Quality Assurance\")]")]
        public IWebElement Tooltab_QATAB { get; set; }

        [FindsBy(How = How.XPath, Using = ".//a[text()=\"Samples Ready for Test\"]")]
        public IWebElement Tooltab_SampleReadyTestTab { get; set; }

        [FindsBy(How = How.XPath, Using = ".//input[@data-hm-id=\"filterControl_orderNo\"]")]
        public IWebElement WebEdit_Search_Ordernumber { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='footerViewId']//input[@value=\"Apply\"]")]
        public IWebElement WebButton_Search_Apply { get; set; }

        [FindsBy(How = How.XPath, Using = "//td[@class=\"testTypeChoice\"]//a[@class=\"select2-choice\"]")]
        public IWebElement WebList_LabTest_Test { get; set; }
        //  [FindsBy(How = How.XPath, Using = ".//*[@id='s2id_id2ed']/a")]
        [FindsBy(How = How.XPath, Using = ".//div[@class=\"content\"]//a[@class=\"select2-choice\"]")]
        public IWebElement WebList_LabTest_Performer { get; set; }
        [FindsBy(How = How.XPath, Using = ".//input[@name=\"initiateTest\"]")]
        public IWebElement WebButton_LabTest_InitiateTest { get; set; }
        [FindsBy(How = How.XPath, Using = ".//input[@value=\"Handle QA Tests\"]")]
        public IWebElement WebButton_LabTest_HandleQatest { get; set; }
        [FindsBy(How = How.XPath, Using = ".//input[@name=\"tvcSelectAll\"]")]
        public IWebElement WebButton_QATest_SelectAllCheckbox { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']//span[contains(.,\"Enter Test Result\")]")]
        public IWebElement WebButton_QATest_Enterresult { get; set; }
        [FindsBy(How = How.XPath, Using = ".//input[@name=\"decisionByQA\"][@value=\"0\"]")]
        public IWebElement WebButton_QATestresult_Pass { get; set; }
        [FindsBy(How = How.XPath, Using = ".//input[@value=\"Confirm\"]")]
        public IWebElement WebButton_QATestresult_Confirm { get; set; }
        [FindsBy(How = How.XPath, Using = "(.//a[contains(.,\"Back to QA Test List\")])[2]")]
        public IWebElement WebButton_QATestresult_BacktoQAList { get; set; }

        [FindsBy(How = How.XPath, Using = ".//div[@class=\"content-wrapper\"]/span[contains(.,\"Create Sample Report\")]")]
        public IWebElement WebButton_ReceivedSample_CreateSampleReport { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[@class=\"contentTable\"]/tbody/tr[1]/td/input[1]")]
        public IWebElement PDMerch_ReceivedSample_FirstCheckBox { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='buttons']/input[@value=\"Next >>\"]")]
        public IWebElement WebButton_Samplereport_Next { get; set; }

        [FindsBy(How = How.XPath, Using = ".//input[@value=\"Approved\"]")]
        public IWebElement WebRadio_Samplereport_Approve { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='buttons']/input[@value=\"Send >>\"]")]
        public IWebElement WebButton_Samplereport_Send { get; set; }

        [FindsBy(How = How.XPath, Using = ".//a[contains(.,\"Certificates\")]")]
        public IWebElement PDMerch_TABS_Certificates { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']//span[contains(.,\"Show Options\")]")]
        public IWebElement PDMerch_Certificates_ShowOptions { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='page0']//img[@title=\"Assign Korea Certificates\"]")]
        public IWebElement PDMerch_Certificates_AssignCertificates { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='okButton']")]
        public IWebElement PDMerch_AssignCertificates_ApplyButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//table[@class=\"contentTable\"]//tr[1]/td[1]/input[1]")]
        public IWebElement PDMerch_AssignCertificates_FirstCheckBox { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='toolbar-container']//span[contains(.,\"Assign\")]")]
        public IWebElement PDMerch_AssignCertificates_AssignButton { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@filter_name='LABEL_TYPE']")]
        public IWebElement LabelType { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='select_LABEL_TYPE']/option[@value='Price Tag']")]
        public IWebElement PriceTag { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='imgProgressDiv']")]
        public IWebElement Spinner { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='full_search']//button[@class='done']")]
        public IWebElement done { get; set; }

        [FindsBy(How = How.XPath, Using = ".//*[@id='input_LABEL_TYPE']")]
        public IWebElement LabelTypeTextbox { get; set; }
        [FindsBy(How = How.XPath, Using = ".//*[@id='tvc-ms-icon']//preceding::img[1]")]
        public IWebElement LogoutImg { get; set; }

        [FindsBy(How = How.XPath, Using = "//*[text()='Exit Castor']")]
        public IWebElement ExitCastor { get; set; }

        [FindsBy(How = How.XPath, Using = "//a[contains(.,\"Fit Descriptions\")]")]
        public IWebElement Weblink_FitDescription { get; set; }
        [FindsBy(How = How.XPath, Using = "//a[contains(.,\"Add or Copy M-chart\")]")]
        public IWebElement Weblink_Mchart { get; set; }

        public void CastorLogin(string user,string pwd)
        {
            try
            {
                if (Login_WebEdit_UserName.Displayed)
                {
                    Login_WebEdit_UserName.SendKeys(user);
                }

            }
            catch
            {

            }



            Login_WebEdit_Password.SendKeys(pwd);
            Login_WebButton_Login.Click();
            Reporter.ReportEvent("Castor Login with "+ user, "Castor Login with " + user, HP.LFT.Report.Status.Passed);
        }

        public void CastorLogout()
        {
            try
            {
                if (LogoutImg.Displayed)
                {
                    LogoutImg.Click();
                    ExitCastor.Click();
                }

            }
            catch
            {

            }



        }

        public void WebCheck_Office(string officename)
        {
            //string sid= driver.FindElement(By.XPath(".//table[@id=\"tbr\"]/tbody/tr/td/div[contains(.,\"" + officename + "\")]/../..")).GetAttribute("id");
            // string aid = sid.Substring(1, (sid.Length - 1));
            // driver.FindElement(By.XPath(".//table[@id=\"tbl\"]/tbody//tr//td//input[@dataid=\"" + aid + "\"]")).Click();
            //if (ConfigUtils.Read("ENV") == "AT")
            //{
                // Console.WriteLine(ele[1].GetAttribute("innerHTML"));
                Thread.Sleep(5000);
                string sid = driver.FindElement(By.XPath(".//table/tbody/tr/td/div[contains(.,\"" + officename + "\")]/../..")).GetAttribute("id");

                // System.Diagnostics.Debug.WriteLine(sid);
                string aid = sid.Substring(1, (sid.Length - 1));
                //  driver.FindElement(By.XPath(".//table[@id=\"tbl\"]/tbody/tr/td/input[@dataid=\"" + aid + "\"]")).Click();

                Thread.Sleep(5000);
                driver.FindElement(By.XPath(".//tbody//input[@dataid='" + aid + "']")).Click();
                Thread.Sleep(2000);
                driver.FindElement(By.XPath(".//tbody//input[@dataid='" + aid + "']")).Click();
                Thread.Sleep(2000);
                driver.FindElement(By.XPath(".//tbody//input[@dataid='" + aid + "']")).Click();

                Thread.Sleep(5000);

                string sval = driver.FindElement(By.XPath(".//input[@dataid=\"" + aid + "\"]")).GetAttribute("dataid");
          //  }
           // if (ConfigUtils.Read("ENV") == "SIT")
          ////  {
                
          //      Thread.Sleep(5000);
          //      driver.FindElement(By.XPath(".//table[@id=\"contentTable\"]/tbody/tr[contains(.,\"" + officename + "\")]//input[@type=\"checkbox\"]")).Click();
               
          //  }


            // driver.FindElement(By.XPath(".//input[@dataid=\"" + aid + "\"]")).Clear();

            // System.Drawing.Point ploc = driver.FindElement(By.XPath(".//input[@dataid=\"" + aid + "\"]")).Location;
            // Mouse.Move(ploc);
            //  Mouse.Click(ploc);
            // driver.FindElement(By.XPath(".//input[@dataid=\"" + aid + "\"]")).SendKeys(OpenQA.Selenium.Keys.Enter);

        }

        public void PDMerch_EditClick_OfficeProduct(string officename)
        {
            try
            {
                driver.FindElement(By.XPath(".//table/tbody/tr[contains(.,\"" + officename + "\")]/td[1]/a")).Click();
                Reporter.ReportEvent("Click on Porduct Link for Merchent", "PD Merch Click on Porduct Link for Merchent office " + officename, HP.LFT.Report.Status.Passed);
            }
            catch(Exception e)
            {
                Reporter.ReportEvent("Click on Porduct Link for Merchent", " Exception message "+e.Message+ " in PD Merch Click on Porduct Link for Merchent office " + officename, HP.LFT.Report.Status.Failed);
            }
           
        }
        public void PDMerch_EditClick_BuyerProduct (string productname)
        {
            try
            {
                driver.FindElement(By.XPath(".//table[@id=\"tbr\"]//a[text()=\""+ productname+"\"]")).Click();
                Reporter.ReportEvent("Click on Porduct Link for Buyer", "PD Buyer Click on Porduct Link for Buyer " + productname, HP.LFT.Report.Status.Passed);
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("Click on Porduct Link for Merchent", " Exception message " + e.Message + " in PD Buyer Click on Porduct Link for Product " + productname, HP.LFT.Report.Status.Failed);
            }

        }

        public void PDMerch_SearchSuppliersByAttribute(string AttributeType)
        {
            try
            {
                driver.FindElement(By.XPath(".//div[@class=\"facet-title\"]/label[contains(.,\""+ AttributeType+"\")]")).Click();
                Reporter.ReportEvent("Click on Search Supplier Attribute", "PD Merch Click on Porduct Attribute for Search Supplier " + AttributeType, HP.LFT.Report.Status.Passed);
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("Click on Search Supplier Attribute", " Exception message " + e.Message + " in PD Merch Click on Porduct Attribute for Search Supplier " + AttributeType, HP.LFT.Report.Status.Failed);
            }

        }
        public void PDMerch_AddSupplierByProductTypeSearch(string ProductTypeOption)
        {
            try
            {
                driver.FindElement(By.XPath(".//*[@id='select_PRODUCTION_TYPE']/option[contains(.,\"" + ProductTypeOption + "\")]")).Click();
                Reporter.ReportEvent("Click on AddSupplier Seach Option", "PD Merch Click on Porduct Attribute for Search Supplier options " + ProductTypeOption, HP.LFT.Report.Status.Passed);
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("Click on AddSupplier Seach Option", " Exception message " + e.Message + " in PD Merch Click on Porduct Attribute for Search Supplier options " + ProductTypeOption, HP.LFT.Report.Status.Failed);
            }

        }

        public void PDMerch_Options_ProductGroup_Click(string ProductGroupname)
        {
            try
            {
                driver.FindElement(By.XPath(".//table[@id=\"tbl\"]/tbody//tr/td/a[contains(.,\""+ ProductGroupname+"\")]")).Click();
                Reporter.ReportEvent("Click on Options Group Name", "PD Merch Click on Options GroupName " + ProductGroupname, HP.LFT.Report.Status.Passed);
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("Click on Options Group Name", " Exception message " + e.Message + " in PD Merch Click on Options GroupName " + ProductGroupname, HP.LFT.Report.Status.Failed);
            }

        }

        public void PDMerch_Options_SetSecureStatus()
        {
            try
            {
                PDMerch_ActionOption_OptionDetails_SetStatus.Click();

               
                GeneralMethods smethod = new GeneralMethods();
                driver.SwitchTo().DefaultContent();
                smethod.WaitFrametoSwitch(driver, "overlayContentFrame");
               // driver.SwitchTo().Frame("detailsDisplay");

               // driver.SwitchTo().Frame("overlayContentFrame");
                smethod.ClickRadioBtnwithValue(driver.FindElements(By.XPath(".//*[@id='stateId']")), "Secured");
                driver.FindElement(By.XPath(".//*[@id='.tvc.btn.apply.a']/a")).Click();
              
            }
            catch (Exception e)
            {
                Reporter.ReportEvent("Set Secure status failed in Option details Page", " Exception message " + e.Message + " Set Secure status failed in Option details Page", HP.LFT.Report.Status.Failed);
            }

        }

        public string CapcityBooking_Verifyrecord(string ProductDevelopName)
        {
            try
            {
                if (driver.FindElement(By.XPath("(.//table[@class=\"contentTable\"]//td[contains(.,\""+ProductDevelopName+"\")])[1]")).Displayed == true)
                {
                    Reporter.ReportEvent("Capacity booking record displyed", "Capacity booking record displyed", HP.LFT.Report.Status.Passed);
                   return driver.FindElement(By.XPath(".//table[@class=\"contentTable\"]/tbody/tr[1]/td[1]")).Text;
                }
                return driver.FindElement(By.XPath(".//table[@class=\"contentTable\"]/tbody/tr[1]/td[1]")).Text;

            }
            catch
            {
                Reporter.ReportEvent("Capacity booking record is not displyed", "Capacity booking record is not displyed", HP.LFT.Report.Status.Failed);
                return "";
            }
        }
        public void PDMerch_Options_LabelSelect()
        {
            try
            {

                GeneralMethods smethod = new GeneralMethods();
                waitForSpinner();
                smethod.WebButton_Click(LabelType);
                Thread.Sleep(1000);
                smethod.WebEdit_SetValue(LabelTypeTextbox, "Price Tag");
                Thread.Sleep(1000);
                smethod.WebButton_Click(PriceTag, "Price Tag");
                Thread.Sleep(1000);
                smethod.WebButton_Click(done, "Done");
                Thread.Sleep(1000);
                waitForSpinner();

            }

            catch (Exception e)
            {

            }
        }

        public void addSupplierDetails(string productiongroup,string productiontype)
        {
            try
            {

                GeneralMethods smethod = new GeneralMethods();
                waitForSpinner();

                //Product Group Selection
                try
                {
                    var productionGroup = driver.FindElement(By.XPath(".//label[@filter_name='PRODUCTION_GROUP']"));
                    smethod.WebButton_Click(productionGroup);
                    var ProductionGroupTextBox = driver.FindElement(By.XPath(".//*[@id='input_PRODUCTION_GROUP']"));
                    smethod.WebEdit_SetValue(ProductionGroupTextBox, productiongroup);
                    var Woven = driver.FindElement(By.XPath(".//option[@value='" + productiongroup + "']"));
                    smethod.WebButton_Click(Woven, productiongroup);
                    Thread.Sleep(1000);
                    smethod.WebButton_Click(done, "Done");

                    waitForSpinner();
                    Thread.Sleep(2000);
                }
                catch
                {

                }
                Thread.Sleep(1000);
                try
                {
                    var productionType = driver.FindElement(By.XPath(".//label[@filter_name='PRODUCTION_TYPE']"));
                    smethod.WebButton_Click(productionType);
                    var ProductionTypeTextBox = driver.FindElement(By.XPath(".//*[@id='input_PRODUCTION_TYPE']"));
                    smethod.WebEdit_SetValue(ProductionTypeTextBox, productiontype);
                    var Dresses = driver.FindElement(By.XPath(".//option[@value='" + productiontype + "']"));
                    smethod.WebButton_Click(Dresses, productiontype);
                    Thread.Sleep(1000);
                    smethod.WebButton_Click(done, "Done");
                    waitForSpinner();
                    Thread.Sleep(2000);
                }
                catch
                {

                }
                Thread.Sleep(1000);
                try
                {
                    var productioncountry = driver.FindElement(By.XPath(".//label[@filter_name='REL_COUNTRY']"));
                    smethod.WebButton_Click(productioncountry);
                    Thread.Sleep(1000);
                    var CountryName = driver.FindElement(By.XPath(".//*[@id='input_REL_COUNTRY']"));
                    smethod.WebEdit_SetValue(CountryName, "China");
                   // var Dresses = driver.FindElement(By.XPath(".//option[@value='" + productiontype + "']"));
                   // smethod.WebButton_Click(Dresses, productiontype);
                    Thread.Sleep(1000);
                    smethod.WebButton_Click(done, "Done");
                    waitForSpinner();
                    Thread.Sleep(2000);
                }
                catch
                {

                }
                //Product type Selection
                Thread.Sleep(1000);
                //SRM Grading
                var SRMGrading = driver.FindElement(By.XPath(".//label[@filter_name='SRM_GRADING']"));
                smethod.WebButton_Click(SRMGrading);
                var SRMGradingTextBox = driver.FindElement(By.XPath(".//*[@id='input_SRM_GRADING']"));
                try {
                    var Gold = driver.FindElement(By.XPath(".//option[@value='Gold']"));
                    smethod.WebButton_Click(Gold, "Gold");
                    Thread.Sleep(1000);
                }
                catch
                {

                }
                //try
                //{
                //    var Gold1 = driver.FindElement(By.XPath(".//option[@value='Silver']"));
                //    smethod.WebButton_Click(Gold1, "Silver");
                //    Thread.Sleep(1000);
                //}
                //catch
                //{

                //}
               
                smethod.WebButton_Click(done, "Done");
                waitForSpinner();

            }

            catch (Exception e)
            {

            }
        }
        public void waitForSpinner()
        {
            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(MAXWAITTIME));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath(".//*[@id='imgProgressDiv']")));
            }
            catch
            {

            }
        }

        public void waitforDataLoadSpinner()
        {
            try
            {
                driver.SwitchTo().DefaultContent();
                driver.SwitchTo().Frame("structure_browser");
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(MAXWAITTIME));
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='uip-loader-spinner']")));
            }
            catch
            {

            }

        }


        public void addMchartDetails(string season,string producttype)
        {
            try
            {

                GeneralMethods smethod = new GeneralMethods();
                waitForSpinner();

                //Product Group Selection
                var sSeason = driver.FindElement(By.XPath(".//label[@filter_name='MEASUREMENT_CHART_SEASON']"));
                smethod.WebButton_Click(sSeason);
                var SSeasoninput= driver.FindElement(By.XPath(".//*[@id='input_MEASUREMENT_CHART_SEASON']"));
                smethod.WebEdit_SetValue(SSeasoninput, season);
                var SSeason1 = driver.FindElement(By.XPath(".//option[@value='" + season + "']"));
                smethod.WebButton_Click(SSeason1, season);
                Thread.Sleep(1000);
                smethod.WebButton_Click(done, "Done");
                waitForSpinner();
                Thread.Sleep(1000);


                var sproducttype = driver.FindElement(By.XPath(".//label[@filter_name='MEASUREMENT_CHART_PRODUCT_TYPE']"));
                smethod.WebButton_Click(sproducttype);
                Thread.Sleep(1000);
                var Sproducttypeinput = driver.FindElement(By.XPath(".//*[@id='input_MEASUREMENT_CHART_PRODUCT_TYPE']"));
                smethod.WebEdit_SetValue(Sproducttypeinput, producttype);
                Thread.Sleep(1000);
                var SProducttype1 = driver.FindElement(By.XPath(".//option[@value='" + producttype + "']"));
                smethod.WebButton_Click(SProducttype1, producttype);
                Thread.Sleep(1000);
                smethod.WebButton_Click(done, "Done");
                waitForSpinner();
                Thread.Sleep(1000);

            }
            catch
            {

            }
        }

    }
}

