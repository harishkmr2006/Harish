using BOPO.NUnit.ParallelTests.Tests.TestRunup;
using NUnit.Framework;
using BOPO.NUnit.ParallelTests;


namespace BOPO.NUnit.ParallelTests.Tests.TAGS
{

    

    [TestFixture]

    class OperationalRoutineWebApplications : TestBase
    {

        [OneTimeTearDown]
        public static void createExcel456789()
        {
            VersionConrol.convertDicToDataTable();
        }

        [Test, Category("FFRegression1"), Order(0)]
        public void TC01_BOSearch()
        {
            TC01_BOSearch_FFBrowser Run = new TC01_BOSearch_FFBrowser();
            bool finalResult = Run.TC01_BOSearch_FF();
            Assert.IsTrue(finalResult);

        }

        [Test, Category("FFRegression1"), Order(1)]
        public void TC02_SizeCurveTool_IEBrowser()
        {
            TC02_SizeCurveTool_IEBrowser Run = new TC02_SizeCurveTool_IEBrowser();
            bool finalResult = Run.TC02_SizeCurveTool_IE();
            Assert.IsTrue(finalResult);

        }


        [Test, Category("FFRegression1"), Order(2)]
        public void TC03_SellPriceTool_IEBrowser()
        {
            TC03_SellPriceTool_IEBrowser Run = new TC03_SellPriceTool_IEBrowser();
            bool finalResult = Run.TC03_SellPriceTool_IE();
            Assert.IsTrue(finalResult);
            

        }

        [Test, Category("FFRegression1"), Order(3)]
        public void TC04_ProductPlan_IEBrowsers()
        {
            TC04_ProductPlan_IEBrowsers Run = new TC04_ProductPlan_IEBrowsers();
            bool finalResult = Run.TC04_ProductPlan_IEBrowser();
            Assert.IsTrue(finalResult);

        }

        [Test, Category("FFRegression1"), Order(4)]
        public void TC05_AllocatedAssortmentTool_FFBrowser()
        {
            TC05_AllocatedAssortmentTool_FFBrowser Run = new TC05_AllocatedAssortmentTool_FFBrowser();
            bool finalResult = Run.TC05_AllocatedAssortmentTool_FF();
            Assert.IsTrue(finalResult);

        }

        [Test, Category("FFRegression1"), Order(5)]
        public void TC06_MarketOptimizationAdminTool_FFBrowser()
        {
            TC06_MarketOptimizationAdminTool_FFBrowser Run = new TC06_MarketOptimizationAdminTool_FFBrowser();
            bool finalResult = Run.TC06_MarketOptimizationAdminTool_FF();
            Assert.IsTrue(finalResult);

        }
        [Test, Category("FFRegression1"), Order(6)]
        public void TC07_PercentageAdminTool_FFBrowser()
        {
            TC07_PercentageAdminTool_FFBrowser Run = new TC07_PercentageAdminTool_FFBrowser();
            bool finalResult = Run.TC07_PercentAdminToolVerificaton();
            Assert.IsTrue(finalResult);

        }

        [Test, Category("FFRegression1"), Order(70)]
        public void TC20_VersionControlInfoTest()
        {
            TC20_VersionInformation version = new TC20_VersionInformation();
            version.TC20_VersionInformationTest();
        }


    }
}
