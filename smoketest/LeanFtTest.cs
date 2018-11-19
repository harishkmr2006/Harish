using System;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.SDK.Web;
using HP.LFT.Verifications;
using System.Threading;

namespace SITSmokeTests
{
    [TestFixture]
    public class LeanFtTest : UnitTestClassBase
    {
        [OneTimeSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture
        }

        [SetUp]
        public void SetUp()
        {
            // Before each test
        }

     //   [Test]
        public void Test()
        {
            // IBrowser browser;
            // CastorObjects oCastor = new CastorObjects();
            // IBrowser Browserobj= BrowserFactory.Launch(BrowserType.InternetExplorer);
            // Thread.Sleep(2000);
            // Browserobj.Navigate("http://castor-sit.hm.com");
            //// browser.Describe<IPage>(new PageDescription());

            // oCastor.MessageFromWebpageDialog.OKButton.Click();
            // //  Browserobj.Sync();
            // // Browserobj.Page.

            // //var pages = Browserobj.FindChildren<IPage>(new PageDescription());
            // //pages[0].Highlight();
            // //pages[0].Describe<IButton>(new ButtonDescription
            // //{
            // //    ButtonType = @"button",
            // //    TagName = @"INPUT",
            // //    Name = @"SEBO PD Buyer"
            // //}).Click();
            //oCastor.Castor.SEBOPDBuyerButton.Click();
            // Browserobj.Sync();
            // oCastor.CastorPage2.ContentFrame.ProductDevelopmentsLink.Click();

            // oCastor.CastorPage2.TvcTabs0HMProductDevelopmentBuyerStartPageTabcontentFrame.ProductDevelopmentsLink.Click();
            // oCastor.CastorPage2.TableContentFrame.CreateProductDevelopmentWebElement.Click();
            //string ssuppliername = "BEIJING TOPNEW IMP.&EXP.CO,LTD";
            //if (ssuppliername.Length > 20)
            //{
            //    string ssup1 = ssuppliername.Substring(1, 19);
            //    Thread.Sleep(1000);
            //    }

            string sadate = "2017-08-04 11:43:52";

           string sadate2= sadate.Substring(0, 16);
            DateTime ddate = DateTime.Parse(sadate2);
            DateTime ddate2= ddate.AddMinutes(1);


            Thread.Sleep(1000);


        }

        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
        }

        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
        }
    }
}
