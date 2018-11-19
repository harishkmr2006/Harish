using System;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.Verifications;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System.Threading;
using System.Linq;
using System.Collections.Generic;
using OpenQA.Selenium.Support.UI;
using System.Configuration;
using InteropLibrary;
using OpenQA.Selenium.Interactions;
using HP.LFT.Report;

namespace SITSmokeTests
{
   
    [TestFixture]
    public class TestRunner : UnitTestClassBase
    {
        private Dictionary<string, string> testResult;

        public TestRunner()
        {
            int i;
            testResult = new Dictionary<string, string>();

            try
            {
                string sprojectpath = this.GetType().Assembly.Location;
                string[] swrokingdirectorypath = sprojectpath.Split('\\');
                string sdatafiledirectory = "";
                foreach (string sitem in swrokingdirectorypath)
                {
                    if (sitem == "SITSmokeTests")
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


        [Test]

        public void TC01_Selenium_Castor_ProductDevelopmentandPublish()
        {
            LeanFtTest1 Testpahse1 = new LeanFtTest1(); ;

            try
            {
                Testpahse1.TC01_Selenium_Castor_ProductDevelopmentandPublish();
            }
            finally
            {
                Testpahse1.TearDown();
            }
        }
        [Test]

        public void TC02_Selenium_ICCVerificationexportsTillPDPublish()
        {
         if (testResult["TC01_Selenium_Castor_ProductDevelopmentandPublish"] == "Passed")
            {

                LeanFtTest1 Testpahse1 = new LeanFtTest1();

                try
                {
                    Testpahse1.TC02_Selenium_ICCVerificationexportsTillPDPublish();
                }
                finally
                {
                    Testpahse1.TearDown();
                }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC01_Selenium_Castor_ProductDevelopmentandPublish and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

       [Test]

        public void TC03_Selenium_Castor_Merch_PublishToTillHandover()
        {
      if (testResult["TC01_Selenium_Castor_ProductDevelopmentandPublish"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();
          
            try
            {
                Testpahse1.TC03_Selenium_Castor_Merch_PublishToTillHandover();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC01_Selenium_Castor_ProductDevelopmentandPublish and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

        [Test]

        public void TC04_Selenium_ICCExportsverificationforHandOver()
        {
          if (testResult["TC03_Selenium_Castor_Merch_PublishToTillHandover"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();
           
            try
            {
                Testpahse1.TC04_Selenium_ICCExportsverificationforHandOver();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC03_Selenium_Castor_Merch_PublishToTillHandover and current test depends on that", Status.Warning);
                Assert.Ignore();
              

            }
        }

        [Test]

        public void TC05_Selenium_Castor_Buyer_PDRTO()
        {
       if (testResult["TC03_Selenium_Castor_Merch_PublishToTillHandover"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();
          
            try
            {
                Testpahse1.TC05_Selenium_Castor_Buyer_PDRTO();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC03_Selenium_Castor_Merch_PublishToTillHandover and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

        [Test]

        public void TC06_Selenium_ICCExportsverificationforRTO()
        {
           if (testResult["TC05_Selenium_Castor_Buyer_PDRTO"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();
         
            try
            {
                Testpahse1.TC06_Selenium_ICCExportsverificationforRTO();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC05_Selenium_Castor_Buyer_PDRTO and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

       [Test]

        public void TC07_Selenium_SmokeTest_ProductPlan()
        {
            RebuildTest Testpahse1 = new RebuildTest();
          
            try
            {
                Testpahse1.TC07_Selenium_SmokeTest_ProductPlan();
            }
            finally
            {
                Testpahse1.TearDown();
            }
        }
        [Test]

        public void TC08_HMORDER_NEWOrderCeation()
        {
       if ((testResult["TC05_Selenium_Castor_Buyer_PDRTO"] == "Passed")&&(testResult["TC07_Selenium_SmokeTest_ProductPlan"] == "Passed"))
            {
                LeanFtTest2 Testpahse1 = new LeanFtTest2();
            

            try
            {
                Testpahse1.TC08_HMORDER_NEWOrderCeation();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in tests TC05_Selenium_Castor_Buyer_PDRTO and TC07_Selenium_SmokeTest_ProductPlan and current test depends on these", Status.Warning);
                Assert.Ignore();

            }
        }

        [Test]

        public void TC09_Selenium_PPlan_EditQunatity_Colorvalidation()
        {
        if (testResult["TC08_HMORDER_NEWOrderCeation"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();

            try
            {
                Testpahse1.TC09_Selenium_PPlan_EditQunatity_Colorvalidation();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC08_HMORDER_NEWOrderCeation and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC10_HMORDER_Definite()
        {
      if (testResult["TC08_HMORDER_NEWOrderCeation"] == "Passed")
            {
                LeanFtTest2 Testpahse1 = new LeanFtTest2();


            try
            {
                Testpahse1.TC10_HMORDER_Definite();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC08_HMORDER_NEWOrderCeation and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

        [Test]

        public void TC11_Selenium_PPlan_Definite_Colorvalidation()
        {
        if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();

            try
            {
                Testpahse1.TC11_Selenium_PPlan_Definite_Colorvalidation();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

       //// [Test]

        public void TC12_HMOrder_DBExports_Verification()
        {
            if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest2 Testpahse1 = new LeanFtTest2();


            try
            {
                Testpahse1.TC12_HMOrder_DBExports_Verification();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC13_Selenium_SmokeTest_HMORder_BOPOInternalOrderexportsVerification()
        {
      if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC13_Selenium_SmokeTest_HMORder_BOPOInternalOrderexportsVerification();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

        [Test]

        public void TC14_Selenium_SmokeTest_HMORder_GenericOrderExports_Verification()
        {
           if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC14_Selenium_SmokeTest_HMORder_GenericOrderExports_Verification();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

        [Test]

        public void TC15_Selenium_SmokeTest_HMORder_OrderedProductSpecification_Verification()
        {
          if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC15_Selenium_SmokeTest_HMORder_OrderedProductSpecification_Verification();
            }
            finally
            {
                Testpahse1.TearDown();
            }
        }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

        [Test]

        public void TC16_Selenium_Castor_CapacityBooking()
        {
           if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC16_Selenium_Castor_CapacityBooking();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

        [Test]

        public void TC17_Selenium_SmokeTest_CapacityBooking_CastorExports_Verification()
        {
          if (testResult["TC16_Selenium_Castor_CapacityBooking"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC17_Selenium_SmokeTest_CapacityBooking_CastorExports_Verification();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC16_Selenium_Castor_CapacityBooking and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC18_Selenium_Castor_SampleOrder_CounterSample_Creation()
        {
          if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC18_Selenium_Castor_SampleOrder_CounterSample_Creation();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC19_Selenium_Castor_SampleOrder_ProductionSample_Creation()
        {
           if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC19_Selenium_Castor_SampleOrder_ProductionSample_Creation();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC20_Selenium_Castor_SampleOrder_RegisterSample()
        {
      if ((testResult["TC18_Selenium_Castor_SampleOrder_CounterSample_Creation"] == "Passed")|| (testResult["TC19_Selenium_Castor_SampleOrder_ProductionSample_Creation"] == "Passed"))
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC20_Selenium_Castor_SampleOrder_RegisterSample();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC18_Selenium_Castor_SampleOrder_CounterSample_Creation or TC19_Selenium_Castor_SampleOrder_ProductionSample_Creation and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC21_Selenium_Castor_SampleOrder_LabTest()
        {
          if (testResult["TC20_Selenium_Castor_SampleOrder_RegisterSample"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC21_Selenium_Castor_SampleOrder_LabTest();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC20_Selenium_Castor_SampleOrder_RegisterSample and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC22_Selenium_Castor_SampleOrder_SampleReportgeneration()
        {
         if (testResult["TC21_Selenium_Castor_SampleOrder_LabTest"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC22_Selenium_Castor_SampleOrder_SampleReportgeneration();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC21_Selenium_Castor_SampleOrder_LabTest and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC23_Selenium_Castor_AssignKoreaCertificate()
        {
            if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC23_Selenium_Castor_AssignKoreaCertificate();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC24_Selenium_ICCExports_CountryCertificates()
        {
            if (testResult["TC23_Selenium_Castor_AssignKoreaCertificate"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC24_Selenium_ICCExports_CountryCertificates();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC23_Selenium_Castor_AssignKoreaCertificate and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC25_SIT_Library_ConsumerPackaging()
        {
            SIT_LibraryTest Testpahse1 = new SIT_LibraryTest();


            try
            {
                Testpahse1.TC25_SIT_Library_ConsumerPackaging();
            }
            finally
            {
                Testpahse1.TearDown();
            }
        }
        [Test]

        public void TC26_SIT_Library_License_Companies()
        {
            SIT_LibraryTest Testpahse1 = new SIT_LibraryTest();


            try
            {
                Testpahse1.TC26_SIT_Library_License_Companies();
            }
            finally
            {
                Testpahse1.TearDown();
            }
        }

        [Test]

        public void TC27_SIT_Library_Materials()
        {
            SIT_LibraryTest Testpahse1 = new SIT_LibraryTest();


            try
            {
                Testpahse1.TC27_SIT_Library_Materials();
            }
            finally
            {
                Testpahse1.TearDown();
            }
        }
       [Test]

        public void TC28_SIT_Library_TransportPacking()
        {
            SIT_LibraryTest Testpahse1 = new SIT_LibraryTest();


            try
            {
                Testpahse1.TC28_SIT_Library_TransportPacking();
            }
            finally
            {
                Testpahse1.TearDown();
            }
        }

        [Test]

        public void TC23_2_Selenium_Castor_Merch_MChart()
        {
           if (testResult["TC10_HMORDER_Definite"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


            try
            {
                Testpahse1.TC23_2_Selenium_Castor_Merch_MChart();
            }
            finally
            {
                Testpahse1.TearDown();
            }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }
        [Test]

        public void TC23_2_Selenium_ICCExports_Merch_MChart()
        {
           if (testResult["TC23_2_Selenium_Castor_Merch_MChart"] == "Passed")
            {
                LeanFtTest1 Testpahse1 = new LeanFtTest1();


                try
                {
                    Testpahse1.TC23_2_Selenium_SmokeTest_MchartICCexportsVerification();
                }
                finally
                {
                    Testpahse1.TearDown();
                }
            }
            else
            {

                Reporter.ReportEvent(TestContext.CurrentContext.Test.MethodName, "due to fail/Skip in test TC10_HMORDER_Definite and current test depends on that", Status.Warning);
                Assert.Ignore();

            }
        }

        [Test]
        public void TC23_3_Selenium_VerifyCITStatus()
        {
            LeanFtTest1 TestPhase = new LeanFtTest1();

            try
            {
                TestPhase.TC23_3_Selenium_VerifyCITStatus();
            }
            finally
            {
                TestPhase.TearDown();
            }
        }

       [Test]
        public void TC23_4_Selenium_VerifyTagsPTAndITPush()
        {
            LeanFtTest1 TestPhase = new LeanFtTest1();

            try
            {
                TestPhase.TC23_4_Selenium_VerifyTagsPTAndITPush();
            }
            finally
            {
                TestPhase.TearDown();
            }
        }


      //  [Test]
        public void TC29_Selenium_OFULogin()
        {
            LeanFtTest1 TestPhase = new LeanFtTest1();

            try
            {
                TestPhase.TC29_Selenium_OFULogin();
            }
            finally
            {
                TestPhase.TearDown();
            }
        }

        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
            string testName = TestContext.CurrentContext.Test.MethodName;
            string Result = TestContext.CurrentContext.Result.Outcome.ToString();
            testResult.Add(testName, Result);


        }

        //[OneTimeTearDown]
        //public void TestFixtureTearDown()
        //{
        //    //try
        //    //{
        //    //    driver.Quit();
        //    //}
        //    //catch
        //    //{

        //    //}
        //    //// Clean up once per fixture
        //    //// Clean up after each test
        //    //xCellFileHelper.CleanUp();
        //    ////            driver.Close();
        //    //Thread.Sleep(2000);
        //}
    }
}
