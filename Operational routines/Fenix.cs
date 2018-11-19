using System;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.Verifications;
using System.Diagnostics;
using FenixAppModel;
using HP.LFT.SDK.StdWin;
using BOPO.NUnit.ParallelTests;

namespace Auto_Fenix
{
    [TestFixture]
    public class Fenix : UnitTestClassBase
    {
        ApplicationModel1 app = new ApplicationModel1();
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
        //Select th Department
        public void selectDepartment(string vDepartmentString)
        {
            try
            {
               // int offset = 0;
                string[] nodetextpath = vDepartmentString.Split(';');
                for (int i = 0; i <= nodetextpath.GetUpperBound(0); i++)
                {
                    if (i == nodetextpath.GetUpperBound(0))
                    {
                        var selectProductGrid = app.FenixWindow.SelectionTreeView;
                        //app.FenixWindow.SelectionTreeView.Click(nodetextpath.GetUpperBound(0));
                        //app.FenixWindow.SelectionTreeView.Select(1);
                        //app.FenixWindow.SelectionTreeView.Click();




                        selectProductGrid.Describe<ICheckBox>(new CheckBoxDescription
                        {
                            // ObjectName = @"check box",
                            Text = "Ladies Everyday",
                            NativeClass = @"System.Windows.Controls.CheckBox",


                            //  Text = @"Ladies Everyday",
                            // ObjectName = @"Ladies Everyday"
                        }).Click();


                        //var t = app.FenixWindow.SelectionTreeView.NativeObject.MEMBERS;
                        //Console.WriteLine(t);
                        //Desktop.Describe<IWindow>(new WindowDescription
                        //{
                        //    FullType = @"window"
                        //}).Describe<IUiObject>(new UiObjectDescription
                        //{
                        //    ObjectName = @"ContentControl"
                        //}).Describe<IUiObject>(new UiObjectDescription
                        //{
                        //    FullType = @"object",
                        //    Index = 0
                        //}).Describe<IButton>(new ButtonDescription
                        //{
                        //    Text = @"Load Products",
                        //    ObjectName = @"LoadProductsButton"
                        //});
                        //Desktop.Describe<IWindow>(new WindowDescription
                        //{
                        //    NativeClass = @"System.Windows.Controls.Primitives.PopupRoot",
                        //    Index = 0
                        //}).Click();


                    }
                 
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
       // [Test]
        public void FenixFlow()
        {
            VersionConrol._initTestData(TestContext.CurrentContext.Test.MethodName);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.StartTime.ToString(), VersionConrol.getTimeStamp());
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.ApplictionName.ToString(), "Fenix");
            string apppath = "C:\\Program Files (x86)\\H & M Hennes & Mauritz AB\\H & M Fenix Client\\HM.Plan.Fenix.exe";
            var appname = FileVersionInfo.GetVersionInfo(apppath);
            var appver = appname.FileVersion;
            VersionConrol.addApplication("Fenix", appver);
            VersionConrol.addSubKey(TestContext.CurrentContext.Test.MethodName, subKeys.keySet.Version.ToString(), appver);

            //launch application
            Process process = Process.Start(apppath);
            if(app.ReadOnlySeasonWindow.Exists())
            {
                
                app.ReadOnlySeasonWindow.OkButton.Click();
            }

            

            app.FenixWindow.Highlight();
            //var AT = app.FenixWindow.GetTextLocations("Assortment Triangle");
            //int x = Convert.ToInt32((AT[0].Width) / 2.0 ) + app.FenixWindow.Location.X;

            //int y = Convert.ToInt32(AT[0].Height / 2.0 ) + app.FenixWindow.Location.Y;
            //Mouse.Click(new System.Drawing.Point(x, y));
            app.FenixWindow.ForwardSeasonComboBox.Click();
            app.FenixWindow.ForwardSeasonComboBox.Select(3);
            //app.FenixWindow.Div02PreEarlyCollection1926Button.Click();
            //var treeviewtext = app.FenixWindow.SelectionTreeView.GetVisibleText();
            //string depttoselect = "H&M.;Div 03 Ladies Everyday;Ladies Everyday";
            //selectDepartment(depttoselect);
            app.FenixWindow.Close();
        }

       
        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
            app.FenixWindow.Close();
        }

        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
        }
    }
}
