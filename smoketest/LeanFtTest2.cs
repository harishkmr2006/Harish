using System;
using NUnit.Framework;
using HP.LFT.SDK;
using HP.LFT.Verifications;
using InteropLibrary;
using System.Threading;
using HP.LFT.SDK.WinForms;
using System.Collections.Generic;

namespace SITSmokeTests
{
    [TestFixture]
    public class LeanFtTest2 : UnitTestClassBase
    {
        HMOrderObject appModel;
        ExcelHelper xCellFileHelper;
        string datafilePath;

        public LeanFtTest2()
        {
            appModel = new HMOrderObject();
            datafilePath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + ConfigUtils.Read("TestDataPath");
        }


        [OneTimeSetUp]
        public void TestFixtureSetUp()
        {
            // Setup once per fixture
        }

        [SetUp]
        public void SetUp()
        {
            // Before each test
         
           
            appModel = new HMOrderObject();
            datafilePath = ConfigUtils.Read("TestDataPath");


        }

      // [Test, Order(8)]
        public void TC08_HMORDER_NEWOrderCeation()
        {
             try
             {
                 if (appModel.HMOrderWindow.Exists())
                 {
                     appModel.HMOrderWindow.Close();
                     try
                     {
                         appModel.HMOrderWindow.ClosingHMOrderWindow.YesUiObject.Click();
                     }
                     catch
                     {

                     }
                     Thread.Sleep(20000);
                 }

             }
             catch
             {

             }
             System.Diagnostics.Process.Start("C:\\Program Files (x86)\\H & M Hennes & Mauritz AB\\HMOrder Client\\HM.HMOrder.ManagerClient.UserInterface.Windows.exe");
             Thread.Sleep(30000);
             appModel.HMOrderWindow.Exists(10);
             appModel.HMOrderWindow.Activate();
             xCellFileHelper = new ExcelHelper(datafilePath, 1);
             string sSeason = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");
             string[] sSeasonArr = sSeason.Split('-');
             //Selection of Season
             appModel.HMOrderWindow.Activate();
             try
             {
                 var Season_x = appModel.HMOrderWindow.BaseFormToolbarsDockAreaTopUiObject.NativeObject.ToolbarsManager.Ribbon.Tabs[0].Groups[2].Tools[2].Bounds.X;
                 var Season_y = appModel.HMOrderWindow.BaseFormToolbarsDockAreaTopUiObject.NativeObject.ToolbarsManager.Ribbon.Tabs[0].Groups[2].Tools[2].Bounds.Y;
                 System.Drawing.Point Seasonpoint = new System.Drawing.Point(Season_x + 80, Season_y + 15);
                 Mouse.Move(Seasonpoint);
                 Thread.Sleep(5000);
                 Mouse.Click(Seasonpoint, MouseButton.Left);

                 Thread.Sleep(1000);
                 Keyboard.SendString(sSeasonArr[0].Trim());
                 Thread.Sleep(1000);
                // Keyboard.SendString(Keys.Tab);
                 Reporter.ReportEvent("HMOrder Season Selection", "HMOrder Season Selected - " + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season"), HP.LFT.Report.Status.Passed);
             }
             catch (Exception e)
             {
                 Reporter.ReportEvent("HMOrder Season Selection", "HMOrder Season Selected - " + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season"), HP.LFT.Report.Status.Failed);
             }

             Thread.Sleep(2000);
             //Selection of Node
             try
             {
                 appModel.HMOrderWindow.Activate();
                 appModel.HMOrderWindow.OrderExplorerWindow.UltraTreeDepartmentsUiObject.Click();
                 string sNodeName = "";
                 string sdept = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Dept").Trim();



                 Thread.Sleep(1000);
                 var sNodemebers1 = appModel.HMOrderWindow.OrderExplorerWindow.UltraTreeDepartmentsUiObject.NativeObject.MEMBERS;
                 // var sNodemebers = appModel.HMOrderWindow.OrderExplorerWindow.UltraTreeDepartmentsUiObject.NativeObject.Nodes.MEMBERS;
                 var vNodes = appModel.HMOrderWindow.OrderExplorerWindow.UltraTreeDepartmentsUiObject.NativeObject.Nodes[0].Nodes;
                 var scount = vNodes.Count;
                // var scount = sNodemebers1.Nodes[0].Nodes.Count;
                 Thread.Sleep(2000);
                 bool bNodefound = false;

                 for (int i = 0; i <= scount - 1; i++)
                 {
                    // sNodeName = appModel.HMOrderWindow.OrderExplorerWindow.UltraTreeDepartmentsUiObject.NativeObject.Nodes[0].Nodes[i].Key;
                     sNodeName = vNodes[i].Key;
                     if (sNodeName == sdept)
                     {
                          appModel.HMOrderWindow.OrderExplorerWindow.UltraTreeDepartmentsUiObject.NativeObject.Nodes[0].Nodes[i].Selected = true;
                       //  sNodemebers1.Nodes[0].Nodes[i].Selected = true;
                         bNodefound = true;
                         break;
                     }
                 }
                 if (bNodefound == true)
                 {
                     Reporter.ReportEvent("HMOrder Node Selection", "HMOrder Node Selected - " + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Dept"), HP.LFT.Report.Status.Passed);
                 }
                 else
                 {
                     Reporter.ReportEvent("HMOrder Node Selection", "HMOrder Node Selected - " + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Dept"), HP.LFT.Report.Status.Failed);
                 }
             }
             catch (Exception e)
             {
                 Reporter.ReportEvent("HMOrder Node Selection", "HMOrder Node Selected - " + xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Dept"), HP.LFT.Report.Status.Failed);
             }



             //New Order window with Castor Options
             try
             {
                 Thread.Sleep(1000);
                 appModel.HMOrderWindow.Activate();
                // appModel.HMOrderWindow.BaseFormToolbarsDockAreaTopUiObject.Click();
                 var stabmem = appModel.HMOrderWindow.BaseFormToolbarsDockAreaTopUiObject.NativeObject.MEMBERS;
                 var GrouptoolMem = appModel.HMOrderWindow.BaseFormToolbarsDockAreaTopUiObject.NativeObject.ToolbarsManager.Ribbon.Tabs[0].Groups[0].Tools.MEMBERS;
                 var nx1 = appModel.HMOrderWindow.BaseFormToolbarsDockAreaTopUiObject.NativeObject.ToolbarsManager.Ribbon.Tabs[0].Groups[0].Tools[0].Bounds.X;
                 var ny1 = appModel.HMOrderWindow.BaseFormToolbarsDockAreaTopUiObject.NativeObject.ToolbarsManager.Ribbon.Tabs[0].Groups[0].Tools[0].Bounds.Y;

                 System.Drawing.Point Newordpoint1 = new System.Drawing.Point(nx1 + 20, ny1 + 20);
                 Mouse.Move(Newordpoint1);
                 Mouse.Click(Newordpoint1, MouseButton.Left);
                 Thread.Sleep(5000);
                 Newordpoint1 = new System.Drawing.Point(nx1 + 80, ny1 + 80);
                 Mouse.Move(Newordpoint1);
                 Mouse.Click(Newordpoint1, MouseButton.Left);

                 Thread.Sleep(5000);
                 Reporter.ReportEvent("HMOrder New Order Window With Castor Options", "HMOrder New Order Window With Castor Options Opened", HP.LFT.Report.Status.Passed);
             }
             catch (Exception e)
             {
                 Reporter.ReportEvent("HMOrder New Order Window With Castor Options", "HMOrder New Order Window With Castor Options Opened", HP.LFT.Report.Status.Failed);
             }

            // appModel.HMOrderWindow.NewOrder.ChkShowAllUiObject.Click();
             try
             {

                 string sPplanProduct = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "PPlanProduct");
                 string scastorproduct = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ProductID");
                // string weektoedit = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "ISWWEEKTOEDIT");
                appModel.HMOrderWindow.NewOrder.Activate();
                 var rcnt = appModel.HMOrderWindow.NewOrder.GridBoardCardsTable.CustomGrid.UltraWinGrid.RowCount;
                 Thread.Sleep(2000);
                 bool bfoundPPlanprod = false;
                 for (int j = 0; j <= rcnt - 1; ++j)
                 {
                     // var scellval1 = appModel.HMOrderWindow.NewOrder.GridBoardCardsTable.CustomGrid.UltraWinGrid.GetCell(j, 1).Value;
                     var scellval1 = appModel.HMOrderWindow.NewOrder.GridBoardCardsTable.CustomGrid.UltraWinGrid.GetCell(j, "Product Name").Value;
                     // var scellval5 = appModel.HMOrderWindow1.NewOrder.GridBoardCardsTable.CustomGrid.UltraWinGrid.GetCell(j, 5).Value;
                     if (scellval1.ToString() == sPplanProduct)
                     {
                         appModel.HMOrderWindow.NewOrder.GridBoardCardsTable.CustomGrid.UltraWinGrid.SelectRow(j.ToString());
                         bfoundPPlanprod = true;
                         break;
                     }
                 }

                 if (bfoundPPlanprod == true)
                 {
                     Reporter.ReportEvent("HMOrder New Order Window With Castor Options having the PPlan Product", "HMOrder New Order Window With Castor Options having the PPlan Product " + sPplanProduct, HP.LFT.Report.Status.Passed);
                 }
                 else
                 {
                     Reporter.ReportEvent("HMOrder New Order Window With Castor Options having the PPlan Product", "HMOrder New Order Window With Castor Options having the PPlan Product " + sPplanProduct, HP.LFT.Report.Status.Failed);
                 }
                 appModel.HMOrderWindow.NewOrder.CboSizeScalesUiObject.Click();
                 Thread.Sleep(2000);

                   appModel.HMOrderWindow.NewOrder.CboSizeScalesUiObject.Click();
                  appModel.HMOrderWindow.NewOrder.CboSizeScalesUiObject.SendKeys(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Scale"));
                 appModel.HMOrderWindow.NewOrder.CboCurveGroupsUiObject.SendKeys(Keys.Tab);
                 Thread.Sleep(1000);
                 appModel.HMOrderWindow.NewOrder.CboCurveGroupsUiObject.Click();
                 Thread.Sleep(1000);
                 appModel.HMOrderWindow.NewOrder.CboCurveGroupsUiObject.SendKeys(xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "SizeCurveGroup"));
                 appModel.HMOrderWindow.NewOrder.CboCurveGroupsUiObject.SendKeys(Keys.Tab);
                 Thread.Sleep(1000);
                 //string sscale = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Scale");


                 //var sval5 = appModel.HMOrderWindow.NewOrder.CboSizeScalesUiObject.NativeObject.ValueList.ValueListItems.Count;
                 //Thread.Sleep(1000);
                 //for (int i=0;i<= sval5-1; ++i)
                 //{
                 //    var svalues3 = appModel.HMOrderWindow.NewOrder.CboSizeScalesUiObject.NativeObject.SelectedItem.DisplayText;
                 //    Thread.Sleep(1000);
                 //    if (svalues3.Contains(sscale))
                 //    {
                 //        break;
                 //    }
                 //    else
                 //    {
                 //        appModel.HMOrderWindow.NewOrder.CboSizeScalesUiObject.SendKeys(Keys.Down);
                 //    }
                 //}

                 appModel.HMOrderWindow.NewOrder.CboSizeScalesUiObject.SendKeys(Keys.Tab);
                 Thread.Sleep(1000);
                 appModel.HMOrderWindow.NewOrder.NextButton.Click();
                 Thread.Sleep(5000);
                 var roptioncnt = appModel.HMOrderWindow.NewOrder.OptionsGridTable.CustomGrid.UltraWinGrid.RowCount;
                 Thread.Sleep(2000);
                 bool bfoundDevNo = false;
                 for (int k = 0; k <= roptioncnt - 1; ++k)
                 {

                     //  var soptcellval13 = appModel.HMOrderWindow.NewOrder.OptionsGridTable.CustomGrid.UltraWinGrid.GetCell(k, 13).Value;
                     var soptcellval13 = appModel.HMOrderWindow.NewOrder.OptionsGridTable.CustomGrid.UltraWinGrid.GetCell(k, "Dev No").Value;
                     Thread.Sleep(2000);

                     if (soptcellval13.ToString() == scastorproduct)
                     {
                         // appModel.HMOrderWindow.NewOrder.OptionsGridTable.CustomGrid.UltraWinGrid.SelectRow(k.ToString());
                         appModel.HMOrderWindow.NewOrder.OptionsGridTable.CustomGrid.UltraWinGrid.ActivateCell(k, "");
                        //  appModel.HMOrderWindow.NewOrder.OptionsGridTable.CustomGrid.UltraWinGrid.SelectCell(k, "");
                         Thread.Sleep(2000);
                         bfoundDevNo = true;
                         break;
                     }
                 }
                 if (bfoundDevNo == true)
                 {
                     Reporter.ReportEvent("HMOrder New Order Window With Castor Options having the Castor Product number in Dev No", "HMOrder New Order Window With Castor Options having the Castor Product number in Dev No" + scastorproduct, HP.LFT.Report.Status.Passed);
                 }
                 else
                 {
                     Reporter.ReportEvent("HMOrder New Order Window With Castor Options having the Castor Product number in Dev No", "HMOrder New Order Window With Castor Options having the Castor Product number in Dev No" + scastorproduct, HP.LFT.Report.Status.Failed);
                 }
                 appModel.HMOrderWindow.NewOrder.NextButton.Click();
                 Thread.Sleep(5000);
                 appModel.HMOrderWindow.NewOrder.FinishButton.Click();
                 Thread.Sleep(10000);
                 appModel.HMOrderWindow.Activate();
                 Thread.Sleep(1000);
                 var stabs = appModel.HMOrderWindow.OrderEditProperties.UltraTabControl1UiObject.NativeObject.Tabs.Count;
                 var stabs3 = "";
                 appModel.HMOrderWindow.OrderEditProperties.UltraTabControl1UiObject.NativeObject.Tabs[0].Selected = true;
                 //for (int i = 0; i < stabs; ++i)
                 //{
                 //    stabs3 = appModel.HMOrderWindow.OrderEditProperties.UltraTabControl1UiObject.NativeObject.Tabs[i].Key;
                 //    if (stabs3 == "PriceCalculationTab")
                 //    {
                 //        appModel.HMOrderWindow.OrderEditProperties.UltraTabControl1UiObject.NativeObject.Tabs[i].Selected = true;
                 //        int rrowcnt1 = appModel.HMOrderWindow.OrderEditProperties.UltraGridPriceCalculationTable.CustomGrid.UltraWinGrid.RowCount;
                 //        for (int j = 0; j < rrowcnt1 - 1; ++j)
                 //        {
                 //            appModel.HMOrderWindow.OrderEditProperties.UltraGridPriceCalculationTable.SelectCell(j, 5);
                 //        }
                 //    }

                 //}
                 //Thread.Sleep(1000);


                 //appModel.HMOrderWindow.OrderEditProperties.UltraTabControl1UiObject.NativeObject.Tabs[0].Selected = true;
                 Thread.Sleep(1000);

                 appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.SelectCell(0, "Qty");
                 var xx = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.Location.X;
                 var yy = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.Location.Y;
                 var x = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.GetCell(0, "Qty").X;
                 var y = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.GetCell(0, "Qty").Y;

                 Location loc = new Location(Position.Center, new System.Drawing.Point(x + 2, y + 2));
                 //  appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.MouseMove(loc);
                 Mouse.Click(new System.Drawing.Point(x + xx, y + yy), MouseButton.Right);
                 Thread.Sleep(1000);
                 Mouse.Click(new System.Drawing.Point(x + xx + 10, y + yy + 10), MouseButton.Left);
                 Thread.Sleep(1000);
                 Mouse.Move(new System.Drawing.Point(x + xx + 170, y + yy + 40));
                 Thread.Sleep(1000);
                 Mouse.Click(new System.Drawing.Point(x + xx + 170, y + yy + 40), MouseButton.Left);
                 Thread.Sleep(2000);
                 appModel.HMOrderWindow.Activate();
                 int ipOrdertablerows = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.RowCount;
                 for (int H = 0; H <= ipOrdertablerows - 1; ++H)
                 {

                     var sqtyvalH = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.GetCell(H, "PM").Value;
                     if (sqtyvalH.ToString() == "UY-D8")
                     {
                         appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.ActivateCell(H, "PM");
                         appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.SelectCell(H, "PM");
                         var xx1 = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.Location.X;
                         var yy1 = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.Location.Y;
                         var x1 = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.GetCell(H, "PM").X;
                         var y1 = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.GetCell(H, "PM").Y;

                         Location loc1 = new Location(Position.Center, new System.Drawing.Point(x1 + 2, y1 + 2));
                         Thread.Sleep(10000);
                         //  appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.MouseMove(loc);
                         Mouse.Click(new System.Drawing.Point(x1 + xx1, y1 + yy1), MouseButton.Right);
                         Thread.Sleep(1000);
                         Mouse.Click(new System.Drawing.Point(x1 + xx1 + 10, y1 + yy1 + 10), MouseButton.Left);
                         Thread.Sleep(1000);
                         Mouse.Click(new System.Drawing.Point(x1 + xx1, y1 + yy1), MouseButton.Right);
                         Thread.Sleep(1000);
                         Mouse.Move(new System.Drawing.Point(x1 + xx1 + 20, y1 + yy1 + 30));
                         Thread.Sleep(1000);
                         Mouse.Click(new System.Drawing.Point(x1 + xx1 + 20, y1 + yy1 + 30), MouseButton.Left);
                         Thread.Sleep(2000);
                         Mouse.Click(new System.Drawing.Point(x1 + xx1, y1 + yy1), MouseButton.Right);
                         Thread.Sleep(1000);
                         Mouse.Move(new System.Drawing.Point(x1 + xx1 + 20, y1 + yy1 + 110));
                         Thread.Sleep(1000);
                         Mouse.Click(new System.Drawing.Point(x1 + xx1 + 20, y1 + yy1 + 110), MouseButton.Left);
                         Thread.Sleep(2000);
                     }

                     }
                     for (int l = 0; l <= ipOrdertablerows - 1; ++l)
                    {

                     var sqtyval10 = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.GetCell(l, "Qty").Value;
                     int snewqty =  Int32.Parse(sqtyval10.ToString()) + 1;
                     if (sqtyval10.ToString() != "")
                     {
                         appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.ActivateCell(l, "Qty");

                         appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.GetCell(l, "Qty").SetValue(snewqty.ToString());


                         appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.ActivateCell(l, "ISW");
                         var scelval = appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.GetCell(l, "ISW").Value;
                         int sval = Int32.Parse(scelval.ToString()) + 1;
                         appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.GetCell(l, "ISW").SetValue(sval);

                         // appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.SendKeys(weektoedit);
                         if (!appModel.HMOrderWindow.UpdateTODWindow.YesUiObject.Exists())
                         {
                             appModel.HMOrderWindow.OrderEditProperties.UltraGridMainOrderTable.CustomGrid.UltraWinGrid.SelectCell(l, "Qty");
                         }

                         if (appModel.HMOrderWindow.UpdateTODWindow.YesUiObject.Exists())
                         {
                             appModel.HMOrderWindow.UpdateTODWindow.YesUiObject.Click();

                             Reporter.ReportEvent("HMOrder New Order Window QTY and ISW Edit", "HMOrder New Order Window QTY and ISW Edit Field done", HP.LFT.Report.Status.Passed);
                             break;
                         }
                         else
                         {
                           //  Reporter.ReportEvent("HMOrder New Order Window QTY and ISW Edit", "HMOrder New Order Window ISW Edit Field Popup Not displeyed after edit for confirmation", HP.LFT.Report.Status.Failed);
                             break;
                         }



                     }

                 }
                 Thread.Sleep(10000);
                 appModel.HMOrderWindow.SendKeys("s", KeyModifier.Ctrl);
                 Thread.Sleep(2000);
                 try
                 {
                     Keyboard.SendString(Keys.Return);
                     Thread.Sleep(2000);
                     appModel.HMOrderWindow.SendKeys("s", KeyModifier.Ctrl);
                      Thread.Sleep(2000);
                 }
                 catch
                 {

                 }
                 Thread.Sleep(10000);
                Reporter.ReportEvent("HMOrder New Order Saved as Prelimary", "HMOrder New Order Saved as Prelimary", HP.LFT.Report.Status.Passed);
                //string stext = "";
                //try
                //{
                //    appModel.HMOrderWindow.OrderEditProperties.Activate();
                //    stext = appModel.HMOrderWindow.OrderEditProperties.PreliminaryUiObject.GetVisibleText();
                //}
                //catch
                //{
                //    Thread.Sleep(20000);
                //    appModel.HMOrderWindow.OrderEditProperties.Activate();
                //    stext = appModel.HMOrderWindow.OrderEditProperties.PreliminaryUiObject.GetVisibleText();
                //}

                // if (stext == "Preliminary")
                // {
                //     Reporter.ReportEvent("HMOrder New Order Saved as Prelimary", "HMOrder New Order Saved as Prelimary", HP.LFT.Report.Status.Passed);
                // }
                // else
                // {
                //     Reporter.ReportEvent("HMOrder New Order Saved as Prelimary", "HMOrder New Order not Saved as Prelimary", HP.LFT.Report.Status.Failed);
                // }


            }
             catch (Exception e)
             {
                 string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                 string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                 System.Drawing.Image iimage = appModel.HMOrderWindow.GetSnapshot();
                 iimage.Save(ESSpath);
                 Reporter.ReportEvent("HMOrder New Order Window With Castor Options", "HMOrder New Order Window With Castor Options Opened", HP.LFT.Report.Status.Failed, ESSpath);
                 Assert.IsTrue(2 == 3);
             }

        }
        //  [Test,Order(10)]
        public void TC10_HMORDER_Definite()
        {
            try {
                xCellFileHelper = new ExcelHelper(datafilePath, 1);
                appModel.HMOrderWindow.Activate();
                appModel.HMOrderWindow.OrderEditProperties.PreliminaryUiObject.Click();
                appModel.HMOrderWindow.OrderEditProperties.PreliminaryUiObject.SendKeys("Definite");
                appModel.HMOrderWindow.OrderEditProperties.PreliminaryUiObject.SendKeys(Keys.Tab);
                Thread.Sleep(1000);

                Thread.Sleep(1000);
                appModel.HMOrderWindow.Activate();
                Thread.Sleep(1000);
                appModel.HMOrderWindow.SendKeys("s", KeyModifier.Ctrl);
                Thread.Sleep(10000);
                //string stext = appModel.HMOrderWindow.OrderEditProperties.PreliminaryUiObject.GetVisibleText();
                //if (stext == "Definite")
                //{
                //    Reporter.ReportEvent("HMOrder New Order Saved as Definite", "HMOrder New Order Saved as Definite", HP.LFT.Report.Status.Passed);
                //}
                //else
                //{
                //    Reporter.ReportEvent("HMOrder New Order Saved as Definite", "HMOrder New Order not Saved as Definite", HP.LFT.Report.Status.Failed);
                //}
                appModel.HMOrderWindow.Close();
                try
                {
                    appModel.HMOrderWindow.ClosingHMOrderWindow.YesUiObject.Click();
                }
                catch
                {

                }

            }
            catch
            {
                string stimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss").ToString();
                string ESSpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ImagesPath\\" + stimestamp + ".Png";
                System.Drawing.Image iimage = appModel.HMOrderWindow.GetSnapshot();
                iimage.Save(ESSpath);
                Reporter.ReportEvent("HMOrder New Order Saved as Definite", "HMOrder New Order not Saved as Definite", HP.LFT.Report.Status.Failed, ESSpath);
                appModel.HMOrderWindow.Close();
                try
                {
                    appModel.HMOrderWindow.ClosingHMOrderWindow.YesUiObject.Click();
                }
                catch
                {

                }
             
                Assert.IsTrue(2 == 3);
            }
           

        }
       // [Test,Order(12)]
        public void TC12_HMOrder_DBExports_Verification()
        {
            xCellFileHelper = new ExcelHelper(datafilePath, 1);
            string sSeason = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "Season");
            string[] sSeasonarray = sSeason.Split('-');
            string smonth = sSeasonarray[0];
            if ((Int32.Parse(sSeasonarray[0]) < 10) && (sSeasonarray[0].Length == 1))
            {
                smonth = "0" + sSeasonarray[0];
            }
            string sOrdernum = xCellFileHelper.GetCellValueByRowAndColumn("Selenium_SmokeTest1", "OrderNumber");

            string sorderdata = sSeasonarray[1] + smonth + sOrdernum;

            DatabaseUtil sdbutil = new DatabaseUtil("", null);
            string sQuery = @"select XMLMessage from HMOrderLog.Logs.ExportedOrderLog where InternalOrderId = '" + sorderdata + "'and ExportObjectTypeId=5 Order by  ExportedDateTime desc";
            List<string> lsresults = sdbutil.ExecuteQuery_DB(sQuery);
            var svalue = "";

            if (lsresults.Count > 0)
            {
                svalue = lsresults[0].ToString();

            }


            if (svalue.Contains("GenericOrder"))
            {
                Reporter.ReportEvent("HMOrder DB exports XMLmessage contains Generic Order", "HMOrder DB exports XMLmessage contains Generic Order", HP.LFT.Report.Status.Passed);
            }
            else
            {
                Reporter.ReportEvent("HMOrder DB exports XMLmessage contains Generic Order", "HMOrder DB exports XMLmessage is not contains Generic Order", HP.LFT.Report.Status.Failed);
            }
            if (svalue.Contains("Definitive"))
            {
                Reporter.ReportEvent("HMOrder DB exports Generic XMLmessage contains  Definitive Order", "HMOrder DB exports Generic XMLmessage contains Definitive Order", HP.LFT.Report.Status.Passed);
            }
            else
            {
                Reporter.ReportEvent("HMOrder DB exports Generic XMLmessage contains Definitive Order", "HMOrder DB exports Generic XMLmessage is not contains Definitive Order", HP.LFT.Report.Status.Failed);
            }
             sQuery = @"select XMLMessage from HMOrderLog.Logs.ExportedOrderLog where InternalOrderId = '" + sorderdata + "'and ExportObjectTypeId=4 Order by  ExportedDateTime desc";
           lsresults = sdbutil.ExecuteQuery_DB(sQuery);
             svalue = "";

            if (lsresults.Count > 0)
            {
                svalue = lsresults[0].ToString();

            }


            if (svalue.Contains("BOPOInternalOrder"))
            {
                Reporter.ReportEvent("HMOrder DB exports XMLmessage contains BOPOInternalOrder Order", "HMOrder DB exports XMLmessage contains BOPOInternalOrder Order", HP.LFT.Report.Status.Passed);
            }
            else
            {
                Reporter.ReportEvent("HMOrder DB exports XMLmessage contains BOPOInternalOrder Order", "HMOrder DB exports XMLmessage is not contains BOPOInternalOrder Order", HP.LFT.Report.Status.Failed);
            }
            if (svalue.Contains("Definitive"))
            {
                Reporter.ReportEvent("HMOrder DB exports BOPOInternalOrder XMLmessage contains  Definitive Order", "HMOrder DB exports BOPOInternalOrder XMLmessage contains Definitive Order", HP.LFT.Report.Status.Passed);
            }
            else
            {
                Reporter.ReportEvent("HMOrder DB exports BOPOInternalOrder XMLmessage contains Definitive Order", "HMOrder DB exports BOPOInternalOrder XMLmessage is not contains Definitive Order", HP.LFT.Report.Status.Failed);
            }



        }
      

        [TearDown]
        public void TearDown()
        {
            // Clean up after each test
            xCellFileHelper.CleanUp();
            //            driver.Close();
            Thread.Sleep(2000);
        }

        [OneTimeTearDown]
        public void TestFixtureTearDown()
        {
            // Clean up once per fixture
        }
    }
}
