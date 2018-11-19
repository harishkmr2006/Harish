using Microsoft.Expression.Encoder.ScreenCapture;
using OpenQA.Selenium;
using RelevantCodes.ExtentReports;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace BOPO.NUnit.ParallelTests
{

    public class ReportCustom : TestBase
    {
        public static List<Result> details = new List<Result>();
        public static string ResultPlaceholder = "<!-- INSERT_RESULTS -->";
        public static string TemplatePath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "report_template.html";
        public static string ReportSummaryTemplatePath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "report_Summary.html";
        public static Boolean IncludeScreenshots = true;
        public static string ScreenshotPath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "Reports\\screenshots\\";
        public static string ExtentScreenshotPath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "Reports\\ExtentScreenShots\\";
        public static string Rreportpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "Reports";
        public static string TestVideospath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "Reports\\TestVideos\\";
      //  public static readonly ScreenCaptureJob vidRec = new ScreenCaptureJob();
        public static string reportsummarypath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ReportSummary";
        public static List<Result> TESTREULTSLIST = new List<Result>();

        //--------------------------------------------------------------------------------------------------------------------------------------------------------

        public ReportCustom()
        {

            bool RPfolderExists = Directory.Exists(Rreportpath);
            if (!RPfolderExists)
            {
                Directory.CreateDirectory(Rreportpath);
                Rreportpath = Rreportpath + "\\";
            }

            bool SNfolderExists = Directory.Exists(ScreenshotPath);
            if (!SNfolderExists)
            {
                Directory.CreateDirectory(ScreenshotPath);
                ScreenshotPath = ScreenshotPath + "\\";
            }

            bool TVfolderExists = Directory.Exists(Rreportpath);
            if (!TVfolderExists)
            {
                Directory.CreateDirectory(Rreportpath);
                TestVideospath = TestVideospath + "\\";
            }

            bool RMfolderExists = Directory.Exists(reportsummarypath);
            if (!RMfolderExists)
            {
                Directory.CreateDirectory(reportsummarypath);
                reportsummarypath = reportsummarypath + "\\";
            }

        }

        public static void Report(List<Result> stestresult, IWebDriver sdriver, string actualValue, string expectedValue)
        {

            try
            {
                if (actualValue.Equals(expectedValue))
                {
                    Result r = new Result("<font color = 'green'><strong>Pass</strong> </font>", "Actual value '" + actualValue + "' matches expected value", "");
                    stestresult.Add(r);

                    //---------------EXTENT REPORT GENERAL STEP LOG PASS------------------------------------------------------------
                    test.Log(LogStatus.Pass, "  " + actualValue, "<font color = 'green'><strong>Pass</strong> </font>");
                    //---------------------------------------------------------------------------------------------------------
                }
                else
                {
                    string ScreenshotPaths = "";
                    ReportCustom Eresult = new ReportCustom();
                    ScreenshotPaths = GetScreenshot(sdriver);
                    Console.WriteLine("ScreenshotPaths value");
                    Console.WriteLine(ScreenshotPaths);
                    Result r = new Result("<font color = 'red'><strong>Fail<strong></font>", "Expected value and  Actual value are not Matching--" + actualValue + "--" + expectedValue, "ScreenShot path-" + ScreenshotPaths);
                    stestresult.Add(r);

                    //---------------EXTENT REPORT GENERAL STEP LOG FAIL------------------------------------------------------------
                    if (actualValue.Contains("Exception"))
                    {
                        test.Log(LogStatus.Warning, " " + actualValue, "<font color = 'orange'><strong>Exception<strong></font>");
                        string ExtentScreenShotPath = ReportCustom.Capture(sdriver);
                        test.Log(LogStatus.Warning, "Snapshot below: " + test.AddScreenCapture(ExtentScreenShotPath), "<font color = 'orange'><strong>Exception<strong></font>");
                    }
                    else
                    {
                        test.Log(LogStatus.Fail, "Mismatch in Actual & Expected: " + actualValue + expectedValue, "<font color = 'red'><strong>Fail<strong></font>");
                        string ExtentScreenShotPath = ReportCustom.Capture(sdriver);
                        test.Log(LogStatus.Fail, "Snapshot below: " + test.AddScreenCapture(ExtentScreenShotPath), "<font color = 'red'><strong>Fail<strong></font>");
                    }

                    //--------------------------------------------------------------------------------------------------------------

                }

            }
            catch
            {
                Console.WriteLine("Log message failed for " + actualValue);
                Result r = new Result("<font color = 'red'><strong> Fail <strong></font>", "Actual value '" + actualValue + "' does not match expected value '" + expectedValue + "'", "screenshot not available");
                stestresult.Add(r);
            }
        }

        public static void Report(List<Result> stestresult, IWebDriver sdriver, string actualValue, string expectedValue, string executionType)
        {

            try
            {
                if (actualValue.Equals(expectedValue))
                {
                    Result r = new Result("<font color = 'green'><strong>Pass</strong> </font>", "Actual value '" + actualValue + "' matches expected value", "");
                    stestresult.Add(r);

                    //---------------EXTENT REPORT GENERAL STEP LOG PASS------------------------------------------------------------
                    test.Log(LogStatus.Pass, "  " + actualValue, "<font color = 'green'><strong>Pass</strong> </font>");
                    //---------------------------------------------------------------------------------------------------------
                }
                else
                {
                    string ScreenshotPaths = "";
                    ReportCustom Eresult = new ReportCustom();
                    Console.WriteLine("ScreenshotPaths value");
                    Console.WriteLine(ScreenshotPaths);
                    Result r = new Result("<font color = 'red'><strong>Fail<strong></font>", "Expected value and  Actual value are not Matching--" + actualValue + "--" + expectedValue, "");
                    stestresult.Add(r);

                    //---------------EXTENT REPORT GENERAL STEP LOG FAIL------------------------------------------------------------
                    if (actualValue.Contains("Exception"))
                    {
                        test.Log(LogStatus.Warning, " " + actualValue, "<font color = 'orange'><strong>Exception<strong></font>");
                        string ExtentScreenShotPath = ReportCustom.Capture(sdriver);
                        test.Log(LogStatus.Warning, "Snapshot below: " + test.AddScreenCapture(ExtentScreenShotPath), "<font color = 'orange'><strong>Exception<strong></font>");
                    }
                    else
                    {
                        test.Log(LogStatus.Fail, "Mismatch in Actual & Expected: " + actualValue + expectedValue, "<font color = 'red'><strong>Fail<strong></font>");
                        string ExtentScreenShotPath = ReportCustom.Capture(sdriver);
                        test.Log(LogStatus.Fail, "Snapshot below: " + test.AddScreenCapture(ExtentScreenShotPath), "<font color = 'red'><strong>Fail<strong></font>");
                    }

                    //--------------------------------------------------------------------------------------------------------------

                }

            }
            catch
            {
                Console.WriteLine("Log message failed for " + actualValue);
                Result r = new Result("<font color = 'red'><strong> Fail <strong></font>", "Actual value '" + actualValue + "' does not match expected value '" + expectedValue + "'", "screenshot not available");
                stestresult.Add(r);
            }
        }

        public static void addScreenShotToReport()
        {
            test.Log(LogStatus.Pass, "  " + test.AddScreenCapture(takeWholeScreenshot()), "");
        }

        public static string takeWholeScreenshot()
        {
            string localPath = ScreenshotPath + GetCurrentDate().ToString() + ".png";
            TestStack.White.Desktop.TakeScreenshot(localPath, System.Drawing.Imaging.ImageFormat.Png);
            System.Drawing.Bitmap bitmap = TestStack.White.Desktop.CaptureScreenshot();
            return localPath;
        }

        public static string GetScreenshot(IWebDriver sdriver)
        {
            string stimestamp = GetCurrentDate().ToString();
            string location = ScreenshotPath + stimestamp + ".Jpeg";

            try
            {
                Screenshot sc = ((ITakesScreenshot)sdriver).GetScreenshot();
                sc.SaveAsFile(location, ImageFormat.Jpeg);
                return stimestamp;

            }
            catch (IOException e)
            {
                Console.WriteLine("Error while generating screenshot:-" + e.Message);
            }

            return stimestamp;

        }

        public static string Capture(IWebDriver sdriver)
        {
            string stimestamp = GetCurrentDate().ToString();
            string ESSpath = ExtentScreenshotPath + stimestamp + ".Jpeg";
            Screenshot sc = ((ITakesScreenshot)sdriver).GetScreenshot();
            sc.SaveAsFile(ESSpath, ImageFormat.Jpeg);
            return ESSpath;
        }



        public static string GetCurrentDate()
        {
            return DateTime.Now.ToString("yyyyMMdd_HHmmss");
        }

        public void InsertStartTime()
        {
            System.Environment.SetEnvironmentVariable("Starttime", GetCurrentDate());
        }
        public void InsertEndtTime()
        {

            System.Environment.SetEnvironmentVariable("Endtime", GetCurrentDate());

        }

        public void WriteResults(string testname, List<Result> stestresult)
        {
            string reportIn = string.Empty;
            if (Rreportpath.PadRight(1) != "\\")
            {
                Rreportpath = Rreportpath + "\\";
            }
            else
            {
                Rreportpath = Rreportpath + "\\";
            }


            string reportPath = Rreportpath + "report_" + testname + "_ " + GetCurrentDate() + ".html";
            try
            {
                using (StreamReader reader = new StreamReader(TemplatePath))
                {
                    String line = String.Empty;

                    if ((line = reader.ReadToEnd()) != null)
                    {
                        reportIn += line;
                    }
                }

                reportIn = reportIn.Replace("Test Case Name: ", "Test Case Name: " + testname);
                reportIn = reportIn.Replace("Objective : ", "Objective : " + System.Environment.GetEnvironmentVariable("TestObjective"));
                reportIn = reportIn.Replace("EXECUTION START TIME---", "EXECUTION START TIME---" + System.Environment.GetEnvironmentVariable("Starttime"));
                reportIn = reportIn.Replace("---EXECUTION END TIME---", "---EXECUTION END TIME---" + System.Environment.GetEnvironmentVariable("Endtime"));
                Boolean Bresult = true;
                for (int i = 0; i < stestresult.Count(); i++)
                {

                    if (stestresult[i].getResultScreenshot().Equals(""))
                    {
                        reportIn = reportIn.Replace(ResultPlaceholder,
                        "<tr><td>" + (i + 1).ToString() + "</td>\n<td>" +
                        stestresult[i].getResult() + "</td>\n<td>" +
                        stestresult[i].getResultText() + "</td></tr>\n" + ResultPlaceholder);
                    }
                    else
                    {
                        Bresult = false;
                        reportIn = reportIn.Replace(ResultPlaceholder,
                        "<tr><td>" + (i + 1).ToString() + "</td>\n<td>" +
                        stestresult[i].getResult() + "</td>\n<td>" +
                        stestresult[i].getResultText() + "<img src =\"" + ScreenshotPath + (stestresult[i].getResultScreenshot().Replace("ScreenShot path-", "")) + ".Jpeg\"</td></tr>\n" + ResultPlaceholder);
                    }
                    File.WriteAllText(reportPath, reportIn);


                }

                if (Bresult == false)
                {
                    reportIn = reportIn.Replace("<h3>Test Results Status---</h3>", "<h3>Test Results Status---<font color = 'red'><strong>Fail</strong></font></h3>");
                    WriteReportSummary_Fail(testname);
                }
                else
                {
                    reportIn = reportIn.Replace("<h3>Test Results Status---</h3>", "<h3>Test Results Status---<font color = 'green'><strong>Pass</strong></font></h3>");
                    WriteReportSummary_Pass(testname);
                }

                File.WriteAllText(reportPath, reportIn);

            }
            catch (Exception e)
            {
                Console.WriteLine("Unable to Write Log" + e.Message);
            }
        }


        public void WriteReportSummary_Fail(string testname)
        {
            string reportIn = string.Empty;

            if (reportsummarypath.PadRight(1) != "\\")
            {
                reportsummarypath = reportsummarypath + "\\";
            }
            else
            {
                reportsummarypath = reportsummarypath + "\\";
            }
            string reportPath = reportsummarypath + "ReportSummary.html";

            using (StreamReader reader = new StreamReader(reportsummarypath + "ReportSummary.html"))
            {
                String line = String.Empty;

                if ((line = reader.ReadToEnd()) != null)
                {
                    reportIn += line;
                }
            }

            reportIn = reportIn.Replace("<td>" + testname + "</td><td><font color = 'green'><strong>&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp PASS</strong></font></td>", "<td>" + testname + "</td><td><font color = 'red'><strong>&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp FAIL</strong></font></td>");

            File.WriteAllText(reportPath, reportIn);
        }
        public void WriteReportSummary_Pass(string testname)
        {
            string reportIn = string.Empty;
            string reportPath = string.Empty;
            string rootFolderPath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ReportSummary";
            bool RMSfolderExists = Directory.Exists(rootFolderPath);
            if (!RMSfolderExists)
            {
                Directory.CreateDirectory(rootFolderPath);

            }

            if (reportsummarypath.PadRight(1) != "\\")
            {
                reportsummarypath = reportsummarypath + "\\";
            }
            else
            {
                reportsummarypath = reportsummarypath + "\\";
            }

            reportPath = reportsummarypath + "ReportSummary.html";
            bool RSFileExist = File.Exists(reportPath);

            if (!RSFileExist)
            {
                using (StreamReader reader = new StreamReader(System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "report_Summary.html"))
                {
                    String line = String.Empty;

                    if ((line = reader.ReadToEnd()) != null)
                    {
                        reportIn += line;
                    }
                }
                File.WriteAllText(reportPath, reportIn);
            }
            reportIn = string.Empty;
            using (StreamReader reader = new StreamReader(reportsummarypath + "ReportSummary.html"))
            {
                String line = String.Empty;

                if ((line = reader.ReadToEnd()) != null)
                {
                    reportIn += line;
                }
            }

            if (reportIn.Contains("<td>" + testname + "</td><td>PASS</td>") == false)
            {
                reportIn = reportIn.Replace("<!-- INSERT_RESULTS -->", "<tr><td>" + testname + "</td><td><font color = 'green'><strong>&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp PASS</strong></font></td></tr>" + "<!-- INSERT_RESULTS -->");

            }
            File.WriteAllText(reportPath, reportIn);
        }

        public void WriteResults(string testname, List<Result> stestresult, bool finalResult)
        {
            string reportIn = string.Empty;
            if (Rreportpath.PadRight(1) != "\\")
            {
                Rreportpath = Rreportpath + "\\";
            }
            else
            {
                Rreportpath = Rreportpath + "\\";
            }


            string reportPath = Rreportpath + "report_" + testname + "_ " + GetCurrentDate() + ".html";
            try
            {
                using (StreamReader reader = new StreamReader(TemplatePath))
                {
                    String line = String.Empty;

                    if ((line = reader.ReadToEnd()) != null)
                    {
                        reportIn += line;
                    }
                }

                reportIn = reportIn.Replace("Test Case Name: ", "Test Case Name: " + testname);
                reportIn = reportIn.Replace("Objective : ", "Objective : " + System.Environment.GetEnvironmentVariable("TestObjective"));
                reportIn = reportIn.Replace("EXECUTION START TIME---", "EXECUTION START TIME---" + System.Environment.GetEnvironmentVariable("Starttime"));
                reportIn = reportIn.Replace("---EXECUTION END TIME---", "---EXECUTION END TIME---" + System.Environment.GetEnvironmentVariable("Endtime"));
                Boolean Bresult = true;

                for (int i = 0; i < stestresult.Count(); i++)
                {

                    if (stestresult[i].getResultScreenshot().Equals(""))
                    {
                        reportIn = reportIn.Replace(ResultPlaceholder,
                        "<tr><td>" + (i + 1).ToString() + "</td>\n<td>" +
                        stestresult[i].getResult() + "</td>\n<td>" +
                        stestresult[i].getResultText() + "</td></tr>\n" + ResultPlaceholder);
                    }
                    else
                    {
                        Bresult = false;
                        reportIn = reportIn.Replace(ResultPlaceholder,
                        "<tr><td>" + (i + 1).ToString() + "</td>\n<td>" +
                        stestresult[i].getResult() + "</td>\n<td>" +
                        stestresult[i].getResultText() + "<img src =\"" + ScreenshotPath + (stestresult[i].getResultScreenshot().Replace("ScreenShot path-", "")) + ".Jpeg\"</td></tr>\n" + ResultPlaceholder);
                    }
                    File.WriteAllText(reportPath, reportIn);


                }

                if (Bresult == false)
                {
                    if (finalResult)
                    {
                        reportIn = reportIn.Replace("<h3>Test Results Status---</h3>", "<h3>Test Results Status---<font color = 'red'><strong>Fail</strong></font></h3>");
                        test.Log(LogStatus.Fail, "Test Case Failed");
                        WriteReportSummary_Fail(testname);
                    }
                    else
                    {
                        reportIn = reportIn.Replace("<h3>Test Results Status---</h3>", "<h3>Test Results Status---<font color = 'red'><strong>Fail (Terminated)</strong></font></h3>");
                        test.Log(LogStatus.Fail, "Fail (Terminated)");
                        WriteReportSummary_Fail(testname);
                    }

                }
                else
                {

                    if (finalResult)
                    {
                        reportIn = reportIn.Replace("<h3>Test Results Status---</h3>", "<h3>Test Results Status---<font color = 'green'><strong>Pass</strong></font></h3>");
                        test.Log(LogStatus.Pass, "Test Case Passed");
                        WriteReportSummary_Pass(testname);
                    }
                    else
                    {
                        reportIn = reportIn.Replace("<h3>Test Results Status---</h3>", "<h3>Test Results Status---<font color = 'green'><strong>Pass (Terminated)</strong></font></h3>");
                        test.Log(LogStatus.Pass, "Pass (Terminated)");
                        WriteReportSummary_Pass(testname);
                    }

                }
                File.WriteAllText(reportPath, reportIn);

            }
            catch (Exception e)
            {
                Console.WriteLine("Unable to Write Log" + e.Message);
            }
        }

        //public static void StartRecordingVideo(string scenarioTitle)

        //{
        //    string timestamp = DateTime.Now.ToString("dd-MM-yyyy-hh-mm-ss");
        //    vidRec.OutputScreenCaptureFileName = TestVideospath + scenarioTitle + " " + timestamp + ".wmv";
        //    vidRec.Start();
        //}

        //public static void EndRecording()
        //{
        //    vidRec.Stop();
        //    //vidRec.Dispose();
        //}

    }
}

