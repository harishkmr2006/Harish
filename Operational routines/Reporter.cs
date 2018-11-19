using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace BOPO.NUnit.ParallelTests
{

    public class Reporter
    {
        public static List<Result> details = new List<Result>();
        public static string ResultPlaceholder = "<!-- INSERT_RESULTS -->";
        public static string TemplatePath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "report_template.html";
        public static string ReportSummaryTemplatePath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "report_Summary.html";
        public static Boolean IncludeScreenshots = true;
        public static string ScreenshotPath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "Reports\\screenshots\\";
        public static string Rreportpath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "Reports";
        public static string reportsummarypath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ReportSummary";
        public static List<Result> TESTREULTSLIST = new List<Result>();
        public Reporter()
        {
            // stestresult= details;
            // details = new List<Result>();
            // if (details.Count > 0)
            //  {
            //     stestresult.Add(details[details.Count - 1]);
            //  }

            //   TESTREULTSLIST = stestresult;
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
                    //  Result r = new Result("Pass", "Actual value '" + actualValue + "' matches expected value", "");
                    //  Result r = new Result("<font color = 'green'><strong> Pass </strong> </font>", "Actual value '" + actualValue + "' matches expected value", "");

                    Result r = new Result("<font color = 'green'><strong>Pass</strong> </font>", "Actual value '" + actualValue + "' matches expected value", "");
                    stestresult.Add(r);
                    // TESTREULTSLIST.Add(r);

                }
                else
                {
                    string ScreenshotPaths = "";

                    //if (IncludeScreenshots)
                    //{
                    ScreenshotPaths = GetScreenshot(sdriver);
                    Console.WriteLine("ScreenshotPaths value");
                    Console.WriteLine(ScreenshotPaths);
                    // }
                    // Result r = new Result("Fail", "expected value '" + expectedValue + "' does not match Actual value '" + actualValue + "'", "ScreenShot path-" + ScreenshotPaths);
                    Result r = new Result("<font color = 'red'><strong>Fail<strong></font>", "expected value and  Actual value are not Matching--" + actualValue + "--" + expectedValue, "ScreenShot path-" + ScreenshotPaths);
                    stestresult.Add(r);
                    // TESTREULTSLIST.Add(r);
                }
                // stestresult = TESTREULTSLIST;
            }
            catch
            {
                Console.WriteLine("Log message failed for " + actualValue);
                // Result r = new Result("Fail", "Actual value '" + actualValue + "' does not match expected value '" + expectedValue + "'", "screenshot not available");
                Result r = new Result("<font color = 'red'><strong> Fail <strong></font>", "Actual value '" + actualValue + "' does not match expected value '" + expectedValue + "'", "screenshot not available");
                stestresult.Add(r);
            }
        }

        public static string GetScreenshot(IWebDriver sdriver)
        {
            string stimestamp = GetCurrentDate().ToString();
            string location = ScreenshotPath + stimestamp + ".Jpeg";

            try
            {
                Screenshot sc = ((ITakesScreenshot)sdriver).GetScreenshot();
                sc.SaveAsFile(location, ImageFormat.Jpeg);
                //return location.Replace("\\", "\\\\");
                return stimestamp;
            }
            catch (IOException e)
            {
                //MessageBox.Show("Error while generating screenshot:-" + e.Message);
                Console.WriteLine("Error while generating screenshot:-" + e.Message);

            }

            return stimestamp;
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

            // string reportPath = "C:\\Project\\Reports\\report_template_ " + System.Environment.GetEnvironmentVariable("TestCaseName") + "_ " + GetCurrentDate() + ".html";
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

                    //File.AppendText(reportPath).WriteLine("!--INSERT_RESULTS--");
                }

                if (Bresult == false)
                {


                    reportIn = reportIn.Replace("<h3>Test Results Status---</h3>", "<h3>Test Results Status---<font color = 'red'><strong>Fail</strong></font></h3>");
                    WriteReportSummary_Fail(testname);
                }
                else
                {

                    reportIn = reportIn.Replace("<h3>Test Results Status---</h3>", "<h3>Test Results Status---<font color = 'green'><strong>Pass</strong></font></h3>");

                }
                File.WriteAllText(reportPath, reportIn);

            }
            catch (Exception e)
            {
                ///MessageBox.Show(“Unable to Write Log” + e.Message);
                Console.WriteLine("Unable to Write Log" + e.Message);
            }
        }


        public void WriteReportSummary_Fail(string testname)
        {
            string reportIn = string.Empty;

            // string reportPath = "C:\\Project\\Reports\\report_template_ " + System.Environment.GetEnvironmentVariable("TestCaseName") + "_ " + GetCurrentDate() + ".html";
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
                // rootFolderPath = rootFolderPath + "\\";


            }

            // string reportPath = "C:\\Project\\Reports\\report_template_ " + System.Environment.GetEnvironmentVariable("TestCaseName") + "_ " + GetCurrentDate() + ".html";
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
    }
}
