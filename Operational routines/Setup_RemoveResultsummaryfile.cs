using NUnit.Framework;
using System;
using System.IO;

namespace BOPO.NUnit.ParallelTests
{
    [TestFixture]

    // [Parallelizable]
    class Setup_RemoveResultsummaryfile
    {
     //   [Test, Category("FFRegression")]
        public void Setup_RemoveexistResultsummaryfile()
        {
            new EnvironmentSetUp();
            string rootFolderPath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ReportSummary";
            bool RMSfolderExists = Directory.Exists(rootFolderPath);
            if (!RMSfolderExists)
            {
                Directory.CreateDirectory(rootFolderPath);
                // rootFolderPath = rootFolderPath + "\\";


            }
            string filesToDelete = @"*ReportSummary*.html";   // Only delete DOC files containing "DeleteMe" in their filenames
            string[] fileList = System.IO.Directory.GetFiles(rootFolderPath, filesToDelete);
            foreach (string file in fileList)
            {
                // System.Diagnostics.Debug.WriteLine(file + "will be deleted");
                System.IO.File.Delete(file);
            }

            string reportIn = string.Empty;

            // string reportPath = "C:\\Project\\Reports\\report_template_ " + System.Environment.GetEnvironmentVariable("TestCaseName") + "_ " + GetCurrentDate() + ".html";
            string reportPath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory") + "ReportSummary\\ReportSummary.html";

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
    }
}
