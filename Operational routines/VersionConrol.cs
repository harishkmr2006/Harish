using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace BOPO.NUnit.ParallelTests
{
    public static class subKeys
    {
       public enum keySet
        {
            TestName,
            ApplictionName,
            Version,
            StartTime,
            EndTime,          
            TestResult
        };
    }
    public static class VersionConrol
    {
        public static List<string> appList = new List<string>();
        public static Dictionary<string, string> appDict = new Dictionary<string, string>();
        public static Dictionary<string, Dictionary<string, string>> appDictList = new Dictionary<string, Dictionary<string, string>>();
        public static DataTable globalTable = new DataTable();

        public static void addApplication(string appName, string appVriosn)
        {
            appDict.Add(appName, appVriosn);
            appList.Add(appName);
        }

        public static void _initTestData(string testName)
        {
            appDictList.Add(testName, new Dictionary<string, string>());
            Dictionary<string, string> localDic = new Dictionary<string, string>();
            foreach (var key in Enum.GetValues(typeof(subKeys.keySet)))
            {
                localDic.Add(key.ToString(), " ");
            }
            appDictList[testName] = localDic;
        }

        public static void addTestDataValue(string testName,string key,string KeyValue)
        {
            Dictionary<string, string> localdic = appDictList[testName];
            localdic[key] = KeyValue;
            appDictList[testName] = localdic;
        }
        
        public static void convertDicToDataTable()
        {
            globalTable = GetTable();


            foreach (var key in appDictList.Keys)
            {
                globalTable = addRow(appDictList[key], globalTable, key);
            }


         string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
         string projectPathAbs = new Uri(path.Substring(0, path.IndexOf("BOPO.NUnit.ParallelTests"))).LocalPath;

        string projectPath = projectPathAbs + @"\ExcelResults\" + DateTime.Now.ToString("ddd MMM d yyyy HH_mm");
            Directory.CreateDirectory(projectPath);

            // ExportToExcel(table, projectPath+"\\ExcelResult");

            MemoryStream ms = DataToExcel(globalTable);
            FileStream fs = new FileStream(projectPath + "\\"+DateTime.Now.ToString("ddd MMM d yyyy HH_mm")+ "ExcelResult.xls", FileMode.Create);
            ms.WriteTo(fs);
            fs.Close();
            ms.Close();
        }

        public static MemoryStream DataToExcel(DataTable dt)
        {
            MemoryStream ms = new MemoryStream();
            using (dt)
            {
                IWorkbook workbook = new HSSFWorkbook(); //Create an excel Workbook
                ISheet sheet = workbook.CreateSheet(); //Create a work table in the table
                IRow headerRow = sheet.CreateRow(0); //To add a row in the table
                foreach (DataColumn column in dt.Columns)
                    headerRow.CreateCell(column.Ordinal).SetCellValue(column.Caption);
                int rowIndex = 1;
                foreach (DataRow row in dt.Rows)
                {
                    IRow dataRow = sheet.CreateRow(rowIndex);
                    foreach (DataColumn column in dt.Columns)
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                    }
                    rowIndex++;
                }
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;
            }
            return ms;
        }

        public static DataTable addRow(Dictionary<string, string> localMap, DataTable table,string TestName)
        {
            //localMap = new Dictionary<string, string>();

            //string appanme = localMap[subKeys.keySet.ApplictionName.ToString()];
            //string version = localMap[subKeys.keySet.Version.ToString()];
            //string start =     localMap[subKeys.keySet.StartTime.ToString()];
            //string end =     localMap[subKeys.keySet.EndTime.ToString()];
            //string res =     localMap[subKeys.keySet.TestResult.ToString()];
            //table.Rows.Add(TestName, appanme, version, start, end, res);
            table.Rows.Add(
                TestName,
                localMap[subKeys.keySet.ApplictionName.ToString()],
                localMap[subKeys.keySet.Version.ToString()],
                localMap[subKeys.keySet.StartTime.ToString()],
                localMap[subKeys.keySet.EndTime.ToString()],
                localMap[subKeys.keySet.TestResult.ToString()]

                );
            return table;
        }

        static DataTable GetTable()
        {
            DataTable table = new DataTable();
            foreach (var enumValue in Enum.GetValues(typeof(subKeys.keySet)))
            {
                table.Columns.Add(enumValue.ToString(), typeof(string));
            }           
            return table;
        }

        //public static void testcasename(string applicationName)
        //{
        //    appDictList.Add(applicationName, new Dictionary<string, string>());
        //    appList.Add(applicationName);
        //   // _initTestData(applicationName);
        //}

        public static string getTimeStamp()
        {
            return DateTime.Now.ToString("yyyy:MM:dd_HH:mm:ss");
        }

        public static void addSubKey(string testName, string subKey, string subval)
        {
            Dictionary<string, string> local = appDictList[testName];
            //local.Add(subKey, subval);
            local[subKey] = subval;
            appDictList[testName] = local;
        }

        public static void printAll()
        {
            foreach(string str in appList)
            {
                Console.WriteLine("appName : " + str + " appVersion:"+appDict[str]);
            }
        }

    }
}
