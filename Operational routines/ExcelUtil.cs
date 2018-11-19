using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Data.OleDb;



namespace BOPO.NUnit.ParallelTests
{
    public class ExcelUtil
    {
        public IExcelDataReader excelReader;
        public DataTable ExcelToDataTable(string fileName, IExcelDataReader stestid, FileStream stream)
        {
            //open file and returns as Stream
            try
            {
                stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                stestid = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            catch
            {

            }
            // FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
            // File sfile = new File(fileName);

            //  FileStream stream = File.OpenRead(fileName);
            //Createopenxmlreader via ExcelReaderFactory
            // IExcelDataReader excelReader1 = ExcelReaderFactory.CreateBinaryReader(new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read));


            //.xlsx
            //Set the First Row as Column Name
            stestid.IsFirstRowAsColumnNames = true;
            //Return as DataSet
            DataSet result = stestid.AsDataSet();
            //Get all the Tables
            DataTableCollection table = result.Tables;
            //Store it in DataTable
            DataTable resultTable = table["Sheet1"];

            stestid.Close();
            //stream.Close();

            //return
            return resultTable;

        }
        public class Datacollection
        {
            public int rowNumber { get; set; }
            public string colName { get; set; }
            public string colValue { get; set; }
        }
        public List<Datacollection> dataCol = new List<Datacollection>();

        public void PopulateInCollection(string fileName, IExcelDataReader stestid, DataTable table, FileStream stream)
        {
            table = ExcelToDataTable(fileName, stestid, stream);
            // DataTable table = TestBase.stable;
            //Iterate through the rows and columns of the Table
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    Datacollection dtTable = new Datacollection()
                    {
                        rowNumber = row,
                        colName = table.Columns[col].ColumnName,
                        colValue = table.Rows[row - 1][col].ToString()
                    };
                    //Add all the details for each row
                    dataCol.Add(dtTable);
                }
            }
        }
        public int RowsCount(string fileName, IExcelDataReader stestid, FileStream stream)
        {
            DataTable table = ExcelToDataTable(fileName, stestid, stream);
            return table.Rows.Count;
        }
        public string ReadData(int rowNumber, string columnName)
        {
            try
            {
                //Retriving Data using LINQ to reduce much of iterations
                string data = (from colData in dataCol
                               where colData.colName == columnName && colData.rowNumber == rowNumber
                               select colData.colValue).SingleOrDefault();

                //var datas = dataCol.Where(x => x.colName == columnName && x.rowNumber == rowNumber).SingleOrDefault().colValue;
                return data.ToString();
            }
            catch
            {
                return null;
            }
        }
        public int GetTestCaseRownumber(String fileName, String sTestCaseName, IExcelDataReader sTestCaseid, FileStream sStream)
        {
            try
            {
                // PopulateInCollection(fileName);
                Reporter summarywriteresult = new Reporter();
                summarywriteresult.WriteReportSummary_Pass(sTestCaseName);
                int iirownumber = 1;
                Console.WriteLine(ReadData(iirownumber, "TestCaseName"));
                while (ReadData(iirownumber, "TestCaseName") != sTestCaseName)
                {
                    // Console.WriteLine(ReadData(iirownumber, "TestCaseName"));

                    if (iirownumber == RowsCount(fileName, sTestCaseid, sStream))
                    {
                        Console.WriteLine("Test case name not found in the test data sheet");
                    }
                    iirownumber++;
                }
                return iirownumber;
            }
            catch
            {
                return 0;
            }
        }

        public void WriteData(string fileName, string stestcaseid, string columnName,string value)
        {
            try
            {
                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();
                string sql = null;
                MyConnection = new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ fileName+";Extended Properties=Excel 8.0;");
                MyConnection.Open();
                myCommand.Connection = MyConnection;
                sql = "Update [Sheet1$] set "+ columnName+" = '"+ value + "' where TestCaseName='"+ stestcaseid+"'";
                myCommand.CommandText = sql;
                myCommand.ExecuteNonQuery();
               
                MyConnection.Close();

            }
            catch
            {
               
            }
        }

    }


}
