using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace SITSmokeTests
{
    public class DatabaseUtil
    {

        IWebDriver mainDriver;
        String sBrowserType;
        private string connectionstrings = string.Empty;
       
        public DatabaseUtil(String Browsertype, IWebDriver drivername)
        {
            sBrowserType = Browsertype;
           

            mainDriver = drivername;
            connectionstrings = ConfigurationManager.ConnectionStrings["SqlConnection"].ToString();
            Console.WriteLine(connectionstrings);
        }
        public List<string> ExecuteQuery_DB(string sQuery)
        {
            SqlConnection sqlcon = new SqlConnection(connectionstrings);
            SqlDataReader sqlrd;
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            try
            {
                //SqlConnection sqlcon = new SqlConnection(connectionstrings);
                // SqlDataReader sqlrd;
                // DataSet ds = new DataSet();
                // SqlCommand cmd = new SqlCommand();
                sqlcon.Open();
                cmd.CommandText = sQuery;
                Console.WriteLine(cmd.CommandText);
                cmd.Connection = sqlcon;

                //sqlda.SelectCommand.Connection = sqlcon;
                sqlrd = cmd.ExecuteReader();
                Console.WriteLine(sqlrd);
                List<string> lsData = new List<string>();
                if (sqlrd.HasRows)
                {
                    // for (int i = 0; i < sqlrd.FieldCount; i++)
                    while (sqlrd.Read())
                    {


                        lsData.Add(sqlrd[0].ToString());

                    }
                    sqlrd.Close();
                    return lsData;
                }
                else
                {
                    Console.WriteLine("No rows found.");
                    sqlrd.Close();
                    return null;
                }

            }
            catch
            {
                
                return null;
            }
            finally
            {
                if (sqlcon.State == ConnectionState.Open)
                {
                    sqlcon.Close();

                }
            }
        }




    }
}
