using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace BOPO.NUnit.ParallelTests
{
    public class XMLUtil
    {
       
        public string readXmlData_SignleNode(string Node)
        {
            try
            {
                new SetUp();
                string xmlfilepath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory")+ "BOPO.NUnit.ParallelTests\\BOPO.NUnit.ParallelTests\\XMLFileData.xml";
                XmlDocument doc = new XmlDocument();
                doc.Load(xmlfilepath);


                string Value = doc.DocumentElement.SelectSingleNode(Node).InnerText;
                return Value;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "";
            }
        }
        public void WriteXmlData_SingleNode(string Node,string value)
        {
            try
            {
                new SetUp();
                string xmlfilepath = System.Environment.GetEnvironmentVariable("ProjectWorkingDirectory")+ "BOPO.NUnit.ParallelTests\\BOPO.NUnit.ParallelTests\\XMLFileData.xml";
                XmlDocument doc = new XmlDocument();
                doc.Load(xmlfilepath);

                doc.DocumentElement.SelectSingleNode(Node).InnerText = value;
                doc.Save(xmlfilepath);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
               
            }
        }
    }
}
