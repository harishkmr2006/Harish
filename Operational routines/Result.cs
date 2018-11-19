using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOPO.NUnit.ParallelTests
{
    public class Result
    {

        string result;
        string resultText;
        string resultScreenshot;

        public Result(string result, string resultText, string resultScreenshot)
        {
            this.result = result;
            this.resultText = resultText;
            this.resultScreenshot = resultScreenshot;
        }

        public void setResult(string result)
        {
            this.result = result;
        }

        public string getResult()
        {
            return this.result;
        }

        public void setResultText(string resultText)
        {
            this.resultText = resultText;
        }

        public string getResultText()
        {
            return this.resultText;
        }
        public void setResultScreenshot(string resultScreenshot)
        {
            this.resultScreenshot = resultScreenshot;
        }

        public string getResultScreenshot()
        {
            return this.resultScreenshot;
        }
    }

}
