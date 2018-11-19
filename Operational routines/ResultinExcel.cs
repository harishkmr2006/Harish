using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using interOp = Microsoft.Office.Interop;

namespace BOPO.NUnit.ParallelTests.Wrappers
{
   class ResultinExcel
    {


        interOp.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        interOp.Excel.Workbook xlWorkBook;
        interOp.Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

    //        xlWorkBook = xlApp.Workbooks.Add(misValue);
    //        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            
    //        xlWorkSheet.Cells[1, 1] = "ID";
    //        xlWorkSheet.Cells[1, 2] = "Name";
    //        xlWorkSheet.Cells[2, 1] = "1";
    //        xlWorkSheet.Cells[2, 2] = "One";
    //        xlWorkSheet.Cells[3, 1] = "2";
    //        xlWorkSheet.Cells[3, 2] = "Two";



    //        xlWorkBook.SaveAs("d:\\csharp-Excel.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
    //        xlWorkBook.Close(true, misValue, misValue);
    //        xlApp.Quit();

    //        Marshal.ReleaseComObject(xlWorkSheet);
    //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
    //        Marshal.ReleaseComObject(xlApp);

    //         Console.WriteLine("Excel file created , you can find the file d:\\csharp-Excel.xls");
    //    }  

    //    }
    }
}
