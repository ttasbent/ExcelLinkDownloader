using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace ExcelLinkDownloader
{
    public class ExcelImporter
    {
        public int ColumnsInExcelFile = 0;
        public int RowsInExcelFile = 0;
        public object[,] valueArray;

        public void GetFile(string filePath)
        {
            //string filePath = @"C:\Users\KOM\Downloads\GRI_2017_2020 (1)";

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;


            valueArray = (object[,])xlRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            ColumnsInExcelFile = xlWorksheet.UsedRange.Columns.Count;
            RowsInExcelFile = xlWorksheet.UsedRange.Rows.Count;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorksheet);
        }
    }
}
