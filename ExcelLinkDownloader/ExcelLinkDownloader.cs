using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ExcelLinkDownloader.Models;

namespace ExcelLinkDownloader
{
    public class ExcelLinkDownloader
    {
        public void Main()
        {
            UI UserInterface = new UI();

            UserInterface.UIFlow();

            PDFDownloader pdfDownloader = new PDFDownloader();

            pdfDownloader.DownloadFiles(UserInterface);

            //var file = new FileInfo(@"C:\Users\KOM\Desktop\PdfReport\DownloadReport.xlsx");
            new ReportGenerator().WriteToExcel(pdfDownloader.DownloadList, UserInterface.ReportFilePath);
        } 
    }
}


