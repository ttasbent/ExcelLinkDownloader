using OfficeOpenXml;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Runtime.InteropServices;
using ExcelLinkDownloader.Models;

namespace ExcelLinkDownloader
{
    public class ReportGenerator
    {

        public void WriteToExcel(List<DownloadInfo> PdfFiles, FileInfo file)
        {
            PdfFiles.Sort((x, y) => x.BRNumber.CompareTo(y.BRNumber));

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (file.Exists)
            {
                file.Delete();
            }

            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.Add("DownloadPdfStatus");

                var range = ws.Cells["A2"].LoadFromCollection(PdfFiles, true);
                range.AutoFitColumns();

                // formatting the header
                ws.Cells["A1"].Value = "Download Report";
                ws.Cells["A1:B1"].Merge = true;
                ws.Row(1).Style.Font.Size = 24;

                ws.Row(2).Style.Font.Bold = true;
                ws.Column(2).Width = 24;

                package.SaveAsync();
            }
        }
    }
}