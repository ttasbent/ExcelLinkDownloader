using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLinkDownloader
{
    public class UI
    {
        string filePath = "";
        public FileInfo ReportFilePath;
        public FileInfo PDFsFilePath;

        public object[,] valueArray;
        public int PrimaryLinkColumn = 38;
        public int SecondLinkColumn = 39;

        public int NumberOfThreads = 0;
        public int NumberOfFilesToDownload = 0;
        public int RowsInExcelFile = 0;



        ExcelImporter ExcelFile = new ExcelImporter();

        public void UIFlow()
        {
            ExcelFileInput();
            ColumnInfo();
            ThreadNumber();
            FilesNumber();
            LocationOfReport();
            LocationOfPDFs();

            Console.WriteLine("Begin?");
            Console.Read();
        }

        private void LocationOfPDFs()
        {
            Console.WriteLine("Specify the path where you want the downloaded PDFs to be located");
            Console.WriteLine(@"Ex: C:\Users\KOM\Desktop\test");
            string pdfFilePath = Console.ReadLine();
            if (Directory.Exists(pdfFilePath))
            {
                PDFsFilePath = new FileInfo(pdfFilePath + @"\");
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }
            else
            {
                Console.WriteLine(@"Invalid path. C:\Users\KOM\Desktop\DownloadedPdfs is chosen instead");
                PDFsFilePath = new FileInfo(@"C:\Users\KOM\Desktop\DownloadedPdfs\");
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }
        }

        private void LocationOfReport()
        {
            Console.WriteLine("Specify the path where you want the Report to be generated");
            Console.WriteLine(@"Ex: C:\Users\KOM\Desktop\test");
            string reportFilePath = Console.ReadLine();
            if (Directory.Exists(reportFilePath))
            {
                ReportFilePath = new FileInfo(reportFilePath + @"\DownloadReport.xlsx");
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }
            else
            {
                Console.WriteLine(@"Invalid path. C:\Users\KOM\Desktop\PdfReport is chosen instead");
                ReportFilePath = new FileInfo(@"C:\Users\KOM\Desktop\PdfReport\DownloadReport.xlsx");
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }
        }

        private void ExcelFileInput()
        {
            Console.WriteLine("Hey");
            Console.WriteLine("Input the path of the excel file with the pdf links, which needs to be downloaded. Remember suffix (for example xlsx)");

            filePath = Console.ReadLine();
            if (File.Exists(filePath))
            {
                try
                {
                    Console.WriteLine("Importing file...");
                    ExcelFile.GetFile(filePath);
                    Console.WriteLine("File Imported");
                    Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Invalid input. The default path " + @"C:\Users\KOM\Downloads\GRI_2017_2020 (1) is chosen");
                    Console.WriteLine("Importing file...");
                    filePath = @"C:\Users\KOM\Downloads\GRI_2017_2020 (1)";
                    ExcelFile.GetFile(filePath);
                    Console.WriteLine("File Imported");
                    Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
                }
            }
            else
            {
                Console.WriteLine("Invalid input. The default path " + @"C:\Users\KOM\Downloads\GRI_2017_2020 (1) is chosen");
                Console.WriteLine("Importing file...");
                filePath = @"C:\Users\KOM\Downloads\GRI_2017_2020 (1)";
                ExcelFile.GetFile(filePath);
                Console.WriteLine("File Imported");
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }

            valueArray = ExcelFile.valueArray;
            RowsInExcelFile = ExcelFile.RowsInExcelFile;
        }

        private void ColumnInfo()
        {
            Console.WriteLine("Column number for first Link");
            try
            {
                PrimaryLinkColumn = Int32.Parse(Console.ReadLine());
            }
            catch (Exception e)
            {
                Console.WriteLine("Invalid input. 38 is chosen as the default column");
                PrimaryLinkColumn = 38;
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }

            Console.WriteLine("Column number for second link");
            try
            {
                SecondLinkColumn = Int32.Parse(Console.ReadLine());
            }
            catch (Exception e)
            {
                Console.WriteLine("Invalid input. 39 is chosen as the default column");
                SecondLinkColumn = 39;
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }
        }

        private void ThreadNumber() 
        {
            Console.WriteLine("Number of threads? Default value is 20. Capped to a maximum of 40.");
            try
            {
                NumberOfThreads = Int32.Parse(Console.ReadLine());
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                NumberOfThreads = 20;
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }
        }

        private void FilesNumber()
        {
            int max = ExcelFile.RowsInExcelFile - 1;
            Console.WriteLine("Number of files that you want to try to download? Maximum is: " + max);
            Console.WriteLine("If no number is supplied, all possible files will be downloaded.");
            try
            {
                NumberOfFilesToDownload = Int32.Parse(Console.ReadLine());
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }
            catch (Exception e)
            {
                NumberOfFilesToDownload = ExcelFile.RowsInExcelFile;
                Console.WriteLine("\r\n" + "------------------------------------------------------------------------");
            }
        }
    }
}