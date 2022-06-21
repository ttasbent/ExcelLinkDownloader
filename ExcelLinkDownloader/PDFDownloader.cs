using ExcelLinkDownloader.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLinkDownloader
{
    public class PDFDownloader
    {
        public List<DownloadInfo> DownloadList = new List<DownloadInfo>();
        public object[,] valueArray;
        public int PrimaryLinkColumn = 38;
        public int SecondLinkColumn = 39;
        public FileInfo PDFsFilePath;

        public void DownloadFiles(UI userInterface)
        {
            PDFsFilePath = new FileInfo(userInterface.PDFsFilePath.FullName);
            valueArray = userInterface.valueArray;
            PrimaryLinkColumn = userInterface.PrimaryLinkColumn;
            SecondLinkColumn = userInterface.SecondLinkColumn;
            ThreadBeginner(userInterface);
            StatusResults();
        }

        private void ThreadBeginner(UI userInterface)
        {
            int Threads = ThreadCheck(userInterface.NumberOfThreads);
            int Interval = IntervalCheck(userInterface.RowsInExcelFile, userInterface.NumberOfFilesToDownload, Threads);

            Console.WriteLine("interval: " + Interval);
            Console.WriteLine("Threads: " + Threads);

            //Interval = 10;

            Thread[] downloadThreads = new Thread[Threads];

            for (int i = 0; i < Threads; i++)
            {
                int start = 0;
                int finish = 0;
                start = i * Interval + 2;
                finish = i * Interval + Interval + 2;
                if (i == 0)
                {
                    start = 2;
                }
                if (i > userInterface.RowsInExcelFile)
                {
                    finish = userInterface.RowsInExcelFile;
                }
                Console.WriteLine("start: " + start);
                Console.WriteLine("finish: " + finish);
                downloadThreads[i] = new Thread(() => DownloadFilesInIntervals(start, finish));
                downloadThreads[i].Start();
            }

            for (int i = 0; i < Threads; i++)
            {
                downloadThreads[i].Join();
            }
        }

        private int ThreadCheck(int numOfThreads)
        {
            int Threads = 0;
            if (numOfThreads <= 0)
            {
                Threads = 1;
            }
            else if (numOfThreads > 40)
            {
                Threads = 40;
            }
            else
            {
                Threads = numOfThreads;
            }
            return Threads;
        }

        private int IntervalCheck(int MaxNumOfFiles, int NumberOfDownloades, int threads)
        {
            int interval = 0;
            if (NumberOfDownloades <= 0 || NumberOfDownloades > MaxNumOfFiles)
            {
                interval = MaxNumOfFiles / threads;
            }
            else
            {
                interval = NumberOfDownloades / threads;
            }

            return interval;
        }

        private void DownloadFilesInIntervals(int start, int finish)
        {
            Console.WriteLine("starting download for " + Thread.CurrentThread.ManagedThreadId);
            List<DownloadInfo> subDownloadList = new List<DownloadInfo>();

            int f = 0;
            //for (int i = 300; i < valueArray.Length; i++)
            for (int i = start; i < finish; i++)
            {
                subDownloadList.Add(DownloadSingleFile(i));
            }
            Console.WriteLine("finishing download for " + Thread.CurrentThread.ManagedThreadId);
            DownloadList.AddRange(subDownloadList);
        }

        private DownloadInfo DownloadSingleFile(int rowNum)
        {
            string pdfAdd = "";
            int RowNumber = rowNum;
            string PdfCheck = "";
            Boolean firstLinkSuccess = false;
            string BRnum = valueArray[RowNumber, 1].ToString();
            var file = new FileInfo(PDFsFilePath.FullName + BRnum + ".pdf");
            if (valueArray[RowNumber, PrimaryLinkColumn] != null && valueArray[RowNumber, PrimaryLinkColumn] != null && valueArray[RowNumber, PrimaryLinkColumn] != "")
            {
                Console.WriteLine(valueArray[RowNumber, PrimaryLinkColumn].ToString());
                pdfAdd = valueArray[RowNumber, PrimaryLinkColumn].ToString();
                using (var client = new HttpClient())
                {
                    try
                    {
                        using (var s = client.GetStreamAsync(pdfAdd).GetAwaiter().GetResult())
                        {

                            using (var fs = new FileStream(file.FullName, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read))
                            {
                                s.CopyTo(fs);
                            }

                            using (var reader = new StreamReader(file.FullName))
                            {
                                PdfCheck = reader.ReadLine();

                                if (PdfCheck.Contains("PDF"))
                                {
                                    Console.WriteLine("First link: Downloadet");
                                    firstLinkSuccess = true;
                                    return new DownloadInfo(BRnum, "Downloadet");
                                }
                            }
                        }
                        file.Delete();
                        throw new Exception("Not a PDF");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(BRnum + " " + e);
                        Console.WriteLine("First link: Ikke Downloadet");
                        firstLinkSuccess = false;
                    }

                    if (firstLinkSuccess == false)
                    {
                        Console.WriteLine("Looking at second column link");
                        if (valueArray[RowNumber, SecondLinkColumn] != null && valueArray[RowNumber, SecondLinkColumn] != null && valueArray[RowNumber, SecondLinkColumn] != "")
                        {
                            Console.WriteLine(valueArray[RowNumber, SecondLinkColumn].ToString());
                            pdfAdd = valueArray[RowNumber, SecondLinkColumn].ToString();
                            try
                            {
                                using (var t = client.GetStreamAsync(pdfAdd).GetAwaiter().GetResult())
                                {
                                    using (var fs = new FileStream(file.FullName, FileMode.OpenOrCreate))
                                    {
                                        t.CopyTo(fs);
                                    }

                                    using (StreamReader reader = new StreamReader(file.FullName))
                                    {
                                        PdfCheck = reader.ReadLine();

                                        Console.WriteLine(BRnum + ". Contains pdf: " + PdfCheck.Contains("PDF"));
                                        if (PdfCheck.Contains("PDF"))
                                        {
                                            Console.WriteLine("Contains PDF");
                                            Console.WriteLine("second link: Downloadet");
                                            return new DownloadInfo(BRnum, "Downloadet");
                                        }
                                    }
                                }
                                file.Delete();
                                throw new Exception("Not a PDF");
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(BRnum + " " + e);
                                Console.WriteLine("second link: Ikke Downloadet");
                                return new DownloadInfo(BRnum, "Ikke Downloadet");
                            }
                        }
                        else
                        {
                            Console.WriteLine("No second link");
                            return new DownloadInfo(BRnum, "Ikke Downloadet");
                        }
                    }
                }
            }
            return new DownloadInfo(BRnum, "Ikke Downloadet");
        }

        private void StatusResults()
        {

            foreach (var item in DownloadList)
            {
                Console.WriteLine(item.BRNumber + " " + item.Status);
            }

            int count = DownloadList.Where(x => x.Status.Equals("Downloadet")).Count();
            int totalcount = DownloadList.Count();
            Console.WriteLine(count + " files out of " + totalcount + " were downloaded");
        }
    }
}
