using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLinkDownloader.Models
{
    public struct DownloadInfo
    {
        public DownloadInfo(string BRnum, string status)
        {
            BRNumber = BRnum;
            Status = status;
        }
        public string BRNumber { get; }
        public string Status { get; }
    }


}
