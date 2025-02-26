using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AfifahExcelTest.Models
{
    public class RenameRule
    {
        public string SenderEmail { get; set; }
        public string Subject { get; set; }
        public string Folder { get; set; }
        public string FileName { get; set; }
        public bool IsRule { get; set; }
        public string Remarks { get; set; }
    }
}
