using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebAppExcelReport.Models
{
    public class OrderBody
    {
        public int id { get; set; }
        public int docId { get; set; }
        public string group { get; set; }
        public string article { get; set; }
        public string barcode { get; set; }
        public string name { get; set; }
        public string supplier { get; set; }
        public string goods { get; set; }
        public string average { get; set; }
        public string stockDay { get; set; }
        public string deliveryDate { get; set; }
        public string managerСomment { get; set; }
        public string departmentComment { get; set; }
    }
}