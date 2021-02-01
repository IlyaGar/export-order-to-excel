using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebAppExcelReport.Models
{
    public class Order
    {
        public int id { get; set; }
        public string creationDate { get; set; }
        public string statusId { get; set; }
        public string deparmentId { get; set; }
        public string storeId { get; set; }
        public string author { get; set; }
    }
}
