using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebAppExcelReport.Models
{
    public class FullOrder
    {
        public Order order { get; set; }
        public List<OrderBody> orderBodies { get; set; }
    }
}
