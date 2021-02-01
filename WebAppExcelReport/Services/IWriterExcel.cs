using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebAppExcelReport.Models;

namespace WebAppExcelReport.Services
{
    public interface IWriterExcel
    {
        Task<Stream> GetFile(FullOrder fullOrder);
    }
}
