using EasyOffice.Models.Excel;
using EasyOffice.Solutions.Entities;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace EasyOffice.Solutions.Repositories
{
    public interface IImportErrorLogRepository
    {
        Task<string> AddImportErrorLogAsync(List<ExcelDataRow> excelDataRows, string tag = "");

        Task<IEnumerable<ExcelImportErrorLog>> GetErrorLogsByTagAsync(string tag);
    }
}
