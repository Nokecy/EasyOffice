using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// Excel导出服务
    /// </summary>
    public interface IExcelExportService
    {
        Task<byte[]> ExportAsync<T>(List<T> data,ExportOption<T> exportOption = null)
            where T : class, new();
    }
}
