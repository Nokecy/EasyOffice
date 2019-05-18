using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// Excel支持接口
    /// </summary>
    public interface IExcelImportProvider
    {
        ExcelData Convert<TTemplate>(string fileUrl
             , int sheetIndex = 0, int headerRowIndex = 0, int dataRowStartIndex = 1)
             where TTemplate : class, new();
    }
}
