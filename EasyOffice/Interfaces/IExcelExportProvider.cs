using EasyOffice.Decorators;
using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// Excel导出Provider接口
    /// </summary>
    public interface IExcelExportProvider
    {
        byte[] Export<T>(List<T> data,ExportOption<T> exportOption)
            where T : class, new();

        byte[] DecorateHeader<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context)
            where T : class, new();

        byte[] MergeCols<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context)
          where T : class, new();

        byte[] WrapText<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context)
            where T : class, new();
    }
}
