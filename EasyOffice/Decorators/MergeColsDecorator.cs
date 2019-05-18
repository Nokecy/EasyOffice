using EasyOffice.Attributes;
using EasyOffice.Decorators;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;

namespace EasyOffice.Decorators
{
    /// <summary>
    /// 合并单元格装饰器
    /// </summary>
    [BindDecorator(typeof(MergeColsAttribute))]
    internal class MergeColsDecorator : IDecorator
    {
        public byte[] Decorate<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context, IExcelExportProvider exportProvider)
         where T : class, new()
        {
            return exportProvider.MergeCols(workbookBytes, exportOption, context);
        }
    }
}
