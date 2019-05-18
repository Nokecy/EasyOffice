using EasyOffice.Attributes;
using EasyOffice.Decorators;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;

namespace EasyOffice.Decorators
{
    /// <summary>
    /// 表头装饰器
    /// </summary>
    [BindDecorator(typeof(HeaderAttribute))]
    internal class HeaderDecorator : IDecorator
    {
        public byte[] Decorate<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context, IExcelExportProvider exportProvider)
            where T : class, new()
        {
            return exportProvider.DecorateHeader(workbookBytes, exportOption, context);
        }
    }
}
