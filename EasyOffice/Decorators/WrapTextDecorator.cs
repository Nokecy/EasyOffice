using EasyOffice.Attributes;
using EasyOffice.Decorators;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;

namespace EasyOffice.Decorators
{
    /// <summary>
    /// 自动换行装饰器
    /// </summary>
    [BindDecorator(typeof(WrapTextAttribute))]
    internal class WrapTextDecorator : IDecorator
    {
        public byte[] Decorate<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context, IExcelExportProvider exportProvider)
      where T : class, new()
        {
            return exportProvider.WrapText(workbookBytes, exportOption, context);
        }
    }
}
