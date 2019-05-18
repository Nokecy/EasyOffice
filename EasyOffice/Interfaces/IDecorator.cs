using EasyOffice.Decorators;
using EasyOffice.Models.Excel;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// 装饰器接口
    /// </summary>
    public interface IDecorator
    {
        byte[] Decorate<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context, IExcelExportProvider exportProvider)
            where T : class, new();
    }
}
