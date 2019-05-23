using EasyOffice.Attributes;
using EasyOffice.Decorators;
using EasyOffice.Factories;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EasyOffice.Services
{
    public class ExcelExportService : IExcelExportService
    {
        private readonly IExcelExportProvider _excelExportProvider;
        public ExcelExportService(IExcelExportProvider excelExportProvider)
        {
            _excelExportProvider = excelExportProvider;
        }

        public Task<byte[]> ExportAsync<T>(ExportOption<T> exportOption)
            where T : class, new()
        {
            var provider = exportOption.CustomExcelExportProvider == null ? _excelExportProvider : _excelExportProvider;

            var workbookBytes = provider.Export(exportOption);

            //设置样式
            workbookBytes = Decorate(workbookBytes, exportOption);

            //返回byte数组
            return Task.FromResult(workbookBytes);
        }

        private byte[] Decorate<T>(byte[] workbookBytes, ExportOption<T> exportOption)
              where T : class, new()
        {
            var provider = exportOption.CustomExcelExportProvider == null ? _excelExportProvider : _excelExportProvider;

            DecoratorContext context = new DecoratorContext()
            {
                TypeDecoratorInfo = TypeDecoratorInfoFactory.CreateInstance(typeof(T))
            };

            GetDecorators<T>().ForEach(d =>
            {
                workbookBytes = d.Decorate(workbookBytes, exportOption, context, provider);
            });

            return workbookBytes;
        }
        /// <summary>
        /// 获取所有的装饰器
        /// </summary>
        /// <returns></returns>
        private static List<IDecorator> GetDecorators<T>()
                where T : class, new()

        {
            List<IDecorator> decorators = new List<IDecorator>();
            List<BaseDecorateAttribute> attrs = new List<BaseDecorateAttribute>();
            TypeDecoratorInfo typeDecoratorInfo = TypeDecoratorInfoFactory.CreateInstance(typeof(T));

            attrs.AddRange(typeDecoratorInfo.TypeDecoratorAttrs);
            typeDecoratorInfo.PropertyDecoratorInfos.ForEach(a => attrs.AddRange(a.DecoratorAttrs));

            attrs.Distinct(new DecoratorAttributeComparer()).ToList().ForEach
            (a =>
            {
                var decorator = DecoratorFactory.CreateInstance(a.GetType());
                if (decorator != null)
                {
                    decorators.Add(decorator);
                }
            });

            return decorators;
        }
    }
}
