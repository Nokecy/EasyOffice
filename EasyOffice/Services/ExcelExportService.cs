using EasyOffice.Attributes;
using EasyOffice.Decorators;
using EasyOffice.Factories;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using EasyOffice.Providers.NPOI;
using System;
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

        public Task<byte[]> ExportAsync<T>(List<T> data,ExportOption<T> exportOption=null)
            where T : class, new()
        {
            exportOption = exportOption ?? new ExportOption<T>();

            var provider = _excelExportProvider;

            switch (exportOption.ExportType)
            {
                case Enums.ExportType.FastXLSX:
                    provider = new OpenXmlExcelExportProvider();
                    break;
                case Enums.ExportType.CSV:
                    provider = new CSVExcelExportProvider();
                    break;
            }

            if (exportOption.CustomExcelExportProvider != null)
            {
                provider = exportOption.CustomExcelExportProvider;
            }

            //if (exportOption.ExcelType == Enums.ExcelTypeEnum.XLS
            //    && exportOption.Data.Count > 65535)
            //{
            //    throw new InvalidOperationException("xls格式文件最多支持65536行数据");
            //}

            //if (exportOption.ExcelType == Enums.ExcelTypeEnum.XLSX
            //   && exportOption.Data.Count > 1048575)
            //{
            //    throw new InvalidOperationException("xlsx格式文件最多支持1048575行数据");
            //}

            var workbookBytes = provider.Export(data, exportOption);

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
