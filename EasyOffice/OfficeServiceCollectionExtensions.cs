using EasyOffice.Enums;
using EasyOffice.Interfaces;
using EasyOffice.Providers.NPOI;
using EasyOffice.Services;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace Bayantu.Extensions.DependencyInjection
{
    /// <summary>
    /// Office基础库依赖注入
    /// </summary>
    public static class OfficeServiceCollectionExtensions
    {

        public static void AddOffice(this IServiceCollection services)
        {
            services.AddOffice(new OfficeOptions());
        }

        public static void AddOffice(this IServiceCollection services, OfficeOptions options)
        {
            services.AddTransient<IExcelImportService, ExcelImportService>();
            services.AddTransient<IExcelExportService, ExcelExportService>();
            services.AddTransient<IWordExportService, WordExportService>();

            //根据配置项动态注入Provider
            if (options.ExcelImportSolution == SolutionEnum.NPOI)
            {
                services.AddTransient<IExcelImportProvider, ExcelImportProvider>();
            }

            if (options.ExcelExportSolution == SolutionEnum.NPOI)
            {
                services.AddTransient<IExcelExportProvider, ExcelExportProvider>();
            }

            if (options.WordExportSolution == SolutionEnum.NPOI)
            {
                services.AddTransient<IWordExportProvider, WordExportProvider>();
            }
        }
    }
}