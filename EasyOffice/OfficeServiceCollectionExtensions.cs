using System;
using Autofac;
using Bayantu.Extensions.Cache;
using EasyOffice.Enums;
using EasyOffice.Providers.NPOI;
using EasyOffice.Services;
using EasyOffice.Solutions;
using EasyOffice.Solutions.DbContexts;
using EasyOffice.Solutions.Repositories;
using Bayantu.Extensions.Persistence;
using Bayantu.Extensions.Persistence.Dapper;
using Bayantu.Extensions.Session;
using Microsoft.Extensions.DependencyInjection;
using EasyOffice.Interfaces;

namespace Bayantu.Extensions.DependencyInjection
{
    /// <summary>
    /// Office基础库依赖注入
    /// </summary>
    public static class OfficeServiceCollectionExtensions
    {

        public static void AddOffice(this IServiceCollection services, OfficeOptions options)
        {
            services.AddTransient<IExcelImportService,ExcelImportService>();
            services.AddTransient<IExcelExportService,ExcelExportService>();
            services.AddTransient<IExcelImportSolutionService,ExcelImportSolutionService>();
            services.AddTransient<IWordExportService,WordExportService>();

            //根据配置项动态注入Provider
            if (options.ExcelImportSolution == SolutionEnum.NPOI)
            {
                services.AddTransient<IExcelImportProvider,ExcelImportProvider>();
            }

            if (options.ExcelExportSolution == SolutionEnum.NPOI)
            {
                services.AddTransient<IExcelExportProvider,ExcelExportProvider>();
            }

            if (options.WordExportSolution == SolutionEnum.NPOI)
            {
                services.AddTransient<IWordExportProvider,WordExportProvider>();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="options"></param>
        public static void AddOffice(this ContainerBuilder builder, OfficeOptions options)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            if (options == null)
            {
                throw new ArgumentNullException(nameof(options));
            }
            
            AddOfficeServices(builder, options);
        }

        public static void AddOffice(this ContainerBuilder builder)
        {
            OfficeOptions options = new OfficeOptions()
            {
            };

            AddOfficeServices(builder, options);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="options"></param>
        internal static void AddOfficeServices(ContainerBuilder builder, OfficeOptions options)
        {
            builder.RegisterType<ExcelImportService>().AsImplementedInterfaces();
            builder.RegisterType<ExcelExportService>().AsImplementedInterfaces();
            builder.RegisterType<ExcelImportSolutionService>().AsImplementedInterfaces();
            builder.RegisterType<WordExportService>().AsImplementedInterfaces();

            //根据配置项动态注入Provider
            if (options.ExcelImportSolution == SolutionEnum.NPOI)
            {
                builder.RegisterType<ExcelImportProvider>().AsImplementedInterfaces();
            }

            if (options.ExcelExportSolution == SolutionEnum.NPOI)
            {
                builder.RegisterType<ExcelExportProvider>().AsImplementedInterfaces();
            }

            if (options.WordExportSolution == SolutionEnum.NPOI)
            {
                builder.RegisterType<WordExportProvider>().AsImplementedInterfaces();
            }
        }
    }
}