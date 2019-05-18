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

namespace Bayantu.Extensions.DependencyInjection
{
    /// <summary>
    /// Office基础库依赖注入
    /// </summary>
    public static class OfficeServiceCollectionExtensions
    {
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

            const string serviceName = "Office";

            builder.Register(c => new ImportErrorLogRepository(c.Resolve<OfficeDbContext>(), c.Resolve<IEvosSession>(), c.ResolveNamed<IDapperAsyncExecutor>($"{serviceName}_IDapperAsyncExecutor")))
                .As<IImportErrorLogRepository>();

            //Session
            builder.AddEvosSession();
        }
    }
}