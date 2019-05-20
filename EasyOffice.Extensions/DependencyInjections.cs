using DinkToPdf;
using DinkToPdf.Contracts;
using EasyOffice.Extension.Converter.Converters;
using EasyOffice.Extensions.Interfaces;
using Microsoft.Extensions.DependencyInjection;

namespace EasyOffice.Extensions
{
    /// <summary>
    /// Office基础库依赖注入
    /// </summary>
    public static class DependencyInjections
    {
        public static void AddEasyOfficeExtensions(this IServiceCollection services)
        {
            services.AddTransient<IWordConverter, WordConverter>();

            services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));
        }
    }
}