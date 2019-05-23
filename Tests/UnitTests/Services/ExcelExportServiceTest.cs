using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using EasyOffice.Providers.NPOI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UnitTests.DependencyInjection;
using UnitTests.Models;
using Xunit;

namespace UnitTests.Services
{
    public class ExcelExportServiceTest : IDisposable
    {
        private readonly IExcelExportService _excelExportService;
        public ExcelExportServiceTest()
        {
            _excelExportService = _excelExportService.Resolve();
        }

        [Fact]
        public async Task ExportTest_导出Excel_正确导出Excel手动观察()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls");

            var carDTO = new ExcelCarTemplateDTO()
            {
                Age = 10,
                CarCode = "鄂A123456",
                Gender = GenderEnum.男,
                IdentityNumber = "test",
                Mobile = "test",
                Name = "test",
                RegisterDate = DateTime.Now
            };

            var list = new List<ExcelCarTemplateDTO>();

            for (int i = 0; i < 10; i++)
            {
                list.Add(carDTO);
            }

            var bytes = await _excelExportService.ExportAsync<ExcelCarTemplateDTO>(new ExportOption<ExcelCarTemplateDTO>()
            {
                Data = list
            });
            File.WriteAllBytes(fileUrl, bytes);
        }

        public const int dataCount = 10000;

        [Fact]
        public async Task ExportTest_NPOI最大条数性能测试_性能测试()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

            var performance = new PerformanceTest();

            var list = new List<PerformanceTest>();

            for (int i = 0; i < dataCount; i++)
            {
                list.Add(performance);
            }

            var bytes = await _excelExportService.ExportAsync(new ExportOption<PerformanceTest>()
            {
                Data = list
            });

            File.WriteAllBytes(fileUrl, bytes);
        }

        [Fact]
        public async Task ExportTest_OpenXml最大条数性能测试_性能测试()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

            var performance = new PerformanceTest();

            var list = new List<PerformanceTest>();

            for (int i = 0; i < dataCount; i++)
            {
                list.Add(performance);
            }

            var bytes = await _excelExportService.ExportAsync(new ExportOption<PerformanceTest>()
            {
                Data = list,
                ExportType = EasyOffice.Enums.ExportType.FastXLSX
            });

            File.WriteAllBytes(fileUrl, bytes);
        }

        [Fact]
        public async Task ExportTest_CSV最大条数性能测试_性能测试()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");

            var performance = new PerformanceTest();

            var list = new List<PerformanceTest>();

            for (int i = 0; i < dataCount; i++)
            {
                list.Add(performance);
            }

            var bytes = await _excelExportService.ExportAsync(new ExportOption<PerformanceTest>()
            {
                Data = list,
                ExportType = EasyOffice.Enums.ExportType.CSV
            });

            File.WriteAllBytes(fileUrl, bytes);
        }

        public void Dispose()
        {
        }
    }
}
