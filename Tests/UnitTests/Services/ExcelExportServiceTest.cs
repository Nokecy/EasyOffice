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

            var bytes = await _excelExportService.ExportAsync<ExcelCarTemplateDTO>(list);
            File.WriteAllBytes(fileUrl, bytes);
        }

        public void Dispose()
        {
        }
    }
}
