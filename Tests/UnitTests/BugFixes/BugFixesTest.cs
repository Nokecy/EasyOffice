using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using EasyOffice.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UnitTests.DependencyInjection;
using UnitTests.Models;
using UnitTests.Models.bugs;
using Xunit;

namespace UnitTests.Services
{
    public class BugFixesTest :IDisposable
    {
        private readonly IExcelImportService _excelImportService;
        public BugFixesTest()
        {
            _excelImportService = _excelImportService.Resolve();
        }


        [Fact]
        public async Task Issue3_行数读取错误导致row为null错误()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, "Resources/bugs", "issue3.xlsx");

            var rows = await _excelImportService.ValidateAsync<PerformanceTest>(new ImportOption()
            {
                FileUrl = fileUrl
            });
        }

        [Fact]
        public async Task issue6_FastConvert转换数据_数据准确()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, "Resources/bugs", "issue6.xlsx");

            var rows = await _excelImportService.ValidateAsync<Issue6>(new ImportOption()
            {
                FileUrl = fileUrl
            });

            var data = rows.Convert<Issue6>().ToList();

            Assert.Equal(2, data.Count);
            Assert.Equal("name1", data[0].Name);
            Assert.Equal("name2", data[1].Name);
        }

        public void Dispose()
        {
        }
    }
}
