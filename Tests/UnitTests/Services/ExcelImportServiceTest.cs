using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
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
    public class ExcelImportServiceTest :IDisposable
    {
        private readonly IExcelImportService _excelImportService;
        private static List<ExcelDataRow> _rows;
        public ExcelImportServiceTest()
        {
            _excelImportService = _excelImportService.Resolve();

            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, "Resources", "CarImport.xlsx");
            _rows = _excelImportService.ValidateAsync<ExcelCarTemplateDTO>(new ImportOption() { FileUrl = fileUrl }).Result;
        }

        [Fact]
        public void ValidateTest_导入Excel_数据行数等于10()
        {
            Assert.True(_rows.Count == 10);
        }

        [Fact]
        public void ValidateTest_导入Excel_车牌号非法校验()
        {
            var row0 = _rows[0];
            Assert.True(!row0.IsValid && row0.ErrorMsg.Contains("车牌号非法"));
        }

        [Fact]
        public void ValidateTest_导入Excel_姓名超长校验()
        {
            var row1 = _rows[1];
            Assert.True(!row1.IsValid && row1.ErrorMsg.Contains("姓名超长"));
        }

        [Fact]
        public void ValidateTest_导入Excel_枚举校验()
        {
            var row2 = _rows[2];
            Assert.True(!row2.IsValid && row2.ErrorMsg.Contains("性别非法"));
        }

        [Fact]
        public void ValidateTest_导入Excel_日期校验()
        {
            var row3 = _rows[3];
            Assert.True(!row3.IsValid && row3.ErrorMsg.Contains("注册日期非法"));
        }

        [Fact]
        public void ValidateTest_导入Excel_数值范围校验()
        {
            var row4 = _rows[4];
            Assert.True(!row4.IsValid && row4.ErrorMsg.Contains("年龄超限，仅允许为0-150"));
        }

        [Fact]
        public void ValidateTest_导入Excel_重复值校验()
        {
            var row5 = _rows[5];
            Assert.True(!row5.IsValid && row5.ErrorMsg.Contains("车牌号重复"));
        }

        [Fact]
        public void ValidateTest_导入Excel_必填校验()
        {
            var row6 = _rows[6];
            Assert.True(!row6.IsValid && row6.ErrorMsg.Contains("车牌号必填"));
        }

        [Fact]
        public void ValidateTest_导入Excel_手机号校验()
        {
            var row7 = _rows[7];
            Assert.True(!row7.IsValid && row7.ErrorMsg.Contains("手机号非法"));
        }

        [Fact]
        public void ValidateTest_导入Excel_国内身份证校验校验()
        {
            var row8 = _rows[8];
            Assert.True(!row8.IsValid && row8.ErrorMsg.Contains("身份证号非法"));
        }

        [Fact]
        public void ValidateTest_导入Excel_有效数据()
        {
            var row9 = _rows[9];
            Assert.True(row9.IsValid);
        }

        [Fact]
        public void ValidateTest_导入Excel_有效数据转换正确()
        {
            var row9 = _rows[9];
            Assert.True(row9.IsValid);

            ExcelCarTemplateDTO dto = row9.Convert<ExcelCarTemplateDTO>();

            Assert.True(dto.CarCode == "鄂A57MG2"
                && dto.Gender == GenderEnum.男
                && dto.Gender.GetHashCode() == 10
                && dto.Name == "龚英韬"
                && dto.RegisterDate == new DateTime(2018, 1, 1)
                && dto.Age == 18);
        }

        [Fact]
        public async Task ConvertTest_Convert转换数据_数据准确()
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

        [Fact]
        public void FastConvert_性能测试_n条()
        {
            var rows = DataLib.GetExcelDataRows(5000, 10);

            var result = rows.FastConvert<PerformanceTest>().ToList();

            for (int i = 0; i < result.Count; i++)
            {
                Assert.Equal($"r{i + 1}c1", result[i].p1);
                Assert.Equal($"r{i + 1}c2", result[i].p2);
                Assert.Equal($"r{i + 1}c3", result[i].p3);
                Assert.Equal($"r{i + 1}c4", result[i].p4);
                Assert.Equal($"r{i + 1}c5", result[i].p5);
                Assert.Equal($"r{i + 1}c6", result[i].p6);
                Assert.Equal($"r{i + 1}c7", result[i].p7);
                Assert.Equal($"r{i + 1}c8", result[i].p8);
                Assert.Equal($"r{i + 1}c9", result[i].p9);
                Assert.Equal($"r{i + 1}c10", result[i].p10);
            }
        }

        [Fact]
        public void Convert_性能测试_n条()
        {
            var rows = DataLib.GetExcelDataRows(5000, 10);

            var result = rows.Convert<PerformanceTest>().ToList();

            for (int i = 0; i < result.Count; i++)
            {
                Assert.Equal($"r{i + 1}c1", result[i].p1);
                Assert.Equal($"r{i + 1}c2", result[i].p2);
                Assert.Equal($"r{i + 1}c3", result[i].p3);
                Assert.Equal($"r{i + 1}c4", result[i].p4);
                Assert.Equal($"r{i + 1}c5", result[i].p5);
                Assert.Equal($"r{i + 1}c6", result[i].p6);
                Assert.Equal($"r{i + 1}c7", result[i].p7);
                Assert.Equal($"r{i + 1}c8", result[i].p8);
                Assert.Equal($"r{i + 1}c9", result[i].p9);
                Assert.Equal($"r{i + 1}c10", result[i].p10);
            }
        }

        [Fact]
        public async Task FastConvertTest_Convert转换数据_数据准确()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, "Resources/bugs", "issue6.xlsx");

            var rows = await _excelImportService.ValidateAsync<Issue6>(new ImportOption()
            {
                FileUrl = fileUrl
            });

            var data = rows.FastConvert<Issue6>().ToList();

            Assert.Equal(2, data.Count);
            Assert.Equal("name1", data[0].Name);
            Assert.Equal("name2", data[1].Name);
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

        public void Dispose()
        {
            _rows = null;
        }
    }
}
