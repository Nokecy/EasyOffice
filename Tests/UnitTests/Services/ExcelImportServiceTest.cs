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
    public class ExcelImportServiceTest : IDisposable
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
        public void FluentApi测试()
        {
            Validator<ExcelCarTemplateDTO> v = new Validator<ExcelCarTemplateDTO>();
            v.AddRule(x => x.Age).NotDuplicate();
        }

        public void Dispose()
        {
            _rows = null;
        }
    }
}
