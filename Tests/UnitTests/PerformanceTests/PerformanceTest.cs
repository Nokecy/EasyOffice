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
    public class PerformanceTest : IDisposable
    {
        private readonly IExcelImportService _excelImportService;
        private readonly IExcelExportService _excelExportService;
        private static readonly int rowsCount = 1000000;
        private static List<ExcelDataRow> _rows = new  List<ExcelDataRow>();

        private static List<PerformanceTestDTO> _datas=new List<PerformanceTestDTO>();

        public PerformanceTest()
        {
            _excelImportService = _excelImportService.Resolve();
            _excelExportService = _excelExportService.Resolve();

            _rows = DataLib.GetExcelDataRows(rowsCount, 10);

            var performance = new PerformanceTestDTO();
            for (int i = 0; i < rowsCount; i++)
            {
                _datas.Add(performance);
            }
        }


        [Fact]
        public void Convert_转换_n条()
        {
            var result = _rows.Convert<PerformanceTestDTO>().ToList();

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

        //[Fact]
        //public void Convert_反射委托转换测试_n条()
        //{
        //    var rows = DataLib.GetExcelDataRows(rowsCount, 10);

        //    var result = _rows.Convert<PerformanceTestDTO>().ToList();

        //    for (int i = 0; i < result.Count; i++)
        //    {
        //        Assert.Equal($"r{i + 1}c1", result[i].p1);
        //        Assert.Equal($"r{i + 1}c2", result[i].p2);
        //        Assert.Equal($"r{i + 1}c3", result[i].p3);
        //        Assert.Equal($"r{i + 1}c4", result[i].p4);
        //        Assert.Equal($"r{i + 1}c5", result[i].p5);
        //        Assert.Equal($"r{i + 1}c6", result[i].p6);
        //        Assert.Equal($"r{i + 1}c7", result[i].p7);
        //        Assert.Equal($"r{i + 1}c8", result[i].p8);
        //        Assert.Equal($"r{i + 1}c9", result[i].p9);
        //        Assert.Equal($"r{i + 1}c10", result[i].p10);
        //    }
        //}

        //[Fact]
        //public void ConvertByRefelection_反射转换测试_n条()
        //{
        //    var rows = DataLib.GetExcelDataRows(rowsCount, 10);

        //    var result = _rows.ConvertByRefelection<PerformanceTestDTO>().ToList();

        //    for (int i = 0; i < result.Count; i++)
        //    {
        //        Assert.Equal($"r{i + 1}c1", result[i].p1);
        //        Assert.Equal($"r{i + 1}c2", result[i].p2);
        //        Assert.Equal($"r{i + 1}c3", result[i].p3);
        //        Assert.Equal($"r{i + 1}c4", result[i].p4);
        //        Assert.Equal($"r{i + 1}c5", result[i].p5);
        //        Assert.Equal($"r{i + 1}c6", result[i].p6);
        //        Assert.Equal($"r{i + 1}c7", result[i].p7);
        //        Assert.Equal($"r{i + 1}c8", result[i].p8);
        //        Assert.Equal($"r{i + 1}c9", result[i].p9);
        //        Assert.Equal($"r{i + 1}c10", result[i].p10);
        //    }
        //}

        [Fact]
        public void ConvertHadrCode_硬编码转换_n条()
        {
            var rows = DataLib.GetExcelDataRows(rowsCount, 10);

            var result = new List<PerformanceTestDTO>();

            foreach (var item in rows)
            {
                var data = new PerformanceTestDTO()
                {
                    p1 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p1")?.ColValue,
                    p2 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p2")?.ColValue,
                    p3 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p3")?.ColValue,
                    p4 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p4")?.ColValue,
                    p5 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p5")?.ColValue,
                    p6 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p6")?.ColValue,
                    p7 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p7")?.ColValue,
                    p8 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p8")?.ColValue,
                    p9 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p9")?.ColValue,
                    p10 = item.DataCols.FirstOrDefault(x => x.PropertyName == "p10")?.ColValue,
                };
            }

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
        public async Task ExportTest_NPOI最大条数性能测试_性能测试()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

            var bytes = await _excelExportService.ExportAsync(new ExportOption<PerformanceTestDTO>()
            {
                Data = _datas
            });

            File.WriteAllBytes(fileUrl, bytes);
        }

        [Fact]
        public async Task ExportTest_OpenXml最大条数性能测试_性能测试()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx");

            var bytes = await _excelExportService.ExportAsync(new ExportOption<PerformanceTestDTO>()
            {
                Data = _datas,
                ExportType = EasyOffice.Enums.ExportType.FastXLSX
            });

            File.WriteAllBytes(fileUrl, bytes);
        }

        [Fact]
        public async Task ExportTest_CSV最大条数性能测试_性能测试()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");

            var bytes = await _excelExportService.ExportAsync(new ExportOption<PerformanceTestDTO>()
            {
                Data = _datas,
                ExportType = EasyOffice.Enums.ExportType.CSV
            });

            File.WriteAllBytes(fileUrl, bytes);
        }

        public void Dispose()
        {
        }
    }
}
