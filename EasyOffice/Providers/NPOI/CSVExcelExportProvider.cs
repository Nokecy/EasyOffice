using CsvHelper;
using EasyOffice.Decorators;
using EasyOffice.Factories;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EasyOffice.Providers.NPOI
{
    public class CSVExcelExportProvider : IExcelExportProvider
    {
        public byte[] DecorateHeader<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context) where T : class, new()
        {
            throw new NotImplementedException();
        }

        public byte[] Export<T>(List<T> data,ExportOption<T> exportOption) where T : class, new()
        {
            string url = Path.Combine(Environment.CurrentDirectory, DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xlsx");
            using (var writer = new StreamWriter(url))
            using (var csv = new CsvWriter(writer))
            {
                csv.WriteRecords(data);
            }

            var bytes = File.ReadAllBytes(url);

            File.Delete(url);

            return bytes;
        }

        public byte[] MergeCols<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context) where T : class, new()
        {
            throw new NotImplementedException();
        }

        public byte[] WrapText<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context) where T : class, new()
        {
            throw new NotImplementedException();
        }
    }
}
