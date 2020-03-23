using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EasyOffice.Decorators;
using EasyOffice.Factories;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace EasyOffice.Providers.NPOI
{
    public class OpenXmlExcelExportProvider : IExcelExportProvider
    {
        public byte[] DecorateHeader<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context) where T : class, new()
        {
            throw new NotImplementedException();
        }

        public byte[] Export<T>(List<T> data,ExportOption<T> exportOption) where T : class, new()
        {
            //var dt = ToDataTable(exportOption.Data);
            //DataSet ds = new DataSet();

            //ds.Tables.Add(dt);

            var headerDict = ExportMappingDictFactory.CreateInstance(typeof(T));

            return Export(data, headerDict);
        }

        public byte[] MergeCols<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context) where T : class, new()
        {
            throw new NotImplementedException();
        }

        public byte[] WrapText<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context) where T : class, new()
        {
            throw new NotImplementedException();
        }

        public static DataTable ToDataTable<T>(IEnumerable<T> collection)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray());
            if (collection.Count() > 0)
            {
                for (int i = 0; i < collection.Count(); i++)
                {
                    ArrayList tempList = new ArrayList();
                    foreach (PropertyInfo pi in props)
                    {
                        object obj = pi.GetValue(collection.ElementAt(i), null);
                        tempList.Add(obj);
                    }
                    object[] array = tempList.ToArray();
                    dt.LoadDataRow(array, true);
                }
            }
            return dt;
        }

        private byte[] Export<T>(IEnumerable<T> datas, Dictionary<string, string> headerDict)
            where T : class, new()
        {
            byte[] result = null;
            var ms = new MemoryStream();
            var fileUrl = Path.Combine(Environment.CurrentDirectory, DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xlsx");
            using (var workbook = SpreadsheetDocument.Create(fileUrl, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
                )
            {
                var type = typeof(T);
                var props = type.GetProperties();

                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new Sheets();


                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                sheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = "sheet1" };
                sheets.Append(sheet);

                Row headerRow = new Row();

                List<string> columns = new List<string>();
                foreach (var kvp in headerDict)
                {
                    columns.Add(kvp.Value);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(kvp.Value);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (var item in datas)
                {
                    Row newRow = new Row();
                    foreach (string col in columns)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;

                        var propName = headerDict.FirstOrDefault(x => x.Value == col).Key;

                        var prop = props.FirstOrDefault(x => x.Name == propName);

                        cell.CellValue = new CellValue(
                            prop.GetValue(item).ToString()
                        );

                        newRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(newRow);
                }

                //var options = new ParallelOptions()
                //{
                //    MaxDegreeOfParallelism = 8,
                //};

                //ConcurrentBag<Row> rows = new ConcurrentBag<Row>();

                //Parallel.ForEach(datas, options,data => {
                //    Row newRow = new Row();
                //    foreach (string col in columns)
                //    {
                //        Cell cell = new Cell();
                //        cell.DataType = CellValues.String;

                //        var propName = headerDict.FirstOrDefault(x => x.Value == col).Key;

                //        var prop = props.FirstOrDefault(x => x.Name == propName);

                //        cell.CellValue = new CellValue(
                //            col
                //        );

                //        newRow.AppendChild(cell);
                //    }

                //    rows.Add(newRow);
                //});
                //sheetData.Append(rows);
            }

            result = File.ReadAllBytes(fileUrl);

            try
            {
                var tf = new TaskFactory();
                tf.StartNew(() =>
                {
                    File.Delete(fileUrl);
                });
            }
            catch
            {
            }

            return result;
        }
    }
}
