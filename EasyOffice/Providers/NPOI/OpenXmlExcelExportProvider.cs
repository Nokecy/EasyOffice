using DocumentFormat.OpenXml.Packaging;
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
using System.Text;

namespace EasyOffice.Providers.NPOI
{
    public class OpenXmlExcelExportProvider : IExcelExportProvider
    {
        public byte[] DecorateHeader<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context) where T : class, new()
        {
            throw new NotImplementedException();
        }

        public byte[] Export<T>(ExportOption<T> exportOption) where T : class, new()
        {
            //var dt = ToDataTable(exportOption.Data);
            //DataSet ds = new DataSet();

            //ds.Tables.Add(dt);

            var headerDict = ExportMappingDictFactory.CreateInstance(typeof(T));

            return Export(exportOption.Data, headerDict);
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

        private byte[] Export<T>(IEnumerable<T> data, Dictionary<string, string> headerDict)
            where T : class, new()
        {
            byte[] result = null;
            var ms = new MemoryStream();
            var fileUrl = Path.Combine(Environment.CurrentDirectory, DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xlsx");
            using (var workbook = SpreadsheetDocument.Create(fileUrl, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
                //(destination, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
                )
            {
                var type = typeof(T);
                var props = type.GetProperties();

                var workbookPart = workbook.AddWorkbookPart();

                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();


                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                {
                    sheetId =
                        sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = "sheet1" };
                sheets.Append(sheet);

                DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                List<String> columns = new List<string>();
                foreach (var kvp in headerDict)
                {
                    columns.Add(kvp.Value);

                    DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                    cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(kvp.Value);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (var dto in data)
                {
                    DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    foreach (String col in columns)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;

                        var propName = headerDict.FirstOrDefault(x => x.Value == col).Key;

                        var prop = props.FirstOrDefault(x => x.Name == propName);

                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(
                            prop.GetValue(dto).ToString()
                            ); //
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }
            }

            result = File.ReadAllBytes(fileUrl);

            File.Delete(fileUrl);

            return result;
        }
    }
}
