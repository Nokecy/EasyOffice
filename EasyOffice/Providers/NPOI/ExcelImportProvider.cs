using EasyOffice.Attributes;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace EasyOffice.Providers.NPOI
{
    public class ExcelImportProvider : IExcelImportProvider
    {
        private static Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// 将Sheet转化为ExcelDataRow集合
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startDataRowIndex"></param>
        /// <returns></returns>
        public ExcelData Convert<TTemplate>(string fileUrl
              , int sheetIndex = 0, int headerRowIndex = 0, int dataRowStartIndex = 1)
              where TTemplate : class, new()
        {
            var workbook = GetWorkbook(fileUrl);

            var sheet = GetSheet(workbook, sheetIndex);

            var headerRow = GetHeaderRow(sheet, headerRowIndex);

            List<ExcelDataRow> dataRows = new List<ExcelDataRow>();
            for (int i = dataRowStartIndex; i < sheet.PhysicalNumberOfRows; i++)
            {
                var row = sheet.GetRow(i);
                if (row == null || !row.Cells.Any(x => !string.IsNullOrWhiteSpace(x?.GetStringValue()))) continue;
                dataRows.Add(Convert<TTemplate>(row, headerRow));
            }

            return new ExcelData()
            {
                Header = headerRow,
                Data = dataRows
            };
        }

        private ExcelHeaderRow GetHeaderRow(ISheet sheet, int headerRowIndex)
        {
            var headerRow = new ExcelHeaderRow();
            IRow row = sheet.GetRow(headerRowIndex);
            ICell cell;
            for (int i = 0; i < row.PhysicalNumberOfCells; i++)
            {
                cell = row.GetCell(i);
                headerRow.Cells.Add(
                    new ExcelCol()
                    {
                        ColIndex = i,
                        ColName = cell.GetStringValue()
                    });
            }

            return headerRow;
        }

        private IWorkbook GetWorkbook(string fileUrl)
        {
            if (string.IsNullOrWhiteSpace(fileUrl))
            {
                throw new ArgumentNullException("fileUrl");
            }

            if (!File.Exists(fileUrl))
            {
                throw new FileNotFoundException();
            }

            string ext = Path.GetExtension(fileUrl).ToLower().Trim();
            if (ext != ".xls" && ext != ".xlsx")
            {
                throw new NotSupportedException("仅支持后缀名为.xls或者.xlsx的文件");
            }

            return WorkbookFactory.Create(fileUrl);
        }

        private ISheet GetSheet(IWorkbook workbook, int sheetIndex = 0)
        {
            return workbook.GetSheetAt(sheetIndex);
        }

        /// <summary>
        /// 将IRow转换为ExcelDataRow
        /// </summary>
        /// <typeparam name="TTemplate"></typeparam>
        /// <param name="row"></param>
        /// <param name="headerRow"></param>
        /// <returns></returns>
        public static ExcelDataRow Convert<TTemplate>(IRow row, ExcelHeaderRow headerRow)
        {
            Type type = typeof(TTemplate);
            var props = type.GetProperties().ToList();
            ExcelDataRow dataRow = new ExcelDataRow()
            {
                DataCols = new List<ExcelDataCol>(),
                ErrorMsg = string.Empty,
                IsValid = true,
                RowIndex = row.RowNum
            };

            ExcelDataCol dataCol;
            string colName;
            string propertyName;
            string key;
            for (int i = 0; i < headerRow.Cells.Count; i++)
            {
                colName = headerRow?.Cells?.SingleOrDefault(h => h.ColIndex == i)?.ColName;

                if (colName == null)
                {
                    continue;
                }

                key = $"{type.FullName}_{i}";

                if (Table[key] == null)
                {
                    //优先匹配ColName特性值
                    var matchProperty = props.FirstOrDefault(p => p.GetCustomAttribute<ColNameAttribute>()?.ColName == colName);

                    if (matchProperty == null)
                    {
                        //次之匹配属性名
                        matchProperty = props.FirstOrDefault(p => p.Name.Equals(colName,StringComparison.CurrentCultureIgnoreCase));
                    }

                    propertyName = matchProperty?.Name;

                    Table[key] = propertyName;
                }

                dataCol = new ExcelDataCol()
                {
                    ColIndex = i,
                    ColName = colName,
                    PropertyName = Table[key]?.ToString(),
                    RowIndex = row.RowNum,
                    ColValue = row.GetCell(i) == null ? string.Empty : row.GetCell(i).GetStringValue()
                };

                dataRow.DataCols.Add(dataCol);
            }

            return dataRow;
        }
    }
}
