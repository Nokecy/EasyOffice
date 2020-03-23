using EasyOffice.Attributes;
using EasyOffice.Decorators;
using EasyOffice.Enums;
using EasyOffice.Factories;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using EasyOffice.Providers.NPOI;
using EasyOffice.Utils;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EasyOffice.Providers.NPOI
{
    public class ExcelExportProvider : IExcelExportProvider
    {
        public byte[] Export<T>(List<T> data,ExportOption<T> exportOption)
            where T : class, new()
        {
            IWorkbook workbook = null;
            if (exportOption.ExportType == ExportType.XLS)
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                workbook = new XSSFWorkbook();
            }

            ISheet sheet = workbook.CreateSheet(exportOption.SheetName);

            var headerDict = ExportMappingDictFactory.CreateInstance(typeof(T));

            SetHeader<T>(sheet, exportOption.HeaderRowIndex, headerDict);

            if (data != null && data.Count > 0)
            {
                SetDataRows(sheet, exportOption.DataRowStartIndex, data, headerDict);
            }

            return workbook?.ToBytes();
        }

        private void SetHeader<T>(ISheet sheet, int headerRowIndex, Dictionary<string, string> headerDict)
      where T : class, new()
        {
            IRow row = sheet.CreateRow(headerRowIndex);
            int colIndex = 0;
            foreach (var kvp in headerDict)
            {
                row.CreateCell(colIndex).SetCellValue(kvp.Value);
                colIndex++;
            }
        }

        private void SetDataRows<T>(ISheet sheet, int dataRowStartIndex, List<T> datas, Dictionary<string, string> headerDict)
                where T : class, new()
        {
            if (datas.Count <= 0) return;

            for (var i = 0; i < datas.Count; i++)
            {
                int colIndex = 0;
                IRow row = sheet.CreateRow(dataRowStartIndex + i);
                T dto = datas[i];

                foreach (var kvp in headerDict)
                {
                    row.CreateCell(colIndex).SetCellValue(dto.GetStringValue(kvp.Key));
                    colIndex++;
                }
            }
        }

        public byte[] DecorateHeader<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context)
            where T : class, new()
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            var attr = (HeaderAttribute)context.TypeDecoratorInfo?.TypeDecoratorAttrs?.SingleOrDefault(a => a.GetType() == typeof(HeaderAttribute));

            if (attr == null)
            {
                return workbookBytes;
            }

            IWorkbook workbook = workbookBytes.ToWorkbook();

            IRow headerRow = workbook?.GetSheet(exportOption.SheetName)?.GetRow(exportOption.HeaderRowIndex);

            if (headerRow == null) return workbookBytes;

            ICellStyle style = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont();

            font.FontName = attr.FontName;
            font.Color = (short)attr.Color.GetHashCode();
            font.FontHeightInPoints = (short)attr.FontSize;
            if (attr.IsBold)
            {
                font.Boldweight = short.MaxValue;
            }

            style.SetFont(font);

            for (int i = 0; i < headerRow.PhysicalNumberOfCells; i++)
            {
                headerRow.GetCell(i).CellStyle = style;
            }

            return workbook.ToBytes();
        }

        public byte[] MergeCols<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context)
            where T : class, new()
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            var propertyDecoratorInfos = context.TypeDecoratorInfo.PropertyDecoratorInfos;
            var workbook = workbookBytes.ToWorkbook();
            ISheet sheet = workbook.GetSheet(exportOption.SheetName);
            foreach (var item in propertyDecoratorInfos)
            {
                if (item.DecoratorAttrs.SingleOrDefault(a => a.GetType() == typeof(MergeColsAttribute)) != null)
                {
                    MergeCols(sheet, item.ColIndex, exportOption);
                }
            }

            return workbook.ToBytes();
        }

        public byte[] WrapText<T>(byte[] workbookBytes, ExportOption<T> exportOption, DecoratorContext context)
            where T : class, new()
        {
            if (context == null)
            {
                throw new ArgumentNullException("context");
            }

            var attr = context.TypeDecoratorInfo.GetDecorateAttr<WrapTextAttribute>();

            if (attr == null)
            {
                return workbookBytes;
            }

            IWorkbook workbook = workbookBytes.ToWorkbook();

            ISheet sheet = workbook.GetSheet(exportOption.SheetName);
            IRow row;
            if (sheet.PhysicalNumberOfRows > 0)
            {
                for (int i = 0; i < sheet.PhysicalNumberOfRows; i++)
                {
                    row = sheet.GetRow(i);
                    for (int colIndex = 0; colIndex < row.PhysicalNumberOfCells; colIndex++)
                    {
                        row.GetCell(colIndex).CellStyle.WrapText = true;
                    }
                }
            }

            return workbook.ToBytes();
        }

        private void MergeCols<T>(ISheet sheet, int colIndex, ExportOption<T> exportOption)
            where T : class, new()
        {
            string currentCellValue;
            int startRowIndex = exportOption.DataRowStartIndex;
            CellRangeAddress mergeRangeAddress;

            var startRow = sheet.GetRow(startRowIndex);
            if (startRow == null) return;

            var startCell = startRow.GetCell(colIndex);
            if (startCell == null) return;

            string startCellValue = startCell.StringCellValue;

            if (string.IsNullOrWhiteSpace(startCellValue)) return;

            for (int rowIndex = exportOption.DataRowStartIndex; rowIndex < sheet.PhysicalNumberOfRows; rowIndex++)
            {
                var cell = sheet.GetRow(rowIndex)?.GetCell(colIndex);
                currentCellValue = cell == null ? string.Empty : cell.StringCellValue;

                if (currentCellValue.Trim() != startCellValue.Trim())
                {
                    mergeRangeAddress = new CellRangeAddress(startRowIndex, rowIndex - 1, colIndex, colIndex);
                    sheet.AddMergedRegion(mergeRangeAddress);

                    startRow.GetCell(colIndex).CellStyle.VerticalAlignment = VerticalAlignment.Center;

                    startRowIndex = rowIndex;
                    startCellValue = currentCellValue;
                }

                if (rowIndex == sheet.PhysicalNumberOfRows - 1 && startRowIndex != rowIndex)
                {
                    mergeRangeAddress = new CellRangeAddress(startRowIndex, rowIndex, colIndex, colIndex);
                    sheet.AddMergedRegion(mergeRangeAddress);

                    startRow.GetCell(colIndex).CellStyle.VerticalAlignment = VerticalAlignment.Center;
                }
            }
        }


    }
}
