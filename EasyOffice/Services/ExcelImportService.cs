using EasyOffice.Factories;
using EasyOffice.Filters;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace EasyOffice.Services
{
    public class ExcelImportService : IExcelImportService
    {
        private readonly IExcelImportProvider _excelImportProvider;

        public ExcelImportService(IExcelImportProvider excelImportProvider)
        {
            _excelImportProvider = excelImportProvider;
        }

        public Task<DataTable> ToTableAsync<T>(string fileUrl
            , int sheetIndex = 0
            , int headerRowIndex = 0
            , int dataRowIndex = 1
            , int dataRowCount = -1)
            where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(fileUrl))
            {
                throw new ArgumentNullException("fieUrl");
            }

            DataTable dt = new DataTable();
            var importOption = new ImportOption()
            {
                FileUrl = fileUrl,
                DataRowStartIndex = dataRowIndex,
                HeaderRowIndex = headerRowIndex,
                SheetIndex = sheetIndex
            };

            var excelData = GetExcelData<T>(importOption);

            excelData.Header.Cells.ForEach(x =>
            {
                dt.Columns.Add(x.ColName);
            });

            var count = excelData.Data.Count;
            if (dataRowCount > -1)
            {
                count = dataRowCount;
            }

            for (int i = 0; i < count && i < excelData.Data.Count; i++)
            {
                var excelRow = excelData.Data[i];
                DataRow dataRow = dt.NewRow();
                dataRow.ItemArray = excelRow.DataCols.Select(c => c.ColValue).ToArray();
                dt.Rows.Add(dataRow);
            }

            return Task.FromResult(dt);
        }

        //public DataTable ToTable(string fileUrl, int sheetIndex = 0, int headerRowIndex = 0, int dataRowStartIndex = 1, int dataRowCount = -1)
        //{
        //    IWorkbook workbook;
        //    using (FileStream stream = new FileStream(fileUrl, FileMode.Open, FileAccess.Read))
        //    {
        //        workbook = new HSSFWorkbook(stream);
        //    }

        //    ISheet sheet = workbook.GetSheetAt(sheetIndex);
        //    DataTable dt = new DataTable(sheet.SheetName);

        //    // 处理表头
        //    IRow headerRow = sheet.GetRow(headerRowIndex);
        //    foreach (ICell headerCell in headerRow)
        //    {
        //        dt.Columns.Add(headerCell.ToString());
        //    }

        //    //默认转换所有数据
        //    var dataRowStopIndex = sheet.PhysicalNumberOfRows - 1;

        //    //只转换指定数量数据
        //    if (dataRowCount > -1)
        //    {
        //        dataRowStopIndex = dataRowStartIndex + dataRowCount;
        //    }

        //    //转换
        //    for (int i = dataRowStartIndex; i < dataRowStopIndex; i++)
        //    {
        //        var row = sheet.GetRow(i);
        //        if (row != null)
        //        {
        //            DataRow dataRow = dt.NewRow();
        //            dataRow.ItemArray = row.Cells.Select(c => c.ToString()).ToArray();
        //            dt.Rows.Add(dataRow);
        //        }
        //    }

        //    return dt;
        //}

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T">模板类</typeparam>
        /// <param name="fileUrl">Excel文件绝对地址</param>
        /// <returns></returns>
        public Task<List<ExcelDataRow>> ValidateAsync<T>(ImportOption importOption) where T : class, new()
        {
            var excelData = GetExcelData<T>(importOption);

            if (importOption.MappingDictionary != null)
            {
                excelData.Data = MappingHeaderDictionary(excelData.Data, importOption.MappingDictionary);
            }

            AndFilter andFilter = new AndFilter() { filters = FiltersFactory.CreateFilters<T>(excelData.Header) };

            FilterContext context = new FilterContext()
            {
                TypeFilterInfo = TypeFilterInfoFactory.CreateInstance(typeof(T), excelData.Header)
            };

            excelData.Data = andFilter.Filter(excelData.Data, context, importOption);

            return Task.FromResult(excelData.Data);
        }

        /// <summary>
        /// 根据表头字段映射数据
        /// </summary>
        /// <param name="rows">Excel原数据</param>
        /// <param name="headerDictionary">映射字典</param>
        /// <returns></returns>
        private List<ExcelDataRow> MappingHeaderDictionary(List<ExcelDataRow> rows, Dictionary<string, string> headerDictionary)
        {
            foreach (var kvp in headerDictionary)
            {
                foreach (var row in rows)
                {
                    foreach (var col in row.DataCols)
                    {
                        if (col.ColName.Equals(kvp.Value, StringComparison.CurrentCultureIgnoreCase))
                        {
                            col.PropertyName = kvp.Key;
                        }
                    }
                }
            }

            return rows;
        }

        private ExcelData GetExcelData<T>(ImportOption importOption)
            where T : class, new()
        {
            return importOption.CustomExcelImportProvider == null ? 
               _excelImportProvider.Convert<T>(
                 importOption.FileUrl
               , importOption.SheetIndex
               , importOption.HeaderRowIndex
               , importOption.DataRowStartIndex
               ) :
               importOption.CustomExcelImportProvider.Convert<T>(
                 importOption.FileUrl
               , importOption.SheetIndex
               , importOption.HeaderRowIndex
               , importOption.DataRowStartIndex
               );
        }
    }
}
