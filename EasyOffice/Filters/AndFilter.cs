using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Filters
{
    /// <summary>
    /// 且 过滤器
    /// </summary>
    public class AndFilter : IFilter
    {
        public List<IFilter> filters { get; set; }

        public Type AttributeType => null;

        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context, ImportOption importOption)
        {
            foreach (var filter in filters)
            {
                excelDataRows = filter.Filter(excelDataRows, context, importOption);
            }

            return excelDataRows;
        }
    }
}
