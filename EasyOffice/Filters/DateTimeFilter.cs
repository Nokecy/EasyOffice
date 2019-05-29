using EasyOffice.Attributes;
using EasyOffice.Enums;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using EasyOffice.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EasyOffice.Filters
{
    /// <summary>
    /// 日期过滤器
    /// </summary>
    [FilterBind(typeof(DateTimeAttribute))]
    public class DateTimeFilter : BaseFilter, IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context, ImportOption importOption)
        {
            foreach (var r in excelDataRows)
            {
                if (!r.IsValid && importOption.ValidateMode == ValidateModeEnum.StopOnFirstFailure)
                    continue;

                r.DataCols.ForEach(c =>
                {
                    var attr = c.GetFilterAttr<DateTimeAttribute>(context.TypeFilterInfo);
                    if (attr != null)
                    {
                        r.SetNotValid(c.IsDateTime(), c, attr.ErrorMsg);
                    }
                });
            }

            return excelDataRows;
        }
    }
}
