using EasyOffice.Attributes;
using EasyOffice.Enums;
using EasyOffice.Filters;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using EasyOffice.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EasyOffice.Filters
{
    /// <summary>
    /// 最大长度过滤器
    /// </summary>
    [FilterBind(typeof(MaxLengthAttribute))]
    public class MaxLengthFilter : BaseFilter,IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context, ImportOption importOption)
        {
            foreach (var r in excelDataRows)
            {
                if (!r.IsValid && importOption.ValidateMode == ValidateModeEnum.StopOnFirstFailure)
                    continue;

                r.DataCols.ForEach(c =>
                {
                    var attr = c.GetFilterAttr<MaxLengthAttribute>(context.TypeFilterInfo);
                    if (attr != null)
                    {
                        r.SetNotValid(c.ColValue.Length <= attr.MaxLength, c, attr.ErrorMsg);
                    }
                });
            }

            return excelDataRows;
        }
        public int MaxLength { get; set; }
    }
}
