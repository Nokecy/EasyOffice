using EasyOffice.Attributes;
using EasyOffice.Enums;
using EasyOffice.Filters;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using EasyOffice.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace EasyOffice.Filters
{
    /// <summary>
    /// 数值范围过滤器
    /// </summary>
    [FilterBind(typeof(RangeAttribute))]
    public class RangeFilter : BaseFilter, IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context, ImportOption importOption)
        {
            foreach (var r in excelDataRows)
            {
                if (!r.IsValid && importOption.ValidateMode == ValidateModeEnum.StopOnFirstFailure)
                    continue;

                r.DataCols.ForEach(c =>
                {
                    var attr = c.GetFilterAttr<RangeAttribute>(context.TypeFilterInfo);
                    if (attr != null)
                    {
                        r.SetNotValid(c.IsInRange(attr.Max, attr.Min), c, attr.ErrorMsg);
                    }
                });
            }

            return excelDataRows;
        }

        public int Min { get; set; }

        public int Max { get; set; }
    }
}
