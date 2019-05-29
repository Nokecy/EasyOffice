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
    /// 正则表达式过滤器
    /// </summary>
    [FilterBind(typeof(RegexAttribute))]
    public class RegexFilter : BaseFilter,IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context, ImportOption importOption)
        {
            foreach (var r in excelDataRows)
            {
                if (!r.IsValid && importOption.ValidateMode == ValidateModeEnum.StopOnFirstFailure)
                    continue;

                r.DataCols.ForEach(c =>
                {
                    var attrs = c.GetFilterAttrs<RegexAttribute>(context.TypeFilterInfo);

                    if (attrs != null && attrs.Count > 0)
                    {
                        attrs.ForEach(a =>
                        {
                            r.SetNotValid(Regex.IsMatch(c.ColValue, a.RegexString), c, a.ErrorMsg);
                        });
                    }
                });
            }

            return excelDataRows;
        }

        public string RegexString { get; set; }
    }
}
