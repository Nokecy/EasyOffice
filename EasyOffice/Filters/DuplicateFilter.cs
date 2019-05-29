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
    /// 重复过滤器
    /// </summary>
    [FilterBind(typeof(DuplicationAttribute))]
    public class DuplicateFilter : BaseFilter,IFilter
    {
        public List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context, ImportOption importOption)
        {
            List<KeyValuePair<int, string>> kvps = new List<KeyValuePair<int, string>>();

            foreach (var r in excelDataRows)
            {
                if (!r.IsValid && importOption.ValidateMode == ValidateModeEnum.StopOnFirstFailure)
                    continue;

                r.DataCols.ForEach(c =>
                 {
                     KeyValuePair<int, string> kvp = new KeyValuePair<int, string>(c.ColIndex, c.ColValue);
                     var attr = c.GetFilterAttr<DuplicationAttribute>(context.TypeFilterInfo);
                     if (attr != null)
                     {
                         r.SetNotValid(!kvps.Contains(kvp), c, attr.ErrorMsg);
                     }

                     kvps.Add(kvp);
                 });
            }



            return excelDataRows;
        }
    }
}
