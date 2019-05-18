using EasyOffice.Filters;
using EasyOffice.Models.Excel;
using System.Collections.Generic;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// 过滤器接口
    /// </summary>
    public interface IFilter
    {
        /// <summary>
        /// 校验
        /// </summary>
        /// <param name="rowWrappers"></param>
        /// <returns></returns>
        List<ExcelDataRow> Filter(List<ExcelDataRow> excelDataRows, FilterContext context, ImportOption importOption);
    }

}
