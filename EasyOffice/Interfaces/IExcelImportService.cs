using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Threading.Tasks;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// Excel导入服务
    /// </summary>
    public interface IExcelImportService
    {
        /// <summary>
        /// 校验Excel数据
        /// </summary>
        /// <typeparam name="T">模板类型</typeparam>
        /// <param name="fileUrl">Excel文件绝对地址</param>
        /// <param name="headerDictionary">表头映射字段 模板类属性名称 ：Excel名称，不传的则按默认配置</param>
        /// <returns>校验后的结果集</returns>
        Task<List<ExcelDataRow>> ValidateAsync<T>(ImportOption importOption) where T : class, new();

        /// <summary>
        /// 将Excel数据读取到DataTable
        /// </summary>
        /// <param name="fileUrl">文件绝对地址</param>
        /// <param name="sheetIndex">sheet索引，默认0</param>
        /// <param name="headerRowIndex">表头行索引，默认0</param>
        /// <param name="dataRowIndex">数据行开始索引，默认1</param>
        /// <param name="dataCount">数据条数，默认-1转换全部</param>
        /// <returns>DataTable</returns>
        Task<DataTable> ToTableAsync<T>(string fileUrl, int sheetIndex = 0, int headerRowIndex = 0, int dataRowIndex = 1, int dataCount = -1)
            where T : class, new();
    }
}
