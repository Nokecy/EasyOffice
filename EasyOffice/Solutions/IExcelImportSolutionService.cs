using EasyOffice.Models.Excel;
using EasyOffice.Solutions.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace EasyOffice.Solutions
{
    /// <summary>
    /// Excel导入解决方案服务
    /// </summary>
    public interface IExcelImportSolutionService
    {
        /// <summary>
        /// 获取导入配置
        /// </summary>
        /// <typeparam name="T">导入模板类</typeparam>
        /// <param name="uploadUrl">导入api地址</param>
        /// <param name="templateUrl">模板下载api地址</param>
        /// <returns>导入配置</returns>
        Task<ImportConfig> GetImportConfigAsync<T>(string uploadUrl = "", string templateUrl = "") where T : class, new();

        /// <summary>
        /// 获取Excel文件表头和预览数据
        /// </summary>
        /// <param name="fileUrl">Excel文件绝对地址</param>
        /// <param name="rowCount">预览数据行数，默认3条</param>
        /// <param name="headerRowIndex">表头行索引，默认第1行</param>
        /// <param name="dataRowIndex">数据行索引，默认第2行</param>
        /// <param name="headerData">自定义表头数据，用于复杂表头情况</param>
        /// <returns></returns>
        Task<List<List<string>>> GetFileHeadersAndRowsAsync<T>(string fileUrl, int sheetIndex = 0, int dataCount = 3, int headerRowIndex = 0, int dataRowIndex = 1, List<string> headerData = null)
            where T : class, new();

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T">导入模板类类型</typeparam>
        /// <param name="fileUrl">Excel文件Url</param>
        /// <param name="importSetData">匹配后的数据</param>
        /// <param name="businessAction">业务操作</param>
        /// <param name="customValidateAction">自定义校验</param>
        /// <returns></returns>
        Task<ImportResult> ImportAsync<T>(ImportOption importOption, object importSetData, Action<List<T>> businessAction, Func<List<ExcelDataRow>, List<ExcelDataRow>> customValidateAction = null)
            where T : class, new();

        /// <summary>
        /// 导出错误信息
        /// </summary>
        /// <param name="tag"></param>
        /// <returns></returns>
        Task<byte[]> ExportErrorMsgAsync(string tag);

        /// <summary>
        /// 获取默认的导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<byte[]> GetImportTemplateAsync<T>() where T : class, new();
    }
}
