using EasyOffice.Interfaces;
using EasyOffice.Models.Word;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// Word导出服务
    /// </summary>
    public interface IWordExportService
    {
        /// <summary>
        /// 根据模板创建Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        Task<Word> CreateFromTemplateAsync<T>(string templateUrl, T data, IWordExportProvider customWordExportProvider = null) where T : class, new();

        /// <summary>
        /// 从空白创建Word文档
        /// </summary>
        /// <param name="elements"></param>
        /// <returns></returns>
        Task<Word> CreateWordAsync(IEnumerable<IWordElement> elements, IWordExportProvider customWordExportProvider = null);

        /// <summary>
        /// 从母版表格中复制样式，N条数据生成N个复制品并替换数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="datas"></param>
        /// <returns></returns>
        Task<Word> CreateFromMasterTableAsync<T>(string templateUrl, IEnumerable<T> datas, IWordExportProvider customWordExportProvider = null)
            where T : class, new();
    }
}
