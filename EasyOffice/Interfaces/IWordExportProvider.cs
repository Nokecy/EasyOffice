using EasyOffice.Models.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// Word导出Provider
    /// </summary>
    public interface IWordExportProvider
    {
        /// <summary>
        /// 根据模板导出Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        Word ExportFromTemplate<T>(string templateUrl, T data) where T : class, new();

        /// <summary>
        /// 创建空白Word
        /// </summary>
        /// <returns></returns>
        Word CreateWord();

        /// <summary>
        /// 插入段落
        /// </summary>
        /// <param name="word"></param>
        /// <returns></returns>
        Word InsertParagraphs(Word word, List<Paragraph> paragraphs);

        /// <summary>
        /// 插入表格
        /// </summary>
        /// <param name="word"></param>
        /// <param name="tables"></param>
        /// <returns></returns>
        Word InsertTables(Word word, List<Table> tables);

        /// <summary>
        /// 根据母版表格创建Word
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="datas"></param>
        /// <returns></returns>
        Word CreateFromMasterTable<T>(string templateUrl, IEnumerable<T> datas) where T : class, new();
    }
}
