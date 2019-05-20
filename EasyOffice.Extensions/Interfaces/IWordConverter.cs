using System;

namespace EasyOffice.Extensions.Interfaces
{
    public interface IWordConverter
    {
        /// <summary>
        /// 将Word文档转换为HTML字符串，仅支持docx格式文档的转换
        /// </summary>
        /// <param name="wordBytes">文档byte数组</param>
        /// <returns>HTML字符串</returns>
        string ConvertToHTML(byte[] wordBytes, string fileName);

        /// <summary>
        /// 将Word文档转换为PDF，仅支持docx格式文档的转换
        /// </summary>
        /// <param name="wordBytes">文档byte数组</param>
        /// <returns>PDF文档的byte数组</returns>
        byte[] ConvertToPDF(byte[] wordBytes, string fileName);
    }
}
