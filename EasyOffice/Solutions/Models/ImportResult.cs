using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Solutions.Models
{
    /// <summary>
    /// 导入结果
    /// </summary>
    public class ImportResult
    {
        /// <summary>
        /// 错误信息数据库标记
        /// </summary>
        public string Tag { get; set; }

        /// <summary>
        /// 成功条数
        /// </summary>
        public int Success { get; set; }

        /// <summary>
        /// 失败条数
        /// </summary>
        public int Failed { get; set; }

        /// <summary>
        /// 总条数
        /// </summary>
        public int Total { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 最大错误条数
        /// </summary>
        public int MaxErrorMsgCount { get; set; }

        /// <summary>
        /// 错误消息
        /// </summary>
        public List<string> ErrorMsgs { get; set; }

        /// <summary>
        /// 操作结果
        /// </summary>
        public bool OperateSuccess { get; set; }
    }
}
