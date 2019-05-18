using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Models.Excel
{
    /// <summary>
    /// 数据行
    /// </summary>
    public class ExcelDataRow
    {
        /// <summary>
        /// 行号
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 单元格数据
        /// </summary>
        public List<ExcelDataCol> DataCols { get; set; } = new List<ExcelDataCol>();

        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMsg { get; set; } = string.Empty;
    }
}
