using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Models.Excel
{
    /// <summary>
    /// Excel数据单元格
    /// </summary>
    public class ExcelDataCol : ExcelCol
    {
        /// <summary>
        /// 对应属性名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 行号
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 字符串值
        /// </summary>
        public string ColValue { get; set; }
    }
}
