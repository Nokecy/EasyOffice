using EasyOffice.Enums;
using EasyOffice.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Models.Excel
{
    public class ExportOption<T>
        where T : class, new()
    {
        /// <summary>
        /// 导出Excel类型
        /// </summary>
        public ExportType ExportType { get; set; } = ExportType.XLSX;

        /// <summary>
        /// 页签名称
        /// </summary>
        public string SheetName { get; set; } = "sheet1";

        /// <summary>
        /// 表头行索引，默认为0
        /// </summary>
        public int HeaderRowIndex { get; set; } = 0;

        /// <summary>
        /// 数据行起始索引，默认为1
        /// </summary>
        public int DataRowStartIndex { get; set; } = 1;

        /// <summary>
        /// 自定义Excel导出Provider
        /// </summary>
        public IExcelExportProvider CustomExcelExportProvider { get; set; }
    }
}
