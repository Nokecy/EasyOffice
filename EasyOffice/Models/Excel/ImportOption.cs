using EasyOffice.Enums;
using EasyOffice.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Models.Excel
{
    public class ImportOption
    {
        public string FileUrl { get; set; }
        public int SheetIndex { get; set; } = 0;
        public int HeaderRowIndex { get; set; } = 0;
        public int DataRowStartIndex { get; set; } = 1;
        public Dictionary<string, string> MappingDictionary { get; set; }
        /// <summary>
        /// 校验模式
        /// </summary>
        public ValidateModeEnum ValidateMode { get; set; } = ValidateModeEnum.StopOnFirstFailure;

        /// <summary>
        /// 自定义导入Provider
        /// </summary>
        public IExcelImportProvider CustomExcelImportProvider { get; set; }
    }
}
