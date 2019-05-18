using EasyOffice.Models.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Models.Excel
{
    /// <summary>
    /// Excel表头行
    /// </summary>
    public class ExcelHeaderRow
    {
        public List<ExcelCol> Cells { get; set; } = new List<ExcelCol>();
    }
}
