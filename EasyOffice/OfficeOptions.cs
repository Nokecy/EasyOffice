using EasyOffice.Enums;

namespace EasyOffice
{
    /// <summary>
    /// 配置参数
    /// </summary>
    public class OfficeOptions
    {
        /// <summary>
        /// Excel导入底层组件，默认NPOI
        /// </summary>
        public SolutionEnum ExcelImportSolution { get; set; } = SolutionEnum.NPOI;

        /// <summary>
        /// Excel导出底层组件，默认NPOI
        /// </summary>
        public SolutionEnum ExcelExportSolution { get; set; } = SolutionEnum.NPOI;

        /// <summary>
        /// Word导入底层组件，默认NPOI
        /// </summary>
        public SolutionEnum WordImportSolution { get; set; } = SolutionEnum.NPOI;

        /// <summary>
        /// Word导出底层组件，默认NPOI
        /// </summary>
        public SolutionEnum WordExportSolution { get; set; } = SolutionEnum.NPOI;
    }
}