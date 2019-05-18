using NPOI.SS.UserModel;
using System.IO;

namespace EasyOffice.Providers.NPOI
{
    public static class NPOIExcelExtensions
    {
        /// <summary>
        /// 将IWorkbook转换为byte数组
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static byte[] ToBytes(this IWorkbook workbook)
        {
            byte[] result;
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                result = ms.ToArray();
            }

            return result;
        }

        /// <summary>
        /// 获取单元格的字符串值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static string GetStringValue(this ICell cell)
        {
            try
            {
                switch (cell.CellType)
                {
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Error:
                        return cell.ErrorCellValue.ToString();
                    case CellType.Numeric:
                        return DateUtil.IsCellDateFormatted(cell)
                       ? cell.DateCellValue.ToString()
                       : cell.NumericCellValue.ToString();
                    case CellType.String:
                        return cell.StringCellValue;
                    default:
                        return cell.ToString();
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        public static IWorkbook ToWorkbook(this byte[] workbookBytes)
        {
            IWorkbook workbook = null;
            using (var stream = new MemoryStream(workbookBytes))
            {
                workbook = WorkbookFactory.Create(stream);
            }
            return workbook;
        }
    }
}
