using EasyOffice.Utils;
using EasyOffice.Models.Word;
using NPOI.OpenXmlFormats.Dml.WordProcessing;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EasyOffice.Providers.NPOI
{
    public static class NPOIHelper
    {
        /// <summary>
        /// 替换Word中的占位符
        /// </summary>
        /// <param name="doc">word</param>
        /// <param name="placeHolderAndValueDict">占位符：值 字典</param>
        /// <param name="hook">钩子委托</param>
        public static void ReplacePlaceholdersInWord(XWPFDocument doc, Dictionary<string, string> stringReplacements
            ,Dictionary<string, IEnumerable<Picture>> pictureReplacements)
        {
            IEnumerator<XWPFParagraph> paragraphEnumerator = doc.GetParagraphsEnumerator();
            XWPFParagraph paragraph;
            while (paragraphEnumerator.MoveNext())
            {
                paragraph = paragraphEnumerator.Current;
                ReplaceInParagraph(stringReplacements, pictureReplacements, paragraph);
            }

            IEnumerator<XWPFTable> tableEnumerator = doc.GetTablesEnumerator();
            XWPFTable table;
            while (tableEnumerator.MoveNext())
            {
                table = tableEnumerator.Current;
                ReplacePlaceholderInTable(table, stringReplacements, pictureReplacements);
            }

            var placeholders = stringReplacements.Select(x => x.Key).Union(pictureReplacements.Select(x => x.Key));

            ClearPlaceholders(placeholders, doc);
        }

        /// <summary>
        /// 清空所有未使用的占位符用空字符串
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="listAllPlaceHolder"></param>
        private static void ClearPlaceholders(IEnumerable<string> placeholders, XWPFDocument doc)
        {
            Dictionary<string, string> replacements = new Dictionary<string, string>();
            foreach (var placeHolder in placeholders)
            {
                replacements.Add(placeHolder, string.Empty);
            }

            IEnumerator<XWPFParagraph> paragraphEnumerator = doc.GetParagraphsEnumerator();
            XWPFParagraph paragraph;
            while (paragraphEnumerator.MoveNext())
            {
                paragraph = paragraphEnumerator.Current;
                ReplaceInParagraph(replacements,null, paragraph);
            }

            IEnumerator<XWPFTable> tableEnumerator = doc.GetTablesEnumerator();
            XWPFTable table;
            while (tableEnumerator.MoveNext())
            {
                table = tableEnumerator.Current;
                ReplacePlaceholderInTable(table, replacements, null);
            }
        }

        /// <summary>
        /// 替换表格中的文本
        /// </summary>
        /// <param name="table"></param>
        /// <param name="placeHolder"></param>
        /// <param name="replaceText"></param>
        /// <param name="setRunStyleDelegate"></param>
        private static void ReplacePlaceholderInTable(XWPFTable table, Dictionary<string, string> stringReplacements,Dictionary<string,IEnumerable<Picture>> pictureReplacements)
        {
            foreach (XWPFTableRow row in table.Rows)
            {
                foreach (XWPFTableCell cell in row.GetTableCells())
                {
                    ReplaceInParagraphs(stringReplacements, pictureReplacements, cell.Paragraphs);
                }
            }
        }

        private static void ReplaceInParagraph(Dictionary<string, string> stringReplacements, 
            Dictionary<string,IEnumerable<Picture>> pictureReplacements,XWPFParagraph paragraph)
        {
            //替换文本
            if (stringReplacements != null)
            {
                foreach (var item in stringReplacements)
                {
                    var matchRun = FindPlaceholderRun(item.Key, paragraph);

                    if (matchRun != null)
                    {
                        matchRun.SetText(item.Value);
                    }
                }
            }

            if (pictureReplacements != null)
            {
                foreach (var item in pictureReplacements)
                {
                    var matchRun = FindPlaceholderRun(item.Key, paragraph);
                    if (matchRun != null)
                    {
                        matchRun.Set(item.Value);
                    }
                }
            }
            
        }

        private static XWPFRun FindPlaceholderRun(string placeholder, XWPFParagraph paragraph)
        {
            XWPFRun matchRun = null;

            var runs = paragraph.Runs;

            TextSegment found = paragraph.SearchText(placeholder, new PositionInParagraph());
            if (found != null)
            {
                if (found.BeginRun == found.EndRun)
                {
                    // 只有一个Run
                    matchRun = runs[found.BeginRun];
                }
                else
                {
                    // 第一个Run设置文本为占位符
                    matchRun = runs[found.BeginRun];

                    // 清空其他Run的文本
                    for (int runPos = found.BeginRun + 1; runPos <= found.EndRun; runPos++)
                    {
                        XWPFRun partNext = runs[runPos];
                        partNext.SetText("", 0);
                    }
                }
            }

            return matchRun;
        }

        private static void ReplaceInParagraphs(Dictionary<string, string> stringReplacements,Dictionary<string,IEnumerable<Picture>> pictureReplacement, 
            IEnumerable<XWPFParagraph> xwpfParagraphs)
        {
            foreach (XWPFParagraph paragraph in xwpfParagraphs)
            {
                ReplaceInParagraph(stringReplacements, pictureReplacement, paragraph);
            }
        }

        /// <summary>
        /// 获取模板文件
        /// </summary>
        /// <param name="contentRootPath"></param>
        /// <returns></returns>
        public static XWPFDocument GetXWPFDocument
            (string fileUrl)
        {
            XWPFDocument word;

            if (!File.Exists(fileUrl))
            {
                throw new Exception("找不到模板文件");
            }

            try
            {
                using (FileStream fs = File.OpenRead(fileUrl))
                {
                    word = new XWPFDocument(fs);
                }
            }
            catch (Exception)
            {
                throw new Exception("打开模板文件失败");
            }

            return word;
        }

        public static void CopyTable(XWPFTable fromTable, XWPFTable toTable)
        {
            var maxCol = fromTable.Rows.Max(x => x.GetTableCells().Count);

            for (int i = 0; i < fromTable.Rows.Count; i++)
            {
                CopyRowToTable(toTable, fromTable.Rows[i], i, maxCol);
            }
        }

        /// <summary>
        /// 将行拷贝到目标表格
        /// 只拷贝了基本的样式
        /// </summary>
        /// <param name="targetTable">目标表格</param>
        /// <param name="sourceRow">源表格行</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="maxCol">表格最大列数</param>
        public static void CopyRowToTable(XWPFTable targetTable, XWPFTableRow sourceRow, int rowIndex,int maxCol)
        {
            //在表格指定位置新增一行
            XWPFTableRow targetRow = rowIndex == 0 ? targetTable.GetRow(0) : targetTable.CreateRow();

            //复制行属性
            targetRow.GetCTRow().trPr = sourceRow.GetCTRow().trPr;
            List<XWPFTableCell> cellList = sourceRow.GetTableCells();
            if (null == cellList)
            {
                return;
            }
            //复制列及其属性和内容
            int colIndex = 0;

            foreach (XWPFTableCell sourceCell in cellList)
            {
                XWPFTableCell targetCell = null;
                //新增行会默认有一个单元格,因此直接获取就好
                try
                {
                    targetCell = targetRow.GetCell(colIndex);
                }
                catch (Exception)
                {
                }

                if (targetCell == null)
                {
                    targetCell = targetRow.CreateCell();
                }

                //列属性
                targetCell.GetCTTc().tcPr = sourceCell.GetCTTc().tcPr;

                //段落属性
                if (sourceCell.Paragraphs != null && sourceCell.Paragraphs.Count > 0)
                {
                    var paragraph = targetCell.Paragraphs[0];
                    var ctp = (CT_P)typeof(XWPFParagraph).GetField("paragraph", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(paragraph);

                    var sourceParagraph = sourceCell.Paragraphs[0];
                    var sourceCtp = (CT_P)typeof(XWPFParagraph).GetField("paragraph", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(sourceParagraph);

                    paragraph.Alignment = sourceParagraph.Alignment;
                    ctp.pPr = sourceCtp.pPr;

                    if (sourceCell.Paragraphs[0].Runs != null && sourceCell.Paragraphs[0].Runs.Count > 0)
                    {
                        XWPFRun cellR = targetCell.Paragraphs[0].CreateRun();
                        var srcRun = sourceCell.Paragraphs[0].Runs[0];
                        cellR.SetText(sourceCell.GetText());
                        cellR.IsBold = srcRun.IsBold;
                        cellR.SetColor(srcRun.GetColor());
                        cellR.FontFamily = srcRun.FontFamily;
                        cellR.IsCapitalized = srcRun.IsCapitalized;
                        cellR.IsDoubleStrikeThrough = srcRun.IsDoubleStrikeThrough;
                        cellR.IsEmbossed = srcRun.IsEmbossed;
                        cellR.IsImprinted = srcRun.IsImprinted;
                        cellR.IsItalic = srcRun.IsItalic;
                        cellR.IsShadowed = srcRun.IsShadowed;
                    }
                    else
                    {
                        targetCell.SetText(sourceCell.GetText());
                    }
                }
                else
                {
                    targetCell.SetText(sourceCell.GetText());
                }

                colIndex++;
            }

            if (cellList.Count < maxCol)
            {
                try
                {
                    targetRow.MergeCells(cellList.Count - 1, maxCol - 1);
                }
                catch
                {
                }
            }
        }

     
    }
}
