using EasyOffice.Attributes;
using EasyOffice.Enums;
using EasyOffice.Utils;
using EasyOffice.Interfaces;
using EasyOffice.Models.Word;
using EasyOffice.Providers.NPOI;
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
    public class WordExportProvider : IWordExportProvider
    {
    
        public Word ExportFromTemplate<T>(string templateUrl, T wordData) where T : class, new()
        {
            XWPFDocument word = NPOIHelper.GetXWPFDocument(templateUrl);

            ReplacePlaceholders(word, wordData);

            var result = new Word()
            {
                Type = SolutionEnum.NPOI,
                WordBytes = word.ToBytes()
            };

            return result;
        }

        /// <summary>
        /// 图片Id
        /// </summary>
        private static uint PicId
        {
            get
            {
                picId++;
                return picId;
            }
        }

        private static uint picId = 0;

        /// <summary>
        /// 替换占位符
        /// </summary>
        /// <param name="word"></param>
        private void ReplacePlaceholders<T>(XWPFDocument word, T wordData)
            where T : class, new()
        {
            if (word == null)
            {
                throw new ArgumentNullException("word");
            }

            Dictionary<string, string> stringReplacements = WordHelper.GetReplacements(wordData);

            Dictionary<string, IEnumerable<Picture>> pictureReplacements = WordHelper.GetPictureReplacements(wordData);

            NPOIHelper.ReplacePlaceholdersInWord(word, stringReplacements, pictureReplacements);
        }

        public Word CreateWord()
        {
            var doc = new XWPFDocument();

            var result = new Word()
            {
                Type = SolutionEnum.NPOI,
                WordBytes = doc.ToBytes()
            };

            return result;
        }

        public Word InsertParagraphs(Word word, List<Paragraph> paragraphs)
        {
            XWPFDocument doc = null;
            using (MemoryStream ms = new MemoryStream(word.WordBytes))
            {
                doc = new XWPFDocument(ms);
            }
            foreach (var p in paragraphs)
            {
                doc = InsertParagraph(doc, p);
            }

            var result = new Word()
            {
                Type = SolutionEnum.NPOI,
                WordBytes = doc.ToBytes()
            };

            return result;
        }

        public Word InsertTables(Word word, List<Table> tables)
        {
            XWPFDocument doc = word.ToNPOI();

            foreach (var t in tables)
            {
                doc = InsertTable(doc, t);
            }

            var result = new Word()
            {
                Type = SolutionEnum.NPOI,
                WordBytes = doc.ToBytes()
            };

            return result;
        }

        private XWPFDocument InsertTable(XWPFDocument doc, Table t)
        {
            var maxColCount = t.Rows.Max(x => x.Cells.Count);

            if (t == null) return doc;

            var table = doc.CreateTable();

            table.Width = t.Width;

            int index = 0;
            t.Rows?.ForEach(r =>
            {
                XWPFTableRow tableRow = index == 0 ? table.GetRow(0) : table.CreateRow();

                for (int i = 0; i < r.Cells.Count; i++)
                {
                    var cell = r.Cells[i];
                    var xwpfCell = i == 0 ? tableRow.GetCell(0) : tableRow.AddNewTableCell();
                    foreach (var para in cell.Paragraphs)
                    {
                        xwpfCell.AddParagraph().Set(para);
                    }

                    if (!string.IsNullOrWhiteSpace(cell.Color))
                    {
                        tableRow.GetCell(i).SetColor(cell.Color);
                    }
                }

                //补全单元格，并合并
                var rowColsCount = tableRow.GetTableICells().Count;
                if (rowColsCount < maxColCount)
                {
                    for (int i = rowColsCount - 1; i < maxColCount; i++)
                    {
                        tableRow.CreateCell();
                    }
                    tableRow.MergeCells(rowColsCount - 1, maxColCount);
                }

                index++;
            });

            return doc;
        }

        private XWPFDocument InsertParagraph(XWPFDocument doc, Paragraph paragraph)
        {
            doc.CreateParagraph().Set(paragraph);

            return doc;
        }

        public Word CreateFromMasterTable<T>(string templateUrl, IEnumerable<T> datas) where T : class, new()
        {
            var template = NPOIHelper.GetXWPFDocument(templateUrl);

            var result = new Word()
            {
                Type = SolutionEnum.NPOI
            };

            var tables = template.GetTablesEnumerator();

            if (tables == null)
            {
                result.WordBytes = template.ToBytes();
                return result;
            }

            var masterTables = new List<XWPFTable>();

            while (tables.MoveNext())
            {
                masterTables.Add(tables.Current);
            }

            var doc = new XWPFDocument();

            foreach (var data in datas)
            {
                foreach (var masterTable in masterTables)
                {
                    var cloneTable = doc.CreateTable();
                    NPOIHelper.CopyTable(masterTable, cloneTable);
                }

                ReplacePlaceholders(doc, data);
            }

            result.WordBytes = doc.ToBytes();

            return result;
        }
    }
}
