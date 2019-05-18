using EasyOffice.Enums;
using EasyOffice.Interfaces;
using EasyOffice.Models.Word;
using EasyOffice.Providers.NPOI;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace EasyOffice.Providers.NPOI
{
    public class WordImportProvider : IWordImportProvider
    {
        public IEnumerable<Paragraph> GetParagraphs(string fileUrl)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Table> GetTables(string fileUrl)
        {
            var result = new List<Table>();

            var doc = NPOIHelper.GetXWPFDocument(fileUrl);

            var xwpfParagraphsEnumerator = doc.GetParagraphsEnumerator();

            while (xwpfParagraphsEnumerator.MoveNext())
            {
                var xwpfParagraph = xwpfParagraphsEnumerator.Current;
            }

            return result;
        }

        private Paragraph Convert(XWPFParagraph xwpfParagraph)
        {
            Paragraph result = null;

            if (xwpfParagraph != null)
            {
                result = new Paragraph();

                result.Alignment = Alignment.LEFT;
                if (xwpfParagraph.Alignment == ParagraphAlignment.CENTER)
                {
                    result.Alignment = Alignment.CENTER;
                }
                if (xwpfParagraph.Alignment == ParagraphAlignment.RIGHT)
                {
                    result.Alignment = Alignment.RIGHT;
                }
            }

            return result;
        }
    }
}
