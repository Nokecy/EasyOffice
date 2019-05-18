using EasyOffice.Interfaces;
using EasyOffice.Models.Word;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// Word导出Provider
    /// </summary>
    public interface IWordImportProvider
    {
        IEnumerable<Table> GetTables(string fileUrl);

        IEnumerable<Paragraph> GetParagraphs(string fileUrl);
    }
}
