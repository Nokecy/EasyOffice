using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using EasyOffice.Interfaces;
using EasyOffice.Models.Word;

namespace EasyOffice.Services
{
    public class WordExportService : IWordExportService
    {
        private readonly IWordExportProvider _wordExportProvider;

        public WordExportService(IWordExportProvider wordExportProvider)
        {
            _wordExportProvider = wordExportProvider;
        }

        public Task<Word> CreateWordAsync(IEnumerable<IWordElement> elements,IWordExportProvider customWordExportProvider=null)
        {
            var provider = customWordExportProvider == null ? _wordExportProvider : customWordExportProvider;

            var word = provider.CreateWord();

            elements?.ToList().ForEach(x =>
            {
                if (x is Paragraph)
                {
                    word = provider.InsertParagraphs(word, new List<Paragraph>() { x as Paragraph });
                }

                if (x is Table)
                {
                    word = provider.InsertTables(word, new List<Table>() { x as Table });
                }
            });

            return Task.FromResult(word);
        }

        public Task<Word> CreateFromTemplateAsync<T>(string templateUrl
            , T wordData
            , IWordExportProvider customWordExportProvider = null) 
        where T : class, new()
        {
            var provider = customWordExportProvider == null ? _wordExportProvider : customWordExportProvider;

            var word = _wordExportProvider.ExportFromTemplate(templateUrl, wordData);
            return Task.FromResult(word);
        }

        public Task<Word> CreateFromMasterTableAsync<T>(string templateUrl
            , IEnumerable<T> datas
            , IWordExportProvider customWordExportProvider = null)
            where T : class, new()
        {
            var provider = customWordExportProvider == null ? _wordExportProvider : customWordExportProvider;
            var word = _wordExportProvider.CreateFromMasterTable(templateUrl, datas);
            return Task.FromResult(word);
        }
    }
}
