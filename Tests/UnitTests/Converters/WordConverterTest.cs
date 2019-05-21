using EasyOffice.Extensions.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using UnitTests.DependencyInjection;
using Xunit;

namespace UnitTests.Converters
{
    public class WordConverterTest
    {
        private readonly IWordConverter _wordConverter;
        public WordConverterTest()
        {
            _wordConverter = _wordConverter.Resolve(); ;
        }

        [Fact]
        public void ConvertToPDF_纯文本转换PDF_手动观察导出效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, "Resources/Convert/text.docx");
            var bytes = File.ReadAllBytes(fileUrl);
            var result = _wordConverter.ConvertToPDF(bytes, "text");

            var saveUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf"); ;

            File.WriteAllBytes(saveUrl, result);
        }

        [Fact]
        public void ConvertToPDF_文字图片转换PDF_手动观察导出效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, "Resources/Convert/textAndPicture.docx");
            var bytes = File.ReadAllBytes(fileUrl);
            var result = _wordConverter.ConvertToPDF(bytes, "text");

            var saveUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf"); ;

            File.WriteAllBytes(saveUrl, result);
        }

        [Fact]
        public void ConvertToPDF_表格文字图片转换PDF_手动观察导出效果()
        {
            string curDir = Environment.CurrentDirectory;
        
            string fileUrl = Path.Combine(curDir, @"C:\学习资料\通知中心概述.docx");
            var bytes = File.ReadAllBytes(fileUrl);
            var result = _wordConverter.ConvertToPDF(bytes, "text");

            var saveUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf"); ;

            File.WriteAllBytes(saveUrl, result);
        }
    }
}
