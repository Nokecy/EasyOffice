using DinkToPdf;
using DocumentFormat.OpenXml.Packaging;
using EasyOffice.Extensions.Interfaces;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace EasyOffice.Extension.Converter.Converters
{
    public class WordConverter : IWordConverter
    {
        public string ConvertToHTML(byte[] wordBytes,string fileName)
        {
            string htmlText = string.Empty;

            try
            {
                htmlText = ParseDOCX(wordBytes, fileName);
            }
            catch (OpenXmlPackageException e)
            {
                if (e.ToString().Contains("Invalid Hyperlink"))
                {
                    using (MemoryStream ms = new MemoryStream(wordBytes))
                    {
                        UriFixer.FixInvalidUri(ms, brokenUri => FixUri(brokenUri));
                    }
                    htmlText = ParseDOCX(wordBytes, fileName);
                }
            }

            return htmlText;
        }

        public byte[] ConvertToPDF(byte[] wordBytes, string fileName)
        {
            var htmlString = ConvertToHTML(wordBytes, fileName);

            var converter = new BasicConverter(new PdfTools());

            var doc = new HtmlToPdfDocument()
                {
                    GlobalSettings = {
                    ColorMode = DinkToPdf.ColorMode.Color,
                    Orientation = Orientation.Portrait,
                    PaperSize = PaperKind.A4Plus,
                },
                Objects =
                {
                    new ObjectSettings() {
                        PagesCount = true,
                        HtmlContent =        htmlString,
                        WebSettings = { DefaultEncoding = "utf-8" },
                        HeaderSettings = { FontSize = 9, Right = "Page [page] of [toPage]", Line = true, Spacing = 2.812 }
                    }
                }
            };

            var result = converter.Convert(doc);

            return result;
        }

        public static Uri FixUri(string brokenUri)
        {
            string newURI = string.Empty;
            if (brokenUri.Contains("mailto:"))
            {
                int mailToCount = "mailto:".Length;
                brokenUri = brokenUri.Remove(0, mailToCount);
                newURI = brokenUri;
            }
            else
            {
                newURI = " ";
            }
            return new Uri(newURI);
        }

        public static string ParseDOCX(byte[] byteArray, string fileName)
        {
            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    memoryStream.Write(byteArray, 0, byteArray.Length);
                    using (WordprocessingDocument wDoc =
                                                WordprocessingDocument.Open(memoryStream, true))
                    {
                        int imageCounter = 0;
                        var pageTitle = fileName;
                        var part = wDoc.CoreFilePropertiesPart;
                        if (part != null)
                            pageTitle = (string)part.GetXDocument()
                                                    .Descendants(DC.title)
                                                    .FirstOrDefault() ?? fileName;

                        WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                        {
                            AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                            PageTitle = pageTitle,
                            FabricateCssClasses = true,
                            CssClassPrefix = "pt-",
                            RestrictToSupportedLanguages = false,
                            RestrictToSupportedNumberingFormats = false,
                            ImageHandler = imageInfo =>
                            {
                                ++imageCounter;
                                string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                                ImageFormat imageFormat = null;
                                if (extension == "png") imageFormat = ImageFormat.Png;
                                else if (extension == "gif") imageFormat = ImageFormat.Gif;
                                else if (extension == "bmp") imageFormat = ImageFormat.Bmp;
                                else if (extension == "jpeg") imageFormat = ImageFormat.Jpeg;
                                else if (extension == "tiff")
                                {
                                    extension = "gif";
                                    imageFormat = ImageFormat.Gif;
                                }
                                else if (extension == "x-wmf")
                                {
                                    extension = "wmf";
                                    imageFormat = ImageFormat.Wmf;
                                }

                                if (imageFormat == null) return null;

                                string base64 = null;
                                try
                                {
                                    using (MemoryStream ms = new MemoryStream())
                                    {
                                        imageInfo.Bitmap.Save(ms, imageFormat);
                                        var ba = ms.ToArray();
                                        base64 = System.Convert.ToBase64String(ba);
                                    }
                                }
                                catch (System.Runtime.InteropServices.ExternalException)
                                { return null; }

                                ImageFormat format = imageInfo.Bitmap.RawFormat;
                                ImageCodecInfo codec = ImageCodecInfo.GetImageDecoders()
                                                            .First(c => c.FormatID == format.Guid);
                                string mimeType = codec.MimeType;

                                string imageSource =
                                        string.Format("data:{0};base64,{1}", mimeType, base64);

                                XElement img = new XElement(Xhtml.img,
                                        new XAttribute(NoNamespace.src, imageSource),
                                        imageInfo.ImgStyleAttribute,
                                        imageInfo.AltText != null ?
                                            new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                                return img;
                            }
                        };

                        XElement htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                        var html = new XDocument(new XDocumentType("html", null, null, null),
                                                                                    htmlElement);
                        var htmlString = html.ToString(SaveOptions.DisableFormatting);
                        return htmlString;
                    }
                }
            }
            catch (Exception)
            {
                return "The file is either open, please close it or contains corrupt data";
            }
        }

        public string ParseDOCX(FileInfo fileInfo)
        {
            byte[] byteArray = File.ReadAllBytes(fileInfo.FullName);
            return ParseDOCX(byteArray, fileInfo.FullName);
        }
    }
}
