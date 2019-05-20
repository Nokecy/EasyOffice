using EasyOffice.Enums;
using EasyOffice.Interfaces;
using EasyOffice.Models.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UnitTests.DependencyInjection;
using UnitTests.Models;
using Xunit;

namespace UnitTests.Services
{
    public class WordExportServiceTest : IDisposable
    {
        private readonly IWordExportService _wordExportService;
        public WordExportServiceTest()
        {
            _wordExportService = _wordExportService.Resolve();
        }

        [Fact]
        public async Task CreateFromTemplateTest_从模板导出Word_导出成功人工校验导出效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");

            string pic1 = Path.Combine(curDir, "Resources", "1.jpg");
            string pic2 = Path.Combine(curDir, "Resources", "2.jpg");
            string pic3 = Path.Combine(curDir, "Resources", "3.jpg");
            string templateurl = Path.Combine(curDir, "Resources", "CarWordTemplate.docx");

            WordCarTemplateDTO car = new WordCarTemplateDTO()
            {
                OwnerName = "龚英韬",
                CarType = "豪华型宾利",
                CarPictures = new List<Picture>() {
                     new Picture(){
                          PictureUrl = pic1
                     },
                     new Picture(){
                          PictureUrl = pic2
                     }
                },
                CarLicense = new Picture { PictureUrl = pic3 }
            };

            var word = await _wordExportService.CreateFromTemplateAsync(templateurl, car);

            File.WriteAllBytes(fileUrl, word.WordBytes);
        }

        [Fact]
        public async Task CreateWordTest_导出空白Word_手动观察生成效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");

            var word = await _wordExportService.CreateWordAsync(null);

            File.WriteAllBytes(fileUrl, word.WordBytes);
        }

        [Fact]
        public async Task CreateWordTest_导出段落_手动观察生成效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");

            var p1 = new Paragraph()
            {
                Run = new Run()
                {
                    Text = "寥落古行宫，宫花寂寞红。白头宫女在，闲坐说玄宗。",
                    Color = ColorEnum.RED.ToString(),
                    FontFamily = "黑体",
                    FontSize = 48,
                    IsBold = true
                }
            };

            List<IWordElement> eles = new List<IWordElement>();
            eles.Add(p1);

            var p2 = new Paragraph()
            {
                Alignment = Alignment.CENTER,
                Run = new Run()
                {
                    Text = "红豆生南国，春来发几枝。愿君多采撷，此物最相思。",
                    Color = ColorEnum.BLUE.ToString(),
                    FontFamily = "宋体",
                    FontSize = 12,
                    IsBold = false
                }
            };
            eles.Add(p2);

            var word = await _wordExportService.CreateWordAsync(eles);

            File.WriteAllBytes(fileUrl, word.WordBytes);
        }

        [Fact]
        public async Task CreateWordTest_多段落换行效果_手动观察生成效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");

            var p1 = new Paragraph()
            {
                Run = new Run()
                {
                    Text = "寥落古行宫",
                    Color = ColorEnum.RED.ToString(),
                    FontFamily = "黑体",
                    FontSize = 24,
                    IsBold = true
                }
            };

            var p2 = new Paragraph()
            {
                Run = new Run()
                {
                    Text = "宫花寂寞红",
                    Color = ColorEnum.RED.ToString(),
                    FontFamily = "黑体",
                    FontSize = 24,
                    IsBold = true
                }
            };

            var p3 = new Paragraph()
            {
                Run = new Run()
                {
                    Text = "白头宫女在",
                    Color = ColorEnum.RED.ToString(),
                    FontFamily = "黑体",
                    FontSize = 24,
                    IsBold = true
                }
            };

            var p4 = new Paragraph()
            {
                Run = new Run()
                {
                    Text = "闲坐说玄宗",
                    Color = ColorEnum.RED.ToString(),
                    FontFamily = "黑体",
                    FontSize = 24,
                    IsBold = true
                }
            };

            var word = await _wordExportService.CreateWordAsync(new List<Paragraph> { p1, p2, p3, p4 });

            File.WriteAllBytes(fileUrl, word.WordBytes);
        }

        [Fact]
        public async Task CreateWordTest_导出段落加表格_手动观察生成效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");

            var p1 = new Paragraph()
            {
                Run = new Run()
                {
                    Text = "寥落古行宫，宫花寂寞红。白头宫女在，闲坐说玄宗。",
                    Color = ColorEnum.RED.ToString(),
                    FontFamily = "黑体",
                    FontSize = 48,
                    IsBold = true
                }
            };

            List<IWordElement> eles = new List<IWordElement>();
            eles.Add(p1);

            var p2 = new Paragraph()
            {
                Alignment = Alignment.CENTER,
                Run = new Run()
                {
                    Text = "红豆生南国，春来发几枝。愿君多采撷，此物最相思。",
                    Color = ColorEnum.BLUE.ToString(),
                    FontFamily = "宋体",
                    FontSize = 12,
                    IsBold = false
                }
            };
            eles.Add(p2);

            var t1 = new Table()
            {
                Rows = new List<TableRow>()
            };

            var r1 = new TableRow()
            {
                Cells = new List<TableCell>()
            };
            var r1c1 = new TableCell()
            {
                Paragraphs = new List<Paragraph>{new Paragraph()
                {
                    Run = new Run()
                    {
                        Text = "时间",
                        IsBold = true
                    }
                }
            }
            };
            var r1c2 = new TableCell()
            {
                Paragraphs = new List<Paragraph>(){ new Paragraph()
                {
                    Run = new Run()
                    {
                        Text = "讲者",
                        IsBold = true
                    }
                }
                }
            };
            var r1c3 = new TableCell()
            {
                Paragraphs = new List<Paragraph>(){ new Paragraph()
                {
                    Run = new Run()
                    {
                        Text = "嘉宾",
                        IsBold = true
                    }
                }
                }
            };
            r1.Cells.Add(r1c1);
            r1.Cells.Add(r1c2);
            r1.Cells.Add(r1c3);

            var r2 = new TableRow()
            {
                Cells = new List<TableCell>()
            };
            var r2c1 = new TableCell()
            {
                Paragraphs = new List<Paragraph>(){new  Paragraph()
                {
                    Run = new Run()
                    {
                        Text = "09:00-09:30"
                    }
                }
                }
            };
            var r2c2 = new TableCell()
            {
                Paragraphs = new List<Paragraph>(){new  Paragraph()
                {
                    Run = new Run()
                    {
                        Text = "嘉宾001"
                    }
                }
                }
            };
            var r2c3 = new TableCell()
            {
                Paragraphs = new List<Paragraph>(){ new Paragraph()
                {
                    Run = new Run()
                    {
                        Text = "分公司的股份"
                    }
                }
                }
            };
            r2.Cells.Add(r2c1);
            r2.Cells.Add(r2c2);
            r2.Cells.Add(r2c3);

            t1.Rows.Add(r1);
            t1.Rows.Add(r2);

            eles.Add(t1);

            var word = await _wordExportService.CreateWordAsync(eles);

            File.WriteAllBytes(fileUrl, word.WordBytes);
        }

        [Fact]
        public async Task CreateWordFromMasterTable_根据母版Table生成Word包括图片占位符_手动观察生成效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");

            string pic1 = Path.Combine(curDir, "Resources", "1.jpg");
            string pic2 = Path.Combine(curDir, "Resources", "2.jpg");
            string pic3 = Path.Combine(curDir, "Resources", "3.jpg");
            string templateurl = Path.Combine(curDir, "Resources", "CarWordTemplate.docx");

            var car1 = new WordCarTemplateDTO()
            {
                OwnerName = "龚英韬",
                CarType = "豪华型宾利",
                CarPictures = new List<Picture>() {
                     new Picture(){
                          PictureUrl = pic1,
                           
                     },
                     new Picture(){
                          PictureUrl = pic2
                     }
                },
                CarLicense = new Picture { PictureUrl = pic3 }
            };

            var car2 = new WordCarTemplateDTO()
            {
                OwnerName = "龚英韬",
                CarType = "豪华型宾利",
                CarPictures = new List<Picture>() {
                     new Picture(){
                          PictureUrl = pic1
                     },
                     new Picture(){
                          PictureUrl = pic2
                     }
                },
                CarLicense = new Picture { PictureUrl = pic3 }
            };

            var datas = new List<WordCarTemplateDTO>() {
                 car1,car2
            };

            var word = await _wordExportService.CreateFromMasterTableAsync(templateurl, datas);

            File.WriteAllBytes(fileUrl, word.WordBytes);
        }

        [Fact]
        public async Task CreateWordFromMasterTable_根据母版Table生成Word_手动观察生成效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");
            string templateurl = Path.Combine(curDir, "Resources", "UserMasterTable.docx");
            var user1 = new UserInfoDTO()
            {
                Name = "张三",
                Age = 15,
                Gender = "男",
                Remarks = "简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介"
            };
            var user2 = new UserInfoDTO()
            {
                Name = "李四",
                Age = 20,
                Gender = "女",
                Remarks = "简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介简介"
            };

            var datas = new List<UserInfoDTO>() { user1, user2 };

            for (int i = 0; i < 10; i++)
            {
                datas.Add(user1);
                datas.Add(user2);
            }

            var word = await _wordExportService.CreateFromMasterTableAsync(templateurl, datas);

            File.WriteAllBytes(fileUrl, word.WordBytes);
        }

        [Fact]
        public async Task CreateWordTest_生成包括段落和图片的Word文档_手动观察生成效果()
        {
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");
            string pic1 = Path.Combine(curDir, "Resources", "gn.jpg");
            string pic2 = Path.Combine(curDir, "Resources", "1.jpg");

            var p1 = new Paragraph()
            {
                Run = new Run()
                {
                    Text = "寥落古行宫，宫花寂寞红。白头宫女在，闲坐说玄宗。",
                    Color = ColorEnum.RED.ToString(),
                    FontFamily = "黑体",
                    FontSize = 12,
                    IsBold = true
                }
            };

            List<IWordElement> eles = new List<IWordElement>();
            eles.Add(p1);

            var p2 = new Paragraph()
            {
                Alignment = Alignment.CENTER,
                Run = new Run()
                {
                    Text = "红豆生南国，春来发几枝。愿君多采撷，此物最相思。",
                    Color = ColorEnum.BLUE.ToString(),
                    FontFamily = "宋体",
                    FontSize = 12,
                    IsBold = false
                }
            };
            eles.Add(p2);

            var p3 = new Paragraph()
            {
                Run = new Run()
                {
                    Pictures = new List<Picture>()
                    {
                        new Picture()
                        {
                            FileName = "宫女",
                            PictureUrl = pic1,
                            PictureType = PictureTypeEnum.JPEG,
                            Height = 8,
                            Width = 8
                            //PictureId = 1
                        },
                        new Picture()
                        {
                            FileName = "宫女",
                            PictureUrl = pic2,
                            PictureType = PictureTypeEnum.JPEG,
                            Height = 8,
                            Width = 8
                        }
                    }
                }
            };

            eles.Add(p3);

            var word = await _wordExportService.CreateWordAsync(eles);

            File.WriteAllBytes(fileUrl, word.WordBytes);
        }

        public void Dispose()
        {
        }
    }
}
