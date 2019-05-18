using EasyOffice.Enums;
using EasyOffice.Interfaces;
using EasyOffice.Models.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using UnitTests.DependencyInjection;
using UnitTests.Models;
using Xunit;

namespace UnitTests.ScheduleDemo
{
    public class ScheduleDemo
    {
        private readonly IWordExportService _wordExportService;
        public ScheduleDemo()
        {
            _wordExportService = _wordExportService.Resolve();
        }

        [Fact]
        public async Task 导出所有日程()
        {
            //准备数据
            string curDir = Environment.CurrentDirectory;
            string fileUrl = Path.Combine(curDir, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");

            var date1 = new ScheduleDate()
            {
                DateTimeStr = "2019年5月5日 星期八",
                Addresses = new List<Address>()
            };

            var address1 = new Address()
            {
                Name = "会场一",
                Categories = new List<Category>()
            };

            var cate1 = new Category()
            {
                Name = "分类1",
                Schedules = new List<Schedule>()
            };

            var schedule1 = new Schedule()
            {
                Name = "日程1",
                TimeString = "上午9：00 - 上午12：00",
                Speakers = new List<Speaker>()
            };
            var schedule2 = new Schedule()
            {
                Name = "日程2",
                TimeString = "下午13：00 - 下午14：00",
                Speakers = new List<Speaker>()
            };

            var speaker1 = new Speaker()
            {
                Name = "张三",
                Position = "总经理"
            };
            var speaker2 = new Speaker()
            {
                Name = "李四",
                Position = "副总经理"
            };

            schedule1.Speakers.Add(speaker1);
            schedule1.Speakers.Add(speaker2);
            cate1.Schedules.Add(schedule1);
            cate1.Schedules.Add(schedule2);
            address1.Categories.Add(cate1);
            date1.Addresses.Add(address1);

            var dates = new List<ScheduleDate>() { date1,date1,date1 };

            var tables = new List<Table>();

            //从空白生成Word
            var table = new Table()
            {
                Rows = new List<TableRow>()
            };

            foreach (var date in dates)
            {
                //会议日期行
                var rowDate = new TableRow()
                {
                    Cells = new List<TableCell>()
                };
                rowDate.Cells.Add(new TableCell()
                {
                    Color = "lightblue",
                    Paragraphs = new List<Paragraph>()
                    { new Paragraph()
                    {
                        Run = new Run()
                        {
                            Text = date.DateTimeStr
                        },
                         Alignment = Alignment.CENTER //段落居中
                    }
                    }
                });
                table.Rows.Add(rowDate);

                //会场
                foreach (var addr in date.Addresses)
                {
                    //分类
                    foreach (var cate in addr.Categories)
                    {
                        var rowCate = new TableRow()
                        {
                            Cells = new List<TableCell>()
                        };

                        //会场名称
                        rowCate.Cells.Add(new TableCell()
                        {
                            Paragraphs = new List<Paragraph>{ new Paragraph()
                            {
                                Run = new Run()
                                {
                                    Text = addr.Name,
                                }
                            }
                            }
                        });

                        rowCate.Cells.Add(new TableCell()
                        {
                            Paragraphs = new List<Paragraph>(){ new Paragraph()
                            {
                                Run = new Run()
                                {
                                    Text = cate.Name,
                                }
                            }
                            }
                        });
                        table.Rows.Add(rowCate);

                        //日程
                        foreach (var sche in cate.Schedules)
                        {
                            var rowSche = new TableRow()
                            {
                                Cells = new List<TableCell>()
                            };

                            var scheCell = new TableCell()
                            {
                                Paragraphs = new List<Paragraph>()
                                {
                                    new Paragraph()
                                    {
                                         Run = new Run()
                                         {
                                              Text = sche.Name
                                         }
                                    },
                                    {
                                    new Paragraph()
                                    {
                                        Run = new Run()
                                        {
                                            Text = sche.TimeString
                                        }
                                    }
                                    }
                                }
                            };

                            foreach (var speaker in sche.Speakers)
                            {
                                scheCell.Paragraphs.Add(new Paragraph()
                                {
                                    Run = new Run()
                                    {
                                        Text = $"{speaker.Position}:{speaker.Name}"
                                    }
                                });
                            }

                            rowSche.Cells.Add(scheCell);

                            table.Rows.Add(rowSche);
                        }
                    }
                }
            }

            tables.Add(table);

            var word = await _wordExportService.CreateWordAsync(tables);

            File.WriteAllBytes(fileUrl, word.WordBytes);
        }
    }
}
