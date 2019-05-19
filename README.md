# 简介
EasyOffice提供了Excel导入导出，Word根据模板生成，根据Table母版生成，从空白创建等功能。
绝大部分Excel、Word导入导出工作只需要通过配置就可以轻松解决。

# 依赖注入
使用Autofac注入

```c
// 注入Office基础服务
builder.AddOffice(new OfficeOptions());
```

---

# IExcelImportService - Excel通用导入
## 定义Excel模板类

```c
 public class Car
    {
        [ColName("车牌号")]  //对应Excel列名
        [Required] //校验必填
        [Regex(RegexConstant.CAR_CODE_REGEX)] //正则表达式校验,RegexConstant预置了一些常用的正则表达式，也可以自定义
        [Duplication] //校验模板类该列数据是否重复
        public string CarCode { get; set; }

        [ColName("手机号")]
        [Regex(RegexConstant.MOBILE_CHINA_REGEX)]
        public string Mobile { get; set; }

        [ColName("身份证号")]
        [Regex(RegexConstant.IDENTITY_NUMBER_REGEX)]
        public string IdentityNumber { get; set; }

        [ColName("姓名")]
        [MaxLength(10)] //最大长度校验
        public string Name { get; set; }

        [ColName("性别")] 
        [Regex(RegexConstant.GENDER_REGEX)]
        public GenderEnum Gender { get; set; }

        [ColName("注册日期")]
        [DateTime] //日期校验
        public DateTime RegisterDate { get; set; }

        [ColName("年龄")]
        [Range(0, 150)] //数值范围校验
        public int Age { get; set; }
    }
```

## 校验数据

```c

    var _rows = _excelImportService.ValidateAsync<ExcelCarTemplateDTO>(new ImportOption()
    {
        FileUrl = fileUrl, //Excel文件绝对地址
        DataRowStartIndex = 1, //数据起始行索引，默认1第二行
        HeaderRowIndex = 0,  //表头起始行索引，默认0第一行
        MappingDictionary = null, //映射字典，可以将模板类与Excel列重新映射， 默认null
        SheetIndex = 0, //页面索引，默认0第一个页签
        ValidateMode = ValidateModeEnum.Continue //校验模式，默认StopOnFirstFailure校验错误后此行停止继续校验，Continue：校验错误后继续校验
    }).Result;
    
    //得到错误行
    var errorDatas = _rows.Where(x => !x.IsValid);
    //错误行业务处理
    
    //将有效数据行转换为指定类型
    var validDatas = _rows.Where(x=>x.IsValid).FastConvert<ExcelCarTemplateDTO>();
    //正确数据业务处理
```

## 转换为DataTable

```
      var dt = _excelImportService.ToTableAsync<ExcelCarTemplateDTO> //模板类型
                (
                fileUrl,  //文件绝对地址
                0,  //页签索引，默认0
                0,  //表头行索引，默认0
                1, //数据行索引，默认1
                -1); //读取多少条数据，默认-1全部
```


---

# IExcelExportService - 通用Excel导出服务
## 定义导出模板类

```
    [Header(Color = ColorEnum.BRIGHT_GREEN, FontSize = 22, IsBold = true)] //表头样式
    [WrapText] //自动换行
    public class ExcelCarTemplateDTO
    {
        [ColName("车牌号")]
        [MergeCols] //相同数据自动合并单元格
        public string CarCode { get; set; }

        [ColName("手机号")]
        public string Mobile { get; set; }

        [ColName("身份证号")]
        public string IdentityNumber { get; set; }

        [ColName("姓名")]
        public string Name { get; set; }

        [ColName("性别")]
        public GenderEnum Gender { get; set; }

        [ColName("注册日期")]
        public DateTime RegisterDate { get; set; }

        [ColName("年龄")]
        public int Age { get; set; }
```

## 导出Excel

```
    var bytes = await _excelExportService.ExportAsync(new ExportOption<ExcelCarTemplateDTO>()
    {
        Data = list,
        DataRowStartIndex = 1, //数据行起始索引，默认1
        ExcelType = Bayantu.Extensions.Office.Enums.ExcelTypeEnum.XLS,//导出Excel类型，默认xls
        HeaderRowIndex = 0, //表头行索引，默认0
        SheetName = "sheet1" //页签名称，默认sheet1
    });

    File.WriteAllBytes(@"c:\test.xls", bytes);
```


---

# IExcelImportSolutionService - Excel导入解决方案服务(与前端控件配套的完整解决方案)

首先定义模板类，参考通用Excel导入

```
   //获取默认导入模板
    var templateBytes = await _excelImportSolutionService.GetImportTemplateAsync<DemoTemplateDTO>();

    //获取导入配置
    var importConfig = await _excelImportSolutionService.GetImportConfigAsync<DemoTemplateDTO>("uploadUrl","templateUrl");

    //获取预览数据
    var previewData = await _excelImportSolutionService.GetFileHeadersAndRowsAsync<DemoTemplateDTO>("fileUrl");

    //导入
    var importOption = new ImportOption()
    {
        FileUrl = "fileUrl",
        ValidateMode = ValidateModeEnum.Continue
    };
    object importSetData = new object(); //前端传过来的映射数据
    var importResult = await _excelImportSolutionService.ImportAsync<DemoTemplateDTO>
        (importOption
        , importSetData
        , BusinessAction //业务方法委托
        , CustomValidate //自定义校验委托
        );

    //获取导入错误消息
    var errorMsg = await _excelImportSolutionService.ExportErrorMsgAsync(importResult.Tag);
```


---


# IWordExportService - Word通用导出服务
## CreateFromTemplateAsync - 根据模板生成Word
```
//step1 - 定义模板类
 public class WordCarTemplateDTO
    {
        //默认占位符为{PropertyName}
        public string OwnerName { get; set; }

        [Placeholder("{Car_Type Car Type}")] //重写占位符
        public string CarType { get; set; }

        //使用Picture或IEnumerable<Picture>类型可以将占位符替换为图片
        public IEnumerable<Picture> CarPictures { get; set; }

        public Picture CarLicense { get; set; }
    }

//step2 - 制作word模板

//step3 - 导出word
string templateUrl = @"c:\template.docx";
WordCarTemplateDTO car = new WordCarTemplateDTO()
{
    OwnerName = "刘德华",
    CarType = "豪华型宾利",
    CarPictures = new List<Picture>() {
         new Picture()
         {
              PictureUrl = pic1, //图片绝对地址，如果设置了PictureData此项不生效
              FileName = "图片1",//文件名称
              Height = 10,//图片高度单位厘米默认8
              Width = 3,//图片宽度单位厘米默认14
              PictureData = null,//图片流数据，优先取这里的数据，没有则取url
              PictureType = PictureTypeEnum.JPEG //图片类型，默认jpeg
         },
         new Picture(){
              PictureUrl = pic2
         }
    },
    CarLicense = new Picture { PictureUrl = pic3 }
};

var word = await _wordExportService.CreateFromTemplateAsync(templateUrl, car);

File.WriteAllBytes(@"c:\file.docx", word.WordBytes);
```

## CreateWordFromMasterTable-根据模板表格循环生成word

```
//step1 - 定义模板类，参考上面

//step2 - 定义word模板,制作一个表格,填好占位符。

//step3 - 调用,如下示例,最终生成的word有两个用户表格
  string templateurl = @"c:\template.docx";
    var user1 = new UserInfoDTO()
    {
        Name = "张三",
        Age = 15,
        Gender = "男",
        Remarks = "简介简介"
    };
    var user2 = new UserInfoDTO()
    {
        Name = "李四",
        Age = 20,
        Gender = "女",
        Remarks = "简介简介简介"
    };
    
    var datas = new List<UserInfoDTO>() { user1, user2 };
    
    for (int i = 0; i < 10; i++)
    {
        datas.Add(user1);
        datas.Add(user2);
    }
    
    var word = await _wordExportService.CreateFromMasterTableAsync(templateurl, datas);
    
    File.WriteAllBytes(@"c:\file.docx", word.WordBytes);
```

## CreateWordAsync - 从空白生成word

```
  [Fact]
        public async Task 导出所有日程()
        {
            //准备数据
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

            //新建一个表格
            var table = new Table()
            {
                Rows = new List<TableRow>()
            };

            foreach (var date in dates)
            {
                //新建一行
                var rowDate = new TableRow()
                {
                    Cells = new List<TableCell>()
                };
                
                //新增单元格
                rowDate.Cells.Add(new TableCell()
                {
                    Color = "lightblue", //设置单元格颜色
                    Paragraphs = new List<Paragraph>()
                    { 
                    //新增段落
                    new Paragraph()
                    {
                       //段落里面新增文本域
                        Run = new Run()
                        {
                           Text = date.DateTimeStr,//文本域文本，Run还可以
                           Color = "red", //设置文本颜色
                           FontFamily = "微软雅黑",//设置文本字体
                           FontSize = 12,//设置文本字号
                           IsBold = true,//是否粗体
                           Pictures = new List<Picture>()//也可以插入图片
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
```












