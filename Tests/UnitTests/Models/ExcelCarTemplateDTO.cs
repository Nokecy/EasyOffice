using EasyOffice.Attributes;
using EasyOffice.Constants;
using EasyOffice.Enums;
using System;
using System.Collections.Generic;
using System.Text;

namespace UnitTests.Models
{
    //[Header(Color = ColorEnum.BRIGHT_GREEN, FontSize = 22, IsBold = true)]
    //[WrapText]
    public class ExcelCarTemplateDTO
    {
        [ColName("车牌号")]
        [Required]
        [Regex(RegexConstant.CAR_CODE_REGEX)]
        [Duplication]
        //[MergeCols]
        public string CarCode { get; set; }

        [ColName("手机号")]
        [Regex(RegexConstant.MOBILE_CHINA_REGEX)]
        public string Mobile { get; set; }

        [ColName("身份证号")]
        [Regex(RegexConstant.IDENTITY_NUMBER_REGEX)]
        public string IdentityNumber { get; set; }

        [ColName("姓名")]
        [MaxLength(10)]
        public string Name { get; set; }

        [ColName("性别")]
        [Regex(RegexConstant.GENDER_REGEX)]
        public GenderEnum Gender { get; set; }

        [ColName("注册日期")]
        [DateTime]
        public DateTime RegisterDate { get; set; }

        [ColName("年龄")]
        [Range(0, 150)]
        public int Age { get; set; }
    }

    public enum GenderEnum
    {
        男 = 10,
        女 = 20
    }
}
