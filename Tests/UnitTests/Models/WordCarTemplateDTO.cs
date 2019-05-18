using EasyOffice.Attributes;
using EasyOffice.Models.Word;
using System.Collections.Generic;

namespace UnitTests.Models
{
    public class WordCarTemplateDTO
    {
        public string OwnerName { get; set; }

        [Placeholder("{Car_Type Car Type}")]
        public string CarType { get; set; }

        //图片占位的属性类型必须为List<string>,存放图片的绝对全地址
        public IEnumerable<Picture> CarPictures { get; set; }

        public Picture CarLicense { get; set; }
    }
}
