using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Solutions.Models
{
    public class ImportErrorLogDTO
    {
        [ColName("错误消息")]
        public string ErrorMsg { get; set; }
    }
}
