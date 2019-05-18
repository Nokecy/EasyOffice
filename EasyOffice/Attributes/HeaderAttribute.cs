using EasyOffice.Attributes;
using EasyOffice.Enums;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 表头特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class HeaderAttribute : BaseDecorateAttribute
    {
        public ColorEnum Color { get; set; } = ColorEnum.BLACK;
        public string FontName { get; set; } = "微软雅黑";
        public int FontSize { get; set; } = 12;
        public bool IsBold { get; set; } = true;
    }
}
