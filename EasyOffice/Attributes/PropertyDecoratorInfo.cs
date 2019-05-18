using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 属性装饰信息
    /// </summary>
    public class PropertyDecoratorInfo
    {
        public int ColIndex { get; set; }
        public List<BaseDecorateAttribute> DecoratorAttrs { get; set; }
    }
}
