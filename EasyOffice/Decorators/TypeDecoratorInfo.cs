using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Decorators
{
    /// <summary>
    /// 模板类装饰器信息
    /// </summary>
    public class TypeDecoratorInfo
    {
        public List<BaseDecorateAttribute> TypeDecoratorAttrs { get; set; }
        public List<PropertyDecoratorInfo> PropertyDecoratorInfos { get; set; }
    }
}
