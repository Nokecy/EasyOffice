using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Decorators
{
    /// <summary>
    /// 仅校验特性的类型是不是一样
    /// </summary>
    public class DecoratorAttributeComparer : IEqualityComparer<BaseDecorateAttribute>
    {
        public bool Equals(BaseDecorateAttribute x, BaseDecorateAttribute y)
        {
            return x.GetType() == y.GetType();
        }

        public int GetHashCode(BaseDecorateAttribute obj)
        {
            return obj.GetType().GetHashCode();
        }
    }
}
