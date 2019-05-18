using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Filters
{
    /// <summary>
    /// 仅校验特性的类型是不是一样
    /// </summary>
    public class FilterAttributeComparer : IEqualityComparer<BaseFilterAttribute>
    {
        public bool Equals(BaseFilterAttribute x, BaseFilterAttribute y)
        {
            return x.GetType() == y.GetType();
        }

        public int GetHashCode(BaseFilterAttribute obj)
        {
            return obj.GetType().GetHashCode();
        }
    }
}
