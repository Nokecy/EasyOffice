using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 过滤器特性绑定标记
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class FilterBindAttribute : Attribute
    {
        public FilterBindAttribute(Type filterAttributeType)
        {
            if (!filterAttributeType.IsSubclassOf(typeof(BaseFilterAttribute)))
            {
                throw new ArgumentOutOfRangeException(filterAttributeType.ToString() + "is not subclass of BaseFilterAttribute");
            }

            FilterAttributeType = filterAttributeType;
        }

        /// <summary>
        /// 过滤器对应特性类型
        /// </summary>
        public Type FilterAttributeType { get; set; }
    }
}
