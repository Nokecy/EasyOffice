using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 最大长度特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class MaxLengthAttribute : BaseFilterAttribute
    {
        public MaxLengthAttribute(int maxLength)
        {
            MaxLength = maxLength;
            ErrorMsg = "超长";
        }

        public int MaxLength { get; set; }
    }
}
