using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 数值区间特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class RangeAttribute : BaseFilterAttribute
    {
        public int Min { get; set; }
        public int Max { get; set; }

        public RangeAttribute(int min, int max)
        {
            Min = min;
            Max = max;
            ErrorMsg = $"超限，仅允许为{min}-{max}";
        }
    }
}
