using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 日期属性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class DateTimeAttribute : BaseFilterAttribute
    {
        public DateTimeAttribute()
        {
            ErrorMsg = "非法";
        }
    }
}
