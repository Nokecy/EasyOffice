using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 必填特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
    public class RequiredAttribute : BaseFilterAttribute
    {
        public RequiredAttribute()
        {
            ErrorMsg = "必填";
        }
    }
}
