using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 正则表达式特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
    public class RegexAttribute : BaseFilterAttribute
    {
        public string RegexString { get; set; }
        public RegexAttribute(string regex)
        {
            RegexString = regex;
            ErrorMsg = "非法";
        }
    }
}
