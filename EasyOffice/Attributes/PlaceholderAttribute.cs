using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 普通占位符特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class PlaceholderAttribute : Attribute
    {
        public PlaceholderAttribute(string placeHolder)
        {
            Placeholder = placeHolder;
        }

        public string Placeholder { get; set; }
    }
}
