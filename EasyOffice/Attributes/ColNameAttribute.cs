using System;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 列名
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColNameAttribute : Attribute
    {
        public ColNameAttribute(string colName)
        {
            ColName = colName;
        }

        public string ColName { get; set; }
    }
}
