using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 合并单元格特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class MergeColsAttribute : BaseDecorateAttribute
    {
    }
}
