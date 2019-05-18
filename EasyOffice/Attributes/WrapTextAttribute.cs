using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 自动换行特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class WrapTextAttribute : BaseDecorateAttribute
    {
    }
}
