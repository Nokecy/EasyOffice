using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 模板数据重复数据校验
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class DuplicationAttribute : BaseFilterAttribute
    {
        public DuplicationAttribute()
        {
            ErrorMsg = "重复";
        }
    }
}
