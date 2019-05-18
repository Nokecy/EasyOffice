using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 过滤器属性基类
    /// </summary>
    public class BaseFilterAttribute : Attribute
    {
        public string ErrorMsg { get; set; } = "非法";
    }
}
