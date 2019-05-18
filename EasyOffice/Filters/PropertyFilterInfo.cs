using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Filters
{
    /// <summary>
    /// 属性过滤器信息
    /// </summary>
    public class PropertyFilterInfo
    {
        public string PropertyName { get; set; }
        public List<BaseFilterAttribute> FilterAttrs { get; set; }
    }
}
