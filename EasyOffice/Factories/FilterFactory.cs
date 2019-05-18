using EasyOffice.Attributes;
using EasyOffice.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EasyOffice.Factories
{
    /// <summary>
    /// 过滤器工厂
    /// </summary>
    internal static class FilterFactory
    {
        /// <summary>
        /// 获取类型绑定的过滤器
        /// </summary>
        /// <param name="attrType">过滤器绑定特性类型</param>
        /// <returns>过滤器</returns>
        public static IFilter CreateInstance(Type attrType)
        {
            IFilter filter = null;

            Type filterType = Assembly.GetAssembly(attrType).GetTypes().ToList()?.
                 Where(t => typeof(IFilter).IsAssignableFrom(t))?.
                 FirstOrDefault(t => t.IsDefined(typeof(FilterBindAttribute))
                 && t.GetCustomAttribute<FilterBindAttribute>()?.FilterAttributeType == attrType);

            if (filterType != null)
            {
                filter = (IFilter)Activator.CreateInstance(filterType);
            }

            return filter;
        }
    }
}
