using EasyOffice.Attributes;
using EasyOffice.Filters;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyOffice.Factories
{
    /// <summary>
    /// 模板类过滤器享元工厂
    /// </summary>
    public static class FiltersFactory
    {
        private static Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// Excel表头与模板类匹配获取所有过滤器
        /// </summary>
        /// <typeparam name="TTemplate"></typeparam>
        /// <param name="headerRow"></param>
        /// <returns></returns>
        public static List<IFilter> CreateFilters<TTemplate>(ExcelHeaderRow headerRow)
        {
            Type templateType = typeof(TTemplate);
            var key = templateType;
            if (Table[key] != null)
            {
                return (List<IFilter>)Table[key];
            }

            List<IFilter> filters = new List<IFilter>();
            List<BaseFilterAttribute> attrs = new List<BaseFilterAttribute>();
            TypeFilterInfo typeFilterInfo = TypeFilterInfoFactory.CreateInstance(typeof(TTemplate), headerRow);

            typeFilterInfo.PropertyFilterInfos.ForEach(a => a.FilterAttrs.ForEach(f => attrs.Add(f)));

            attrs.Distinct(new FilterAttributeComparer()).ToList().ForEach
            (a =>
            {
                var filter = FilterFactory.CreateInstance(a.GetType());
                if (filter != null)
                {
                    filters.Add(filter);
                }
            });

            Table[key] = filters;
            return filters;
        }
    }
}
