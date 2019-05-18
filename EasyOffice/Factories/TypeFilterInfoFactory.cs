using EasyOffice.Attributes;
using EasyOffice.Filters;
using EasyOffice.Models.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EasyOffice.Factories
{
    /// <summary>
    /// 模板类过滤器信息工厂
    /// </summary>
    public static class TypeFilterInfoFactory
    {
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// 获取模板类过滤器信息
        /// </summary>
        /// <param name="importType">导入模板类</param>
        /// <param name="excelHeaderRow">Excel表头</param>
        /// <returns>过滤器信息</returns>
        public static TypeFilterInfo CreateInstance(Type importType, ExcelHeaderRow excelHeaderRow)
        {
            if (importType == null)
            {
                throw new ArgumentNullException("importDTOType");
            }

            if (excelHeaderRow == null)
            {
                throw new ArgumentNullException("excelHeaderRow");
            }

            var key = importType;
            if (Table[key] != null)
            {
                return (TypeFilterInfo)Table[key];
            }

            TypeFilterInfo typeFilterInfo = new TypeFilterInfo() { PropertyFilterInfos = new List<PropertyFilterInfo>() { } };

            IEnumerable<PropertyInfo> props = importType.GetProperties().ToList().Where(p => p.IsDefined(typeof(ColNameAttribute)));
            props.ToList().ForEach(p =>
            {
                typeFilterInfo.PropertyFilterInfos.Add(
                    new PropertyFilterInfo
                    {
                        PropertyName = p.Name,
                        FilterAttrs = p.GetCustomAttributes<BaseFilterAttribute>()?.ToList()
                    });
            });

            Table[key] = typeFilterInfo;

            return typeFilterInfo;
        }
    }
}
