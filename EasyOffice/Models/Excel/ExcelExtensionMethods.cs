using EasyOffice.Attributes;
using EasyOffice.Decorators;
using EasyOffice.Filters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace EasyOffice.Models.Excel
{
    public static class ExcelExtensionMethods
    {
        /// <summary>
        /// 从类型所有的装饰特性中，获取某类型的第一个匹配特性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="typeDecoratorInfo"></param>
        /// <param name="decoratorAttrType"></param>
        /// <returns></returns>
        public static T GetDecorateAttr<T>(this TypeDecoratorInfo typeDecoratorInfo)
            where T : BaseDecorateAttribute
        {
            var attr = typeDecoratorInfo.TypeDecoratorAttrs.SingleOrDefault(a => a.GetType() == typeof(T));
            return attr == null ? null : (T)attr;
        }

        /// <summary>
        /// 从类型所有的装饰特性中，获取某类型的所有匹配特性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="typeDecoratorInfo"></param>
        /// <param name="decoratorAttrType"></param>
        /// <returns></returns>
        //public static List<T> GetDecorateAttrs<T>(this TypeDecoratorInfo typeDecoratorInfo)
        //  where T : BaseDecorateAttribute
        //{
        //    var attrs = typeDecoratorInfo.TypeDecoratorAttrs.Where(a => a.GetType() == typeof(T));
        //    return attrs == null ? null : attrs.Cast<T>().ToList();
        //}

        /// <summary>
        /// 反射获取导出DTO某个属性的值
        /// </summary>
        /// <param name="export"></param>
        /// <param name="propertyName"></param>
        /// <returns></returns>
        public static string GetStringValue<T>(this T export, string propertyName)
        {
            string strVal = string.Empty;
            var prop = export.GetType().GetProperties().Where(p => p.Name.Equals(propertyName)).SingleOrDefault();
            if (prop != null)
            {
                strVal = prop.GetValue(export).ToString();
            }

            return strVal;
        }

        /// <summary>
        /// 获取某单元格的某校验特性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="col"></param>
        /// <param name="typeAttrs"></param>
        /// <returns></returns>
        public static T GetFilterAttr<T>(this ExcelDataCol col, TypeFilterInfo typeFilterInfo)
            where T : BaseFilterAttribute
        {
            return (T)typeFilterInfo.PropertyFilterInfos.SingleOrDefault(a => a.PropertyName.Equals(col.PropertyName, StringComparison.CurrentCultureIgnoreCase))?.
            FilterAttrs?.SingleOrDefault(e => e.GetType() == typeof(T));
        }


        /// <summary>
        /// 获取某单元格的某校验类型集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="col"></param>
        /// <param name="typeFilterInfo"></param>
        /// <returns></returns>
        public static List<T> GetFilterAttrs<T>(this ExcelDataCol col, TypeFilterInfo typeFilterInfo)
           where T : BaseFilterAttribute
        {
            return typeFilterInfo.PropertyFilterInfos.SingleOrDefault(a => a.PropertyName.Equals(col.PropertyName, StringComparison.CurrentCultureIgnoreCase))?.
                   FilterAttrs?.Where(e => e.GetType() == typeof(T)).Cast<T>().ToList();
        }

        /// <summary>
        /// 设置Excel行的校验结果
        /// </summary>
        /// <param name="row"></param>
        /// <param name="isValid"></param>
        /// <param name="col"></param>
        /// <param name="errorMsg"></param>
        public static void SetNotValid(this ExcelDataRow row, bool isValid, ExcelDataCol dataCol, string errorMsg)
        {
            if (!isValid)
            {
                row.IsValid = false;
                row.ErrorMsg += dataCol.ColName + errorMsg + ";";
            }
        }

        /// <summary>
        /// 是否是日期
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        public static bool IsDateTime(this ExcelDataCol col)
        {
            return DateTime.TryParse(col.ColValue, out DateTime dt);
        }

        /// <summary>
        /// 判断数值是否在范围内
        /// </summary>
        /// <param name="col"></param>
        /// <param name="max"></param>
        /// <param name="min"></param>
        /// <returns></returns>
        public static bool IsInRange(this ExcelDataCol col, decimal max, decimal min)
        {
            if (!decimal.TryParse(col.ColValue, out decimal val))
            {
                return false;
            }

            if (val > max || val < min)
            {
                return false;
            }

            return true;
        }

        /// 直接反射
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        //public static T Convert<T>(this ExcelDataRow row)
        //{
        //    return Convert<T>(row, GetValue);
        //}

        /// <summary>
        /// 将ExcelDataRow转换为指定类型
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        //private static T Convert<T>(this ExcelDataRow row, Func<ExcelDataRow, Type, string, object> func)
        //{
        //    Type t = typeof(T);
        //    object o = Activator.CreateInstance(t);
        //    t.GetProperties().ToList().ForEach(p =>
        //    {
        //        if (p.IsDefined(typeof(ColNameAttribute)))
        //        {
        //            p.SetValue(o, func(row, p.PropertyType, p.GetCustomAttribute<ColNameAttribute>().ColName));
        //        }
        //    });

        //    return (T)o;
        //}

        /// <summary>
        /// 利用反射将ExcelDataRow转换为制定类型，性能较差
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        //public static T ConvertByRefelection<T>(this ExcelDataRow row)
        //{
        //    Type t = typeof(T);
        //    object o = Activator.CreateInstance(t);
        //    t.GetProperties().ToList().ForEach(p => {
        //        if (p.IsDefined(typeof(ColNameAttribute)))
        //        {
        //            ExcelDataCol col = row.DataCols.SingleOrDefault(c => c.ColName == p.GetCustomAttribute<ColNameAttribute>().ColName);

        //            if (col != null)
        //            {
        //                p.SetValue(o, ExpressionMapper.ChangeType(col.ColValue, p.PropertyType));
        //            }
        //        }
        //    });

        //    return (T)o;
        //}

        /// <summary>
        /// 利用表达式树，将ExcelDataRow快速转换为指定类型
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <returns></returns>
        public static T FastConvert<T>(this ExcelDataRow row)
        {
            return ExpressionMapper.FastConvert<T>(row);
        }

        /// <summary>
        /// 利用表达式树，将ExcelDataRow集合快速转换为指定类型集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <returns></returns>
        public static IEnumerable<T> FastConvert<T>(this IEnumerable<ExcelDataRow> rows)
        {
            List<T> list = new List<T>();

            rows?.ToList().ForEach(r => list.Add(r.FastConvert<T>()));

            return list;
        }

        //private static object GetValue(ExcelDataRow row,Type propType ,string colName)
        //{
        //    string val = row.DataCols.SingleOrDefault(c => c.ColName == colName)?.ColValue;
        //    if (!string.IsNullOrWhiteSpace(val))
        //    {
        //        return ExpressionMapper.ChangeType(val, propType);
        //    }

        //    return val;
        //}
    }
}
