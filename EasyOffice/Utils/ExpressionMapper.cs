using EasyOffice.Models.Excel;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace EasyOffice.Utils
{
    /// <summary>
    /// 生成表达式目录树 缓存
    /// </summary>
    public static class ExpressionMapper
    {
        private static Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// 将ExcelDataRow快速转换为指定类型
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataRow"></param>
        /// <returns></returns>
        public static T FastConvert<T>(ExcelDataRow dataRow, Func<List<ExcelDataCol>, T> func)
        {
            //利用表达式树，动态生成委托并缓存，得到接近于硬编码的性能
            //最终生成的代码近似于(假设T为Person类)
            //Func<ExcelDataRow,Person>
            //      new Person(){
            //          Name = Convert(ChangeType(dataRow.DataCols.SingleOrDefault(c=>c.PropertyName == prop.Name).ColValue,prop.PropertyType),prop.ProertyType),
            //          Age = Convert(ChangeType(dataRow.DataCols.SingleOrDefault(c=>c.PropertyName == prop.Name).ColValue,prop.PropertyType),prop.ProertyType)
            //      }
            // }
            return func.Invoke(dataRow.DataCols);
        }

        public static Func<List<ExcelDataCol>, T> GetFunc<T>(string key, IEnumerable<PropertyInfo> props)
        {
            if (!Table.ContainsKey(key))
            {
                Expression<Func<IEnumerable<FieldInfo>, Func<FieldInfo, bool>, FieldInfo>> singleOrDefaultExpr = (l, p) => l.SingleOrDefault(p);

                List<MemberBinding> memberBindingList = new List<MemberBinding>();

                //得到FirstOrDefault方法
                MethodInfo firstOrDefaultMethod = typeof(Enumerable)
                                                            .GetMethods()
                                                            .Single(m => m.Name == "FirstOrDefault" && m.GetParameters().Count() == 2)
                                                            .MakeGenericMethod(new[] { typeof(ExcelDataCol) });


                var dataColsParam = Expression.Parameter(typeof(List<ExcelDataCol>), "dataCols");

                foreach (var prop in props)
                {
                    //lambda表达式： PropertyName=prop.Name
                    Expression<Func<ExcelDataCol, bool>> propertyEqualExpr = c => c.PropertyName == prop.Name;

                    //调用ChangeType方法
                    MethodInfo changeTypeMethod = typeof(ExpressionMapper).GetMethods().Where(m => m.Name == "ChangeType").First();

                    //firstOrDefault方法
                    var firstOrDefaultMethodExpr = Expression.Call(firstOrDefaultMethod, dataColsParam, propertyEqualExpr);

                    //当前propertytype
                    var propTypeConst = Expression.Constant(prop.PropertyType);

                    //得到DataCols.SingleOrDefault(c=>c.PropertyName == prop.Name).ColValue
                    var colValueExpr = Expression.Property(firstOrDefaultMethodExpr, typeof(ExcelDataCol), "ColValue");

                    //changeType表达式
                    var changeTypeExpr = Expression.Call(changeTypeMethod, colValueExpr, propTypeConst);

                    Expression expr = Expression.Convert(changeTypeExpr, prop.PropertyType);

                    memberBindingList.Add(Expression.Bind(prop, expr));
                }

                MemberInitExpression memberInitExpression = Expression.MemberInit(Expression.New(typeof(T)), memberBindingList.ToArray());

                var blockExpr = Expression.Block(memberInitExpression);

                Expression<Func<List<ExcelDataCol>, T>> lambda = Expression.Lambda<Func<List<ExcelDataCol>, T>>(blockExpr, new ParameterExpression[]
                {
                  dataColsParam
                });

                Func<List<ExcelDataCol>, T> func = lambda.Compile();
                Table[key] = func;
            }

            return (Func<List<ExcelDataCol>, T>)Table[key];
        }

        public static object ChangeType(string stringValue, Type type)
        {
            object obj = null;

            Type nullableType = Nullable.GetUnderlyingType(type);
            if (nullableType != null)
            {
                if (stringValue == null)
                {
                    obj = null;
                }

            }
            else if (typeof(Enum).IsAssignableFrom(type))
            {
                obj = Enum.Parse(type, stringValue);
            }
            else
            {
                obj = Convert.ChangeType(stringValue, type);
            }

            return obj;
        }
    }
}
