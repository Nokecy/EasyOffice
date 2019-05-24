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

namespace EasyOffice.Models.Excel
{
    /// <summary>
    /// 生成表达式目录树 缓存
    /// </summary>
    public class ExpressionMapper
    {
        private static Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// 将ExcelDataRow快速转换为指定类型
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataRow"></param>
        /// <returns></returns>
        public static T FastConvert<T>(ExcelDataRow dataRow)
        {
            //利用表达式树，动态生成委托并缓存，得到接近于硬编码的性能
            //最终生成的代码近似于(假设T为Person类)
            //Func<ExcelDataRow,Person>
            //      new Person(){
            //          Name = Convert(ChangeType(dataRow.DataCols.SingleOrDefault(c=>c.PropertyName == prop.Name).ColValue,prop.PropertyType),prop.ProertyType),
            //          Age = Convert(ChangeType(dataRow.DataCols.SingleOrDefault(c=>c.PropertyName == prop.Name).ColValue,prop.PropertyType),prop.ProertyType)
            //      }
            // }

            string propertyNames = string.Empty;
            dataRow.DataCols.ForEach(c => propertyNames += c.PropertyName + "_");
            var key = typeof(T).FullName + "_" + propertyNames.Trim('_');


            if (!Table.ContainsKey(key))
            {
                List<MemberBinding> memberBindingList = new List<MemberBinding>();

                MethodInfo firstOrDefaultMethod = typeof(Enumerable)
                                                            .GetMethods()
                                                            .Single(m => m.Name == "FirstOrDefault" && m.GetParameters().Count() == 2)
                                                            .MakeGenericMethod(new[] { typeof(ExcelDataCol) });

                foreach (var prop in typeof(T).GetProperties())
                {
                    Expression<Func<ExcelDataCol, bool>> lambdaExpr = c => c.PropertyName == prop.Name;

                    MethodInfo changeTypeMethod = typeof(ExpressionMapper).GetMethods().Where(m => m.Name == "ChangeType").First();

                    Expression expr =
                        Expression.Convert(
                            Expression.Call(changeTypeMethod
                                , Expression.Property(
                                    Expression.Call(
                                          firstOrDefaultMethod
                                        , Expression.Variable(typeof(IEnumerable<ExcelDataCol>))
                                        , lambdaExpr)
                                        , typeof(ExcelDataCol), "ColValue"), Expression.Constant(prop.PropertyType))
                                    , prop.PropertyType);

                    memberBindingList.Add(Expression.Bind(prop, expr));
                }

                MemberInitExpression memberInitExpression = Expression.MemberInit(Expression.New(typeof(T)), memberBindingList.ToArray());
                Expression<Func<ExcelDataRow, T>> lambda = Expression.Lambda<Func<ExcelDataRow, T>>(memberInitExpression, new ParameterExpression[]
                {
                    Expression.Parameter(typeof(ExcelDataRow), "p")
                });

                Func<ExcelDataRow, T> func = lambda.Compile();//拼装是一次性的
                Table[key] = func;
            }

            var result = ((Func<ExcelDataRow, T>)Table[key]).Invoke(dataRow);

            return result;
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
