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
    /// 装饰器工厂
    /// </summary>
    public static class DecoratorFactory
    {
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

        /// <summary>
        /// /获取类型所有装饰器
        /// </summary>
        /// <param name="attrType"></param>
        /// <returns></returns>
        public static IDecorator CreateInstance(Type attrType)
        {
            if (Table[attrType] != null)
            {
                return (IDecorator)Table[attrType];
            }

            IDecorator filter = null;

            Type decoratorType = Assembly.GetAssembly(attrType).GetTypes().ToList()?.
                 Where(t => typeof(IDecorator).IsAssignableFrom(t))?.
                 FirstOrDefault(t => t.IsDefined(typeof(BindDecoratorAttribute))
                 && t.GetCustomAttribute<BindDecoratorAttribute>()?.DecoratorType == attrType);

            if (decoratorType != null)
            {
                filter = (IDecorator)Activator.CreateInstance(decoratorType);
                Table[attrType] = filter;
            }

            return filter;
        }
    }
}
