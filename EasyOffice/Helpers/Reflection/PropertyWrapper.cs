using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EasyOffice.Helpers.Reflection
{
    /// <summary>
    /// 一些扩展方法，用于访问属性，它们都可以优化反射性能。
    /// </summary>
    public static class PropertyExtensions
    {
        /// <summary>
        /// 快速调用PropertyInfo的GetValue方法
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static object FastGetValue(this PropertyInfo propertyInfo, object obj)
        {
            if (propertyInfo == null)
            {
                throw new ArgumentNullException("propertyInfo");
            }

            return GetterSetterFactory.GetPropertyGetterWrapper(propertyInfo).Get(obj);
        }

        /// <summary>
        /// 快速调用PropertyInfo的SetValue方法
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <param name="obj"></param>
        /// <param name="value"></param>
        public static void FastSetValue(this PropertyInfo propertyInfo, object obj, object value)
        {
            if (propertyInfo == null)
            {
                throw new ArgumentNullException("propertyInfo");
            }

            GetterSetterFactory.GetPropertySetterWrapper(propertyInfo).Set(obj, value);
        }
    }

    /// <summary>
    /// 定义读属性操作的接口
    /// </summary>
    internal interface IGetValue
    {
        object Get(object target);
    }

    /// <summary>
    /// 定义写属性操作的接口
    /// </summary>
    internal interface ISetValue
    {
        void Set(object target, object val);
    }

    /// <summary>
    /// 创建IGetValue或者ISetValue实例的工厂方法类
    /// </summary>
    internal static class GetterSetterFactory
    {
        private static readonly Hashtable GetterDict = Hashtable.Synchronized(new Hashtable(10240));
        private static readonly Hashtable SetterDict = Hashtable.Synchronized(new Hashtable(10240));

        /// <summary>
		/// 根据指定的PropertyInfo对象，返回对应的IGetValue实例
		/// </summary>
		/// <param name="propertyInfo"></param>
		/// <returns></returns>
		public static IGetValue CreatePropertyGetterWrapper(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
            {
                throw new ArgumentNullException("propertyInfo");
            }

            if (propertyInfo.CanRead == false)
            {
                throw new NotSupportedException("属性不支持读操作。");
            }

            MethodInfo mi = propertyInfo.GetGetMethod(true);

            if (mi.GetParameters().Length > 0)
            {
                throw new NotSupportedException("不支持构造索引器属性的委托。");
            }

            if (mi.IsStatic)
            {
                Type instanceType = typeof(StaticGetterWrapper<>).MakeGenericType(propertyInfo.PropertyType);
                return (IGetValue)Activator.CreateInstance(instanceType, propertyInfo);
            }
            else
            {
                Type instanceType = typeof(GetterWrapper<,>).MakeGenericType(propertyInfo.DeclaringType, propertyInfo.PropertyType);
                return (IGetValue)Activator.CreateInstance(instanceType, propertyInfo);
            }
        }

        /// <summary>
        /// 根据指定的PropertyInfo对象，返回对应的ISetValue实例
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <returns></returns>
        public static ISetValue CreatePropertySetterWrapper(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
            {
                throw new ArgumentNullException("propertyInfo");
            }

            if (propertyInfo.CanWrite == false)
            {
                throw new NotSupportedException("属性不支持写操作。");
            }

            MethodInfo mi = propertyInfo.GetSetMethod(true);

            if (mi.GetParameters().Length > 1)
            {
                throw new NotSupportedException("不支持构造索引器属性的委托。");
            }

            if (mi.IsStatic)
            {
                Type instanceType = typeof(StaticSetterWrapper<>).MakeGenericType(propertyInfo.PropertyType);
                return (ISetValue)Activator.CreateInstance(instanceType, propertyInfo);
            }
            else
            {
                Type instanceType = typeof(SetterWrapper<,>).MakeGenericType(propertyInfo.DeclaringType, propertyInfo.PropertyType);
                return (ISetValue)Activator.CreateInstance(instanceType, propertyInfo);
            }
        }

        internal static IGetValue GetPropertyGetterWrapper(PropertyInfo propertyInfo)
        {
            IGetValue property = (IGetValue)GetterDict[propertyInfo];
            if (property == null)
            {
                property = CreatePropertyGetterWrapper(propertyInfo);
                GetterDict[propertyInfo] = property;
            }

            return property;
        }

        internal static ISetValue GetPropertySetterWrapper(PropertyInfo propertyInfo)
        {
            ISetValue property = (ISetValue)SetterDict[propertyInfo];
            if (property == null)
            {
                property = CreatePropertySetterWrapper(propertyInfo);
                SetterDict[propertyInfo] = property;
            }

            return property;
        }
    }

    internal class GetterWrapper<TTarget, TValue> : IGetValue
    {
        private Func<TTarget, TValue> getter;

        public GetterWrapper(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
            {
                throw new ArgumentNullException("propertyInfo");
            }

            if (propertyInfo.CanRead == false)
            {
                throw new NotSupportedException("属性不支持读操作。");
            }

            if (propertyInfo.GetIndexParameters().Length > 0)
            {
                throw new NotSupportedException("不支持构造索引器属性的委托。");
            }

            MethodInfo m = propertyInfo.GetGetMethod(true);
            getter = (Func<TTarget, TValue>)Delegate.CreateDelegate(typeof(Func<TTarget, TValue>), null, m);
        }

        public TValue GetValue(TTarget target)
        {
            return getter(target);
        }

        object IGetValue.Get(object target)
        {
            return getter((TTarget)target);
        }
    }

    internal class SetterWrapper<TTarget, TValue> : ISetValue
    {
        private Action<TTarget, TValue> setter;

        public SetterWrapper(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
            {
                throw new ArgumentNullException("propertyInfo");
            }

            if (propertyInfo.CanWrite == false)
            {
                throw new NotSupportedException("属性不支持写操作。");
            }

            if (propertyInfo.GetIndexParameters().Length > 0)
            {
                throw new NotSupportedException("不支持构造索引器属性的委托。");
            }

            MethodInfo m = propertyInfo.GetSetMethod(true);
            setter = (Action<TTarget, TValue>)Delegate.CreateDelegate(typeof(Action<TTarget, TValue>), null, m);
        }

        public void SetValue(TTarget target, TValue val)
        {
            setter(target, val);
        }

        void ISetValue.Set(object target, object val)
        {
            setter((TTarget)target, (TValue)val);
        }
    }

    internal class StaticGetterWrapper<TValue> : IGetValue
    {
        private Func<TValue> getter;

        public StaticGetterWrapper(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
            {
                throw new ArgumentNullException("propertyInfo");
            }

            if (propertyInfo.CanRead == false)
            {
                throw new NotSupportedException("属性不支持读操作。");
            }

            if (propertyInfo.GetIndexParameters().Length > 0)
            {
                throw new NotSupportedException("不支持构造索引器属性的委托。");
            }

            MethodInfo m = propertyInfo.GetGetMethod(true);

            getter = (Func<TValue>)Delegate.CreateDelegate(typeof(Func<TValue>), m);
        }

        public TValue GetValue()
        {
            return getter();
        }

        object IGetValue.Get(object target)
        {
            return getter();
        }
    }

    internal class StaticSetterWrapper<TValue> : ISetValue
    {
        private Action<TValue> setter;

        public StaticSetterWrapper(PropertyInfo propertyInfo)
        {
            if (propertyInfo == null)
            {
                throw new ArgumentNullException("propertyInfo");
            }

            if (propertyInfo.CanWrite == false)
            {
                throw new NotSupportedException("属性不支持写操作。");
            }

            if (propertyInfo.GetIndexParameters().Length > 0)
            {
                throw new NotSupportedException("不支持构造索引器属性的委托。");
            }

            MethodInfo m = propertyInfo.GetSetMethod(true);
            setter = (Action<TValue>)Delegate.CreateDelegate(typeof(Action<TValue>), m);
        }

        public void SetValue(TValue val)
        {
            setter(val);
        }

        void ISetValue.Set(object target, object val)
        {
            setter((TValue)val);
        }
    }
}
