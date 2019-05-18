using Autofac;
using Bayantu.Extensions.DependencyInjection;
using EasyOffice;
using NSubstitute;
using System;
using System.Reflection;

namespace UnitTests.DependencyInjection
{
    public static class DependencyContainer
    {
        /// <summary>
        /// 容器
        /// </summary>
        private static readonly Autofac.IContainer Container;

        /// <summary>
        /// 构造函数进行容器注册
        /// </summary>
        static DependencyContainer()
        {
            var builder = new ContainerBuilder();

            builder.AddOffice(new OfficeOptions());

            Container = builder.Build();
        }

        /// <summary>
        /// 返回实例
        /// </summary>
        /// <returns></returns>
        public static TComponent GetComponent<TComponent>()
        {
            return Container.Resolve<TComponent>();
        }

        /// <summary>
        /// 根据类型返回Object
        /// </summary>
        /// <remarks>
        /// 给单元测试用
        /// </remarks>
        /// <param name="type"></param>
        /// <returns></returns>
        public static object Resolve(Type type)
        {
            return Container.Resolve(type);
        }

        /// <summary>
        /// 利用类型推断和扩展方法简化Resolve的写法
        /// </summary>
        /// <typeparam name="TComponent"></typeparam>
        /// <param name="component"></param>
        /// <returns></returns>
        public static TComponent Resolve<TComponent>(this TComponent component)
        {
            return Container.Resolve<TComponent>();
        }

        /// <summary>
        /// 利用类型推断和扩展方法简化Resolve的写法
        /// </summary>
        /// <typeparam name="TComponent"></typeparam>
        /// <param name="component"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        public static TComponent ResolveNamed<TComponent>(this TComponent component, string name)
        {
            return Container.ResolveNamed<TComponent>(name);
        }

    }
}
