using Autofac;
using Bayantu.Extensions.DependencyInjection;
using EasyOffice;
using Microsoft.Extensions.DependencyInjection;
using NSubstitute;
using System;
using System.Reflection;

namespace UnitTests.DependencyInjection
{
    public static class DependencyContainerServiceCollection
    {
        /// <summary>
        /// 容器
        /// </summary>
        private static readonly IServiceProvider _serviceProvider;

        /// <summary>
        /// 构造函数进行容器注册
        /// </summary>
        static DependencyContainerServiceCollection()
        {
            var serviceCollection = new ServiceCollection();

            serviceCollection.AddOffice(new OfficeOptions());

            _serviceProvider = serviceCollection.BuildServiceProvider();
        }

        /// <summary>
        /// 利用类型推断和扩展方法简化Resolve的写法
        /// </summary>
        /// <typeparam name="TComponent"></typeparam>
        /// <param name="component"></param>
        /// <returns></returns>
        public static TComponent ResolveByServiceCollection<TComponent>(this TComponent component)
        {
            return _serviceProvider.GetRequiredService<TComponent>();
        }
    }
}
