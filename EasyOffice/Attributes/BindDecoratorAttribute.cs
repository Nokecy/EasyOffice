using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Attributes
{
    /// <summary>
    /// 装饰器绑定属性
    /// </summary>
    public class BindDecoratorAttribute : Attribute
    {
        public BindDecoratorAttribute(Type decoratorType)
        {
            DecoratorType = decoratorType;
        }

        /// <summary>
        /// 装饰器类型
        /// </summary>
        public Type DecoratorType { get; set; }
    }
}
