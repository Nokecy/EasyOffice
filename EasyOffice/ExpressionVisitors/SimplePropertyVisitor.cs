using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace EasyOffice.ExpressionVisitors
{
    public class SimplePropertyVisitor : ExpressionVisitor
    {
        public PropertyInfo PropertyInfo { get; set; }

        protected override Expression VisitMember(MemberExpression node)
        {
            if (node.Member.MemberType == MemberTypes.Property)
            {
                PropertyInfo = (PropertyInfo)node.Member;
            }

            return base.VisitMember(node);
        }
    }
}
