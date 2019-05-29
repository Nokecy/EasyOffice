using EasyOffice.ExpressionVisitors;
using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;

namespace EasyOffice.Models.Excel
{
    public class Validator<TTemplate>
    {
        public IEnumerable<Rule> Rules { get; set; }

        public Rule AddRule(Expression<Func<TTemplate,object>> expr)
        {
            SimplePropertyVisitor visitor = new SimplePropertyVisitor();
            visitor.Visit(expr);

            var rule = new Rule();

            rule.Property = visitor.PropertyInfo;

            return rule;
        }
    }
}
