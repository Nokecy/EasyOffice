using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EasyOffice.Filters
{
    public class FilterContext
    {

        /// <summary>
        /// DTO所有校验特性集合
        /// </summary>
        public TypeFilterInfo TypeFilterInfo { get; set; }


        /// <summary>
        /// 数据库校验委托(参数为表名，字段名)
        /// </summary>
        //public Func<DatabaseFilterContext, bool> DelegateNotExistInDatabase { get; set; }
    }
}
