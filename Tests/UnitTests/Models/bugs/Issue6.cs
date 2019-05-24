using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace UnitTests.Models.bugs
{
    public class Issue6
    {
        [ColName("姓名")]
        public string Name { get;set;}

        [ColName("年龄")]
        public string Age { get; set; }
    }
}
