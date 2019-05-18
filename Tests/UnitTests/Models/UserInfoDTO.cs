using EasyOffice.Attributes;
using System;
using System.Collections.Generic;
using System.Text;

namespace UnitTests.Models
{
    public class UserInfoDTO
    {
        public string Name { get; set; }

        public string Gender { get; set; }
        public int Age { get; set; }
        public string Remarks { get; set; }
    }
}
