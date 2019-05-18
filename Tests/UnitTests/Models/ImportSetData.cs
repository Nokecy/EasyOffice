using System;
using System.Collections.Generic;
using System.Text;

namespace UnitTests
{
    public class ImportSetData
    {
        public ImportConfig CarCode { get; set; }
        public ImportConfig Mobile { get; set; }
        public ImportConfig IdentityNumber { get; set; }
        public ImportConfig Name{get;set;}
        public ImportConfig Gender { get; set; }
        public ImportConfig RegisterDate { get; set; }
        public ImportConfig Age { get; set; }
    }

    public class ImportConfig
    {
        public string ExcelName { get; set; }
    }
}
