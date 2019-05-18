using System;
using System.Collections.Generic;
using System.Text;

namespace UnitTests
{
    public static class DataLib
    {
        public static ImportSetData GetImportSetData()
        {
            return new ImportSetData()
            {
                Age = new ImportConfig()
                {
                    ExcelName = "年龄"
                },
                CarCode = new ImportConfig()
                {
                    ExcelName = "车牌号"
                },
                Gender = new ImportConfig()
                {
                    ExcelName = "性别"
                },
                IdentityNumber = new ImportConfig()
                {
                    ExcelName = "身份证号"
                },
                Mobile = new ImportConfig()
                {
                    ExcelName = "手机号"
                },
                Name = new ImportConfig()
                {
                    ExcelName = "姓名"
                },
                RegisterDate = new ImportConfig()
                {
                    ExcelName = "注册日期"
                }
            };
        }
    }
}
