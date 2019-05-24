using EasyOffice.Models.Excel;
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


        public static List<ExcelDataRow> GetExcelDataRows(int rowsCount,int colsCount)
        {
            var rows = new List<ExcelDataRow>();

            for (int i = 0; i < rowsCount; i++)
            {
                var row = new ExcelDataRow()
                {
                    DataCols = new List<ExcelDataCol>(),
                    IsValid = true
                };
                for (int j = 0; j < colsCount; j++)
                {
                    var col = new ExcelDataCol()
                    {
                        ColIndex = j,
                        ColName = $"p{j + 1}",
                        PropertyName = $"p{j + 1}",
                        RowIndex = i,
                        ColValue = $"r{i + 1}c{j + 1}"
                    };

                    row.DataCols.Add(col);
                }
                rows.Add(row);
            }

            return rows;
        }
    }
}
