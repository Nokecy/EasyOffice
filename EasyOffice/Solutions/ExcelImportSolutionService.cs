using EasyOffice.Attributes;
using EasyOffice.Interfaces;
using EasyOffice.Models.Excel;
using EasyOffice.Solutions.Models;
using EasyOffice.Solutions.Repositories;
using Bayantu.Extensions.Session;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace EasyOffice.Solutions
{
    public class ExcelImportSolutionService : IExcelImportSolutionService
    {
        private readonly IExcelImportService _excelImportService;
        private readonly IExcelExportService _excelExportService;
        private readonly IImportErrorLogRepository _importErrorLogRepository;
        public ExcelImportSolutionService(IExcelImportService excelImportService
            , IImportErrorLogRepository importErrorLogRepository
            , IExcelExportService excelExportService
            )
        {
            _excelImportService = excelImportService;
            _importErrorLogRepository = importErrorLogRepository;
            _excelExportService = excelExportService;
        }

        /// <inheritdoc/>
        public Task<ImportConfig> GetImportConfigAsync<T>(string uploadUrl = "", string templateUrl = "") where T : class, new()
        {
            var result = new ImportConfig()
            {
                UploadUrl = uploadUrl,
                TemplateUrl = templateUrl,
                Property = GetImportPropertyConfigs(typeof(T))
            };

            return Task.FromResult(result);
        }

        /// <inheritdoc/>
        public async Task<List<List<string>>> GetFileHeadersAndRowsAsync<T>(string fileUrl, int sheetIndex = 0, int dataCount = 3, int headerRowIndex = 0, int dataRowIndex = 1, List<string> headerData = null)
            where T : class, new()
        {
            var table = await _excelImportService.ToTableAsync<T>(fileUrl, sheetIndex, headerRowIndex, dataRowIndex, dataCount);

            var result = new List<List<string>>();

            if (headerData == null)
            {
                headerData = new List<string>();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    headerData.Add(table.Columns[i].ColumnName);
                }
            }

            result.Add(headerData);

            for (int i = 0; i < table.Rows.Count; i++)
            {
                var rowData = new List<string>();
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    rowData.Add(table.Rows[i][j].ToString());
                }

                result.Add(rowData);
            }

            if (!ValidateTemplate<T>(result, out string errorMsg))
            {
                throw new InvalidOperationException(errorMsg);
            }

            return result;
        }

        /// <inheritdoc/>
        public async Task<ImportResult> ImportAsync<T>(ImportOption importOption, object importSetData, Action<List<T>> businessAction, Func<List<ExcelDataRow>, List<ExcelDataRow>> customValidateAction = null)
            where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(importOption.FileUrl))
            {
                throw new ArgumentNullException("fileUrl");
            }

            if (businessAction == null)
            {
                throw new ArgumentNullException("businessAction");
            }

            if (importSetData == null)
            {
                throw new ArgumentNullException("importSetData");
            }
            //获取表头映射字典
            var mappdingDictionary = GetHeaderDictFromImportSetData(importSetData);

            //校验数据
            importOption.MappingDictionary = mappdingDictionary;
            var validateResult = await _excelImportService.ValidateAsync<T>(importOption);

            //执行自定义校验委托
            if (customValidateAction != null)
            {
                validateResult = customValidateAction.Invoke(validateResult);
            }

            //对有效数据执行业务操作
            var validRow = validateResult.Where(x => x.IsValid);
            var validDatas = validRow.FastConvert<T>().ToList();
            businessAction.Invoke(validDatas);

            //导入结果
            var result = new ImportResult()
            {
                Total = validateResult.Count,
                OperateSuccess = true,
                Message = "操作成功",
                Success = validateResult.Where(x => x.IsValid).Count(),
                Failed = validateResult.Where(x => !x.IsValid).Count(),
                MaxErrorMsgCount = 100,
                Tag = Guid.NewGuid().ToString(),
                ErrorMsgs = validateResult.Where(x => !x.IsValid).Select(x => $"第{x.RowIndex + 1}行，{x.ErrorMsg}").ToList()
            };

            //错误消息入库
            await _importErrorLogRepository.AddImportErrorLogAsync(validateResult, result.Tag);

            //返回导入结果
            return result;
        }

        /// <summary>
        /// 校验导入的模板是否正确
        /// </summary>
        /// <typeparam name="T">模板类类型</typeparam>
        /// <param name="listHeaderAndRow">Excel模板数据</param>
        /// <returns>bool</returns>
        private bool ValidateTemplate<T>(List<List<string>> listHeaderAndRow, out string errorMsg)
            where T : class, new()
        {
            bool isValid = true;
            errorMsg = string.Empty;

            //Excel模板是否有数据
            if (listHeaderAndRow.Count == 0)
            {
                errorMsg = "Excel模板没有任何数据";
                return false;
            }

            //表头行与数据行的列数是否匹配
            if (listHeaderAndRow.Count > 1)
            {
                int colCount = listHeaderAndRow[0].Count;
                for (int i = 1; i < listHeaderAndRow.Count; i++)
                {
                    if (listHeaderAndRow[i].Count != colCount)
                    {
                        errorMsg = $"第{i}行数据行与表头行列数不匹配";
                        return false;
                    }
                }
            }

            //表头是否存在任何空白单元格
            if (listHeaderAndRow[0].Any(e => string.IsNullOrWhiteSpace(e)))
            {
                errorMsg = "Excel表头行存在空白单元格";
                return false;
            }

            return isValid;
        }

        /// <summary>
        /// 获取导入模板类类型的配置信息
        /// </summary>
        /// <param name="type">导入模板类类型</param>
        /// <returns>配置信息集合</returns>
        private List<ImportPropertyConfig> GetImportPropertyConfigs(Type type)
        {
            var results = new List<ImportPropertyConfig>();
            if (type == null)
            {
                return results;
            }

            PropertyInfo[] infos = type.GetProperties();
            if (infos == null)
            {
                return results;
            }

            foreach (var info in infos)
            {
                if (info.IsDefined(typeof(ColNameAttribute)))
                {
                    var result = new ImportPropertyConfig()
                    {
                        Field = info.Name,
                        DisplayName = info.GetCustomAttribute<ColNameAttribute>().ColName,
                        IsRequired = false
                    };

                    if (info.IsDefined(typeof(RequiredAttribute)))
                    {
                        result.IsRequired = true;
                    }

                    results.Add(result);
                }
            }

            return results;
        }

        private Dictionary<string, string> GetHeaderDictFromImportSetData(object importSetData)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();

            string s = JsonConvert.SerializeObject(importSetData);
            dynamic dynObj = JsonConvert.DeserializeObject<dynamic>(s);
            var jObj = (JObject)dynObj;

            foreach (JToken token in jObj.Children())
            {
                if (token is JProperty)
                {
                    var prop = token as JProperty;
                    dict.Add(prop.Name, string.Empty);
                    var j = (JObject)prop.Value;
                    foreach (JToken t in j.Children())
                    {
                        if (t is JProperty)
                        {
                            var p = t as JProperty;
                            if (p.Name == "ExcelName")
                            {
                                dict[prop.Name] = p.Value.ToString();
                            }
                        }
                    }
                }
            }

            return dict;
        }

        public async Task<byte[]> ExportErrorMsgAsync(string tag)
        {
            var importLogs = await _importErrorLogRepository.GetErrorLogsByTagAsync(tag);
            List<ImportErrorLogDTO> dtos = new List<ImportErrorLogDTO>();

            importLogs.ToList().ForEach(x =>
            {
                dtos.Add(new ImportErrorLogDTO()
                {
                    ErrorMsg = $"第{x.RowNumber}行，{x.Message}"
                });
            });

            return await _excelExportService.ExportAsync(new ExportOption<ImportErrorLogDTO>()
            {
                Data = dtos,
            });
        }

        public async Task<byte[]> GetImportTemplateAsync<T>() where T : class, new()
        {
            return await _excelExportService.ExportAsync(new ExportOption<T>()
            {
                Data = new List<T>()
            });
        }
    }
}
