using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Solutions.Models
{
    public class ImportConfig
    {
        /// <summary>
        /// 导入服务地址
        /// </summary>
        public string UploadUrl { get; set; }

        /// <summary>
        /// 获取导入服务地址模板
        /// </summary>
        public string TemplateUrl { get; set; }

        /// <summary>
        /// 导入字段配置
        /// </summary>
        public List<ImportPropertyConfig> Property { get; set; }
    }
}
