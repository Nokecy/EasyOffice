using EasyOffice.Filters;
using System.Collections.Generic;
using System.Reflection;

namespace EasyOffice.Models.Excel
{
    public class Rule
    {
        public PropertyInfo Property { get; set; }
        public List<BaseFilter> Filters { get; set; } = new List<BaseFilter>();

        public Rule NotEmpty(string errorMsg = "必填")
        {
            var filter = new RequiredFilter()
            {
                ErrorMsg = errorMsg
            };

            Filters.Add(filter);

            return this;
        }

        public Rule IsDateTime(string errorMsg = "非法日期")
        {
            var filter = new DateTimeFilter()
            {
                ErrorMsg = errorMsg
            };

            Filters.Add(filter);

            return this;
        }

        public Rule NotDuplicate(string errorMsg = "重复")
        {
            var filter = new DuplicateFilter()
            {
                ErrorMsg = errorMsg
            };

            Filters.Add(filter);

            return this;
        }

        public Rule MaxLength(int maxLength,string errorMsg = "重复")
        {
            var filter = new MaxLengthFilter()
            {
                ErrorMsg = errorMsg,
                MaxLength = maxLength
            };

            Filters.Add(filter);

            return this;
        }

        public Rule Range(int min,int max, string errorMsg = "")
        {
            if (string.IsNullOrWhiteSpace(errorMsg))
            {
                errorMsg = $"超限，仅允许为{min}-{max}";
            }

            var filter = new RangeFilter()
            {
                ErrorMsg = errorMsg,
                Min = min,
                Max = max
            };

            Filters.Add(filter);

            return this;
        }

        public Rule Regex(string regex, string errorMsg = "重复")
        {
            var filter = new RegexFilter()
            {
                ErrorMsg = errorMsg,
                RegexString = regex
            };

            Filters.Add(filter);

            return this;
        }
    }
}
