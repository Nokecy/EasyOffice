using EasyOffice.Attributes;
using EasyOffice.Models.Word;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace EasyOffice.Utils
{
    public static class WordHelper
    {
        /// <summary>
        /// 获取 占位符：值 字典
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetReplacements<T>(T wordData)
           where T : class, new()
        {
            Dictionary<string, string> replacements = new Dictionary<string, string>();
            Type type = typeof(T);
            PropertyInfo[] props = type.GetProperties();

            foreach (PropertyInfo prop in props)
            {
                if (prop.PropertyType == typeof(Picture) || typeof(IEnumerable<Picture>).IsAssignableFrom(prop.PropertyType))
                    break;

                var replacement = prop.GetValue(wordData)?.ToString();

                var placeholder = prop.IsDefined(typeof(PlaceholderAttribute)) ?
                   prop.GetCustomAttribute<PlaceholderAttribute>().Placeholder.ToString() : "{" + prop.Name + "}";

                replacements.Add(placeholder, replacement);
            }

            return replacements;
        }

        public static Dictionary<string, IEnumerable<Picture>> GetPictureReplacements<T>(T wordData)
          where T : class, new()
        {
            Dictionary<string, IEnumerable<Picture>> replacements = new Dictionary<string, IEnumerable<Picture>>();
            Type type = typeof(T);
            PropertyInfo[] props = type.GetProperties();

            foreach (PropertyInfo prop in props)
            {
                var placeholder = prop.IsDefined(typeof(PlaceholderAttribute)) ?
                   prop.GetCustomAttribute<PlaceholderAttribute>().Placeholder.ToString() : "{" + prop.Name + "}";

                if (prop.PropertyType == typeof(Picture))
                {
                    var picture = (Picture)prop.GetValue(wordData);
                    replacements.Add(placeholder, new List<Picture>() { picture });
                }

                if (typeof(IEnumerable<Picture>).IsAssignableFrom(prop.PropertyType))
                {
                    var pictures = (IEnumerable<Picture>)prop.GetValue(wordData);
                    replacements.Add(placeholder, pictures);
                }
            }

            return replacements;
        }
    }
}
