using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace EasyOffice.Utils
{
    public static class CommonHelper
    {
        public static void Copy<T>(T from, T to)
        {
            Type type = typeof(T);
            PropertyInfo[] props = type.GetProperties();
            foreach (PropertyInfo prop in props)
            {
                if (prop.CanWrite)
                {
                    prop.SetValue(from, prop.GetValue(to));
                }
            }
        }

        public static string GetExtByUrl(string srcFileUrl)
        {
            string ext = string.Empty;
            if (string.IsNullOrWhiteSpace(srcFileUrl) || !srcFileUrl.Contains("."))
            {
                return ext;
            }

            string[] arrayUrls = srcFileUrl.Split('.');
            ext = arrayUrls[arrayUrls.Length - 1];
            return ext;
        }
    }
}
