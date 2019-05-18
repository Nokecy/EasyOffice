using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Constants
{
    public static class RegexConstant
    {
        /// <summary>
        /// 邮箱
        /// </summary>
        public const string EMAIL_REGEX = @"\w[-\w.+]*@([A-Za-z0-9][-A-Za-z0-9]+\.)+[A-Za-z]{2,14}";

        /// <summary>
        /// 国内手机号
        /// </summary>
        public const string MOBILE_CHINA_REGEX = @"0?(13|14|15|17|18|19)[0-9]{9}";

        /// <summary>
        /// 身份证号
        /// </summary>
        public const string IDENTITY_NUMBER_REGEX = @"^(^\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$";

        /// <summary>
        /// 非空
        /// </summary>
        public const string NOT_EMPTY_REGEX = @"\S";

        /// <summary>
        /// 性别
        /// </summary>
        public const string GENDER_REGEX = @"^男$|^女$";

        /// <summary>
        /// 网址URL
        /// </summary>
        public const string URL_REGEX = @"^((https|http|ftp|rtsp|mms)?:\/\/)[^\s]+";

        /// <summary>
        /// 国内车牌号
        /// </summary>
        public const string CAR_CODE_REGEX = @"^[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领A-Z]{1}[A-Z]{1}[A-Z0-9]{4}[A-Z0-9挂学警港澳]{1}$";
    }
}
