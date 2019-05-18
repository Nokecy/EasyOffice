using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Enums
{
    public enum ValidateModeEnum
    {
        /// <summary>
        /// 校验失败后继续校验
        /// </summary>
        Continue,

        /// <summary>
        /// 校验失败后停止校验
        /// </summary>
        StopOnFirstFailure
    }
}
