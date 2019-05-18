using EasyOffice.Enums;
using EasyOffice.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Models.Word
{
    /// <summary>
    /// 段落
    /// </summary>
    public class Paragraph : IWordElement
    {
        /// <summary>
        /// 文本
        /// </summary>
        public Run Run { get; set; }
        public Alignment Alignment { get; set; } = Alignment.LEFT;
    }
}
