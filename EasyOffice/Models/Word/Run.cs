using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Models.Word
{
    public class Run
    {
        public string Text { get; set; } = "";
        public string Color { get; set; } = "black";
        public int FontSize { get; set; } = 12;
        public string FontFamily { get; set; } = "等线 (中文正文)";
        public bool IsBold { get; set; } = false;
        public List<Picture> Pictures { get; set; }
    }
}
