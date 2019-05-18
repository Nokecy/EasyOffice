using EasyOffice.Enums;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Models.Word
{
    public class Word
    {
        public SolutionEnum Type { get; set; }

        public byte[] WordBytes { get; set; }
    }
}
