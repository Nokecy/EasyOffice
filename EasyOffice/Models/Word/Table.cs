using EasyOffice.Interfaces;
using EasyOffice.Models.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace EasyOffice.Models.Word
{
    public class Table : IWordElement
    {
        public List<TableRow> Rows { get; set; }

        public int Width = 5000;
    }
}
