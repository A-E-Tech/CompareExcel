using System;
using System.Collections.Generic;
using System.Text;

namespace CompareExcelCore.Class
{
    public class Tag
    {
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public string MV { get; set; }
        public string CV { get; set; }
        public string? Value { get; set; }
    }
}
