using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Data;

namespace RCG
{
    public interface IFormatter
    {
        string Rule { get; set; }
        string FormatString { get; set; }
        FormatTypes FormaterType { get; set; }
        bool Match(DataRow dr, FormatterConfig formatterConfig);
        void Execute(int currentExcelRowIndex);
    }
    public enum FormatTypes
    {
        None = 0,
        ForeColor = 1,
        BackColor = 2,
        FontBold = 4,
        FontItalic = 8,
        FontStrikethrough = 16
    }
}
