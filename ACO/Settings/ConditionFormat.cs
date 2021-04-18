using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.Settings
{
    public class ConditionFormat
    {
        public int ID { set; get; }
        public string ColumnName { set; get; } = "";

        public Excel.XlFormatConditionOperator  xlOperator{set; get;}
        public string Operator { set; get; } = "Больше";


        public Color ForeColor { set; get; } = Color.Black;

        public Color InteriorColor { set; get; } = Color.White;

        public string FontName { set; get; } = "Arial";

        public float FontSize { set; get; } = 8;

        public FontStyle FontStyle { set; get; } = FontStyle.Regular;

        public double Formula1 { set; get; } = 0;

        public double Formula2 { set; get; } = 0;

        public Excel.Range Range { set; get; }

    }
}
