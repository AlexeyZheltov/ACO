using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO
{
    public class ConditionFormat
    {
        public int ID { set; get; }
        public string ColumnName { set; get; } = "";

        public Excel.XlFormatConditionOperator XlFormatConditionOperator
        {
            set
            {
                _xlFormatConditionOperator = value;
                Operator = Operators.First(x => x.Value == _xlFormatConditionOperator).Key;
            }

            get
            {
                return Operators[Operator];
            }
        }
        Excel.XlFormatConditionOperator _xlFormatConditionOperator;

        public string Operator { set; get; } = "Больше";


        public Color ForeColor { set; get; } = Color.Black;

        public Color InteriorColor { set; get; } = Color.White;

        public bool FontBold { set; get; } = false;

        public double Formula1 { set; get; } = 0;

        public double Formula2 { set; get; }
        public string Text { set; get; }


        private static readonly Dictionary<string, Excel.XlFormatConditionOperator> Operators =
               new Dictionary<string, Excel.XlFormatConditionOperator>()
           {
                {"Больше",Excel.XlFormatConditionOperator.xlGreater},
                {"Больше равно",Excel.XlFormatConditionOperator.xlGreaterEqual },
                {"Меньше",Excel.XlFormatConditionOperator.xlLess },
                {"Меньше равно",Excel.XlFormatConditionOperator.xlLessEqual },
                {"Между",Excel.XlFormatConditionOperator.xlBetween },
                {"Равно",Excel.XlFormatConditionOperator.xlEqual },
                {"Не равно",Excel.XlFormatConditionOperator.xlNotEqual },
           };

        public void SetCondition(Excel.Range range)
        {
            //  Excel.Application app = Globals.ThisAddIn.Application;
            Excel.FormatCondition condition;
            if (string.IsNullOrWhiteSpace(Operator)) return;
            try
            {
                if (Operator == "Содержит")
                {
                    condition = range.FormatConditions.Add(
              Type: Excel.XlFormatConditionType.xlTextString,
              TextOperator: Excel.XlContainsOperator.xlContains,
              String: Text
              );
                }
                else if (Operator == "Между")
                {
                    condition = range.FormatConditions.Add(
                    Type: Excel.XlFormatConditionType.xlCellValue,
                    Operator: XlFormatConditionOperator,
                    Formula1: $"={Formula1}",
                    Formula2: $"={Formula2}"
                    );
                }
                else
                {
                    condition = range.FormatConditions.Add(
               Type: Excel.XlFormatConditionType.xlCellValue,
               Operator: XlFormatConditionOperator,
               Formula1: $"={Formula1}");
                }
                condition.Interior.Color = InteriorColor;
                condition.Font.Color = ForeColor;
                condition.Font.Bold = FontBold;
                condition.StopIfTrue = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("условное форманировение вызвало ошибку. " + ex.Message);
            }
        }
    }
}
