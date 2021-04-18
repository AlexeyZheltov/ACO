using System;
using Excel = Microsoft.Office.Interop.Excel;

using System.Windows.Forms;
using ACO.Settings;
using System.Collections.Generic;
using System.Drawing;

namespace ACO
{
    public partial class FrmColorCommentsFomat : Form
    {
        public List<ConditionFormat> ListCondintions;

        Dictionary<string, Excel.XlFormatConditionOperator> ConditionOperator;
        public FrmColorCommentsFomat()
        {
            InitializeComponent();

            ConditionOperator = new Dictionary<string, Excel.XlFormatConditionOperator>()
            {
                {"Больше",Excel.XlFormatConditionOperator.xlGreater },
                {"Больше равно",Excel.XlFormatConditionOperator.xlGreaterEqual },
                {"Меньше",Excel.XlFormatConditionOperator.xlLess },
                {"Меньше равно",Excel.XlFormatConditionOperator.xlLessEqual },
                {"Между",Excel.XlFormatConditionOperator.xlBetween }
            };
            ListCondintions = new List<ConditionFormat>();
            ListCondintions.Add(
                new ConditionFormat()
                {
                    ColumnName = "",
                    Operator = "Между",
                    FontName = "Tahoma",
                    FontSize = 10,
                    FontStyle = FontStyle.Regular,
                    ForeColor = Color.AliceBlue,
                    InteriorColor = Color.Yellow,
                    Formula1 = -0.1,
                    Formula2 = -0.15
                }
            );

            ListCondintions.Add(
            new ConditionFormat()
            {
                ColumnName = "",
                Operator = "Меньше равно",
                FontName = "Tahoma",
                FontSize = 10,
                FontStyle = FontStyle.Regular,
                ForeColor = Color.White,
                InteriorColor = Color.Red,
                Formula1 = -0.15,
            }
            );

            ListCondintions.Add(
                 new ConditionFormat()
                 {
                     ColumnName = "",
                     Operator = "Между",
                     FontName = "Tahoma",
                     FontSize = 10,
                     FontStyle = FontStyle.Regular,
                     ForeColor = Color.Red,
                     InteriorColor = Color.Yellow,
                     Formula1 = 0.1,
                     Formula2 = 0.15
                 }
            );

            FillData();
        }


        private void FillData()
        {
            customDataGrid.Rows.Clear();

            foreach (DataGridViewRow row in customDataGrid.Rows)
            {
                var cellOprerators = (DataGridViewComboBoxCell)(row.Cells[customDataGrid.Columns[1].Name]);
                BindingSource source = new BindingSource();
                foreach (string operatorEqual in ConditionOperator.Keys)
                {
                    source.Add(operatorEqual);
                }
                cellOprerators.DataSource = source;
            }

            foreach (ConditionFormat conditionFormat in ListCondintions)
            {
                customDataGrid.Rows.Add(conditionFormat.ColumnName, conditionFormat.Operator,
                                        conditionFormat.Formula1, conditionFormat.Formula2);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            colorDialog.Color = richTextBox1.BackColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.BackColor = colorDialog.Color;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            fontDialog.Font = richTextBox1.Font;
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.Font = fontDialog.Font;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            colorDialog.Color = richTextBox1.ForeColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.ForeColor = colorDialog.Color;
            }
        }

        private void BtnSet_Click(object sender, EventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range cell = app.ActiveCell;
            //  Excel.Worksheet sh = app.ActiveSheet;
            // Excel.Range rng = sh.UsedRange;
            cell.Interior.Color = richTextBox1.BackColor;
            cell.Font.Name = richTextBox1.Font.Name;
            cell.Font.Bold = richTextBox1.Font.Bold;
            cell.Font.Size = richTextBox1.Font.Size;

            cell.Font.Color = richTextBox1.ForeColor;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range cell = app.ActiveCell;
            Excel.FormatCondition condition = cell.FormatConditions.Add(
            Type: Excel.XlFormatConditionType.xlCellValue,
            Operator: Excel.XlFormatConditionOperator.xlEqual,
            Formula1: "=\"1\"");

            condition.Interior.Color = richTextBox1.BackColor;
            condition.Font.Name = richTextBox1.Font.Name;
            condition.Font.Bold = richTextBox1.Font.Bold;
            condition.Font.Size = richTextBox1.Font.Size;
            condition.Font.Color = richTextBox1.ForeColor;
           
            //Excel.FormatCondition condition1 = cell.FormatConditions.Add(
            //Excel.XlFormatConditionType.cu ConditionalFormatType.custom);
            //Type: Excel.XlFormatConditionType.xlTextString,
            //Operator: Excel.XlFormatConditionOperator.xl,
            //Formula1: "=\"1\"");

        }


        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range cell = app.ActiveCell;
            Excel.FormatCondition condition = cell.FormatConditions.Add(
            Type: Excel.XlFormatConditionType.xlCellValue,
            Operator: Excel.XlFormatConditionOperator.xlBetween,
            Formula1: "=\"1\"",
            Formula2: "=\"100\""
            );
            condition.Interior.Color = richTextBox1.BackColor;
            condition.Font.Name = richTextBox1.Font.Name;
            condition.Font.Bold = richTextBox1.Font.Bold;
            condition.Font.Size = richTextBox1.Font.Size;
            condition.Font.Color = richTextBox1.ForeColor;
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void BtnAccept_Click(object sender, EventArgs e)
        {

        }
    }
}
