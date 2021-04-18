using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Collections.Generic;

namespace ACO
{
    public partial class FrmColorCommentsFomat : Form
    {
        public List<ConditionFormat> _ListCondintions;
        ConditionFormat _ConditionFormat;

        ConditonsFormatManager manager;
        //ConditionsFormatManager formatManager;
        Dictionary<string, Excel.XlFormatConditionOperator> ConditionOperator;
        public FrmColorCommentsFomat()
        {
            InitializeComponent();
            manager = new ConditonsFormatManager();
            _ListCondintions = manager.ListConditionFormats;
            ConditionOperator = new Dictionary<string, Excel.XlFormatConditionOperator>()
            {
                {"Больше",Excel.XlFormatConditionOperator.xlGreater },
                {"Больше равно",Excel.XlFormatConditionOperator.xlGreaterEqual },
                {"Меньше",Excel.XlFormatConditionOperator.xlLess },
                {"Меньше равно",Excel.XlFormatConditionOperator.xlLessEqual },
                {"Между",Excel.XlFormatConditionOperator.xlBetween }
            };

            FillData();
        }


        private void FillData()
        {
            customDataGrid.Rows.Clear();          

            foreach (ConditionFormat conditionFormat in _ListCondintions)
            {              
                conditionFormat.ID = _ListCondintions.IndexOf(conditionFormat);
                customDataGrid.Rows.Add(conditionFormat.ID,
                                        conditionFormat.ColumnName, conditionFormat.Operator,
                                        conditionFormat.Formula1, conditionFormat.Formula2);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            colorDialog.Color = richTextBox1.BackColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.BackColor = colorDialog.Color;
                if (_ConditionFormat != null)
                _ConditionFormat.InteriorColor = colorDialog.Color;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            fontDialog.Font = richTextBox1.Font;
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.Font = fontDialog.Font;
                if (_ConditionFormat != null)
                    _ConditionFormat.Font = fontDialog.Font;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            colorDialog.Color = richTextBox1.ForeColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.ForeColor = colorDialog.Color;
                if (_ConditionFormat != null)
                    _ConditionFormat.ForeColor = colorDialog.Color;
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
            foreach(DataGridViewRow row in customDataGrid.Rows)
            {
                int id =(int) row.Cells[0].Value;
                string oprator = row.Cells[2].Value.ToString();
                string formula1 = row.Cells[3].Value.ToString();
                string formula2 = row.Cells[4].Value.ToString();

                _ConditionFormat = _ListCondintions[id];
                _ConditionFormat.Operator = oprator;
                _ConditionFormat.Formula1 = double.TryParse(formula1,out double d)? d: 0;
                _ConditionFormat.Formula2 = double.TryParse(formula2, out double d2) ? d2 : 0;

            }

            manager.ListConditionFormats = _ListCondintions;
            manager.Save();
        }

        private void customDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            int id = (int)customDataGrid.Rows[e.RowIndex].Cells[0].Value;
            if (id < _ListCondintions.Count)
            {
                _ConditionFormat = _ListCondintions[id];
                richTextBox1.ForeColor = _ConditionFormat.ForeColor;
                richTextBox1.BackColor = _ConditionFormat.InteriorColor;
                richTextBox1.Font = _ConditionFormat.Font;

               
            }
        }
    }
}
