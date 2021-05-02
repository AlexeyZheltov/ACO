using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;

namespace ACO
{
    public partial class FrmColorCommentsFomat : Form
    {
        public List<ConditionFormat> _ListCondintions;
        ConditionFormat _ConditionFormat;
        readonly ConditonsFormatManager manager;

        public FrmColorCommentsFomat()
        {
            InitializeComponent();
            manager = new ConditonsFormatManager();
            _ListCondintions = manager.ListConditionFormats;
            FillData();
        }


        private void FillData()
        {
            RulesDataGrid.Rows.Clear();

            foreach (ConditionFormat conditionFormat in _ListCondintions)
            {
                conditionFormat.ID = _ListCondintions.IndexOf(conditionFormat);
                if (conditionFormat.Operator == "Содержит")
                {
                RulesDataGrid.Rows.Add(conditionFormat.ID,
                                        conditionFormat.ColumnName, conditionFormat.Operator,
                                        conditionFormat.Text);
                }
                else 
                {
                    RulesDataGrid.Rows.Add(conditionFormat.ID,
                                       conditionFormat.ColumnName, conditionFormat.Operator,
                                       conditionFormat.Formula1, conditionFormat.Formula2);
                }
            }
        }
        private void BtnInteriorColor_Click(object sender, EventArgs e)
        {
            colorDialog.Color = richTextBox1.BackColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.BackColor = colorDialog.Color;
                if (_ConditionFormat != null)
                    _ConditionFormat.InteriorColor = colorDialog.Color;
            }
        }

        private void ChkBoxBold_CheckedChanged(object sender, EventArgs e)
        {
            FontStyle style = ChkBoxBold.Checked ? FontStyle.Bold : FontStyle.Regular;
            richTextBox1.Font = new Font("Tahoma", 10, style);
            _ConditionFormat.FontBold = ChkBoxBold.Checked;
        }


        private void BtnForeColor_Click(object sender, EventArgs e)
        {
            colorDialog.Color = richTextBox1.ForeColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.ForeColor = colorDialog.Color;
                if (_ConditionFormat != null)
                    _ConditionFormat.ForeColor = colorDialog.Color;
            }
        }


        private void BtnAccept_Click(object sender, EventArgs e)
        {
            IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "," };
            List<ConditionFormat> listCondintions = new List<ConditionFormat>();

            foreach (DataGridViewRow row in RulesDataGrid.Rows)
            {
                int id = (int)row.Cells[0].Value;
                ConditionFormat cf = _ListCondintions.Find(x => x.ID == id);
                if (cf is null)
                {
                    cf = new ConditionFormat
                    {
                        ID = _ListCondintions.Count
                    };
                }
                
                string name = row.Cells[1].Value?.ToString()??"";
                string operatorCon = row.Cells[2].Value?.ToString();
                string formula1 = row.Cells[3].Value?.ToString()??"";
                string formula2 = row.Cells[4].Value?.ToString()??"";

                cf = _ListCondintions[id];
                cf.Operator = operatorCon;
                cf.ColumnName = name;
                if (operatorCon == "Содержит" )
                {
                cf.Text = formula1;
                }
                else
                {
                cf.Formula1 = double.Parse(formula1, formatter);
                cf.Formula2 = double.Parse(formula2, formatter);
                }

                listCondintions.Add(cf);
            }

            manager.ListConditionFormats = listCondintions;
            manager.Save();
            Close();
        }

        private void CustomDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            int id = (int)RulesDataGrid.Rows[e.RowIndex].Cells[0].Value;
            if (id < _ListCondintions.Count)
            {
                _ConditionFormat = _ListCondintions[id];
                richTextBox1.ForeColor = _ConditionFormat.ForeColor;
                richTextBox1.BackColor = _ConditionFormat.InteriorColor;
                ChkBoxBold.Checked = _ConditionFormat.FontBold;
                FontStyle style = _ConditionFormat.FontBold ? FontStyle.Bold : FontStyle.Regular;
                richTextBox1.Font = new Font("Tahoma", 10, style);
            }
        }

        private void CustomDataGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex != 2) return;
            string operatorequal = RulesDataGrid.Rows[e.RowIndex].Cells[2].Value?.ToString();
            if (operatorequal == "Между")
            {
                RulesDataGrid.Rows[e.RowIndex].Cells[4].ReadOnly = false;
                RulesDataGrid.Rows[e.RowIndex].Cells[4].Style.BackColor = Color.White;
            }
            else
            {
                RulesDataGrid.Rows[e.RowIndex].Cells[4].ReadOnly = true;
                RulesDataGrid.Rows[e.RowIndex].Cells[4].Style.BackColor = Color.LightGray;
                RulesDataGrid.Rows[e.RowIndex].Cells[4].Value = "";
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            _ListCondintions.Add(new ConditionFormat());
            FillData();
            RulesDataGrid.Rows[RulesDataGrid.Rows.Count - 1].Cells[3].Selected = true;
        }
               

        //private void Closing(object sender, FormClosingEventArgs e)
        //{
        //    if (e.CloseReason == CloseReason.UserClosing) DialogResult = DialogResult.Cancel;

        //}

        private void BtnCancel_Click(object sender, EventArgs e)
        {

        }
    }
}
