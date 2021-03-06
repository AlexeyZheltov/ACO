using ACO.ExcelHelpers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.ProjectManager
{
    public partial class FormManager : Form
    {
        private ProjectManager _projectManager;
        private List<ColumnMapping> _mappingColumns;
        private ColumnMapping _selectedCell ; 
        public FormManager()
        {
            InitializeComponent();
            LoadProjects();
        }

        private void LoadProjects()
        {
            _projectManager = new ProjectManager();
            if (_projectManager.Projects.Count > 0)
            {
                TableProjects.DataSource = _projectManager.Projects;
                TableProjects.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                TableProjects.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                TableProjects.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            if (_projectManager.ActiveProject != null)
            {
                _mappingColumns = _projectManager.ActiveProject.Columns;
                if ((_mappingColumns?.Count ?? 0) > 0)
                {
                    //  TableColumns.DataSource = _mappingColumns;
                    UpdateTableColumns();
                    TableColumns.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[TableColumns.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    TableColumns.Columns[0].HeaderText = "Ячейка";
                    TableColumns.Columns[1].HeaderText = "Проверять";
                    TableColumns.Columns[2].HeaderText = "Обязательный";
                    TableColumns.Columns[3].HeaderText = "Адрес";
                    TableColumns.Columns[4].HeaderText = "Строка";
                    TableColumns.Columns[5].HeaderText = "Столбец";
                    TableColumns.Columns[6].HeaderText = "Значение";
                }
            }
        }

        private void BtnAddProject_Click(object sender, EventArgs e)
        {
            string name = TboxProjectName.Text;
            new ProjectManager().CreateProject(name);
            LoadProjects();
        }

        //private void BtnActiveCell_Click(object sender, EventArgs e)
        //{
        //    Excel.Range cell = Globals.ThisAddIn.Application.ActiveCell;
        //    if (cell != null)
        //    {
        //        TextBoxRow.Text = cell.Row.ToString();
        //        TextBoxColumn.Text = cell.Column.ToString();
        //        TextBoxAddres.Text = cell.Address;
        //        ChkBoxCheck.Checked = false;
        //        ChkBoxObligatory.Checked = false;
        //        try
        //        {
        //            TextBoxCellName.Text = cell.Value?.ToString() ?? "";
        //            TextBoxCellName.Text = cell.Name?.Range?.Name ?? "";
        //        }
        //        catch (Exception) { }
        //        TextBoxValue.Text = cell.Value?.ToString() ?? "";
        //    }
        //}


        //private void BtnAdd_Click(object sender, EventArgs e)
        //{
        //    ColumnMapping cell = new ColumnMapping();
        //    string name = TextBoxCellName.Text;
        //    if (string.IsNullOrEmpty(name)) return;
        //    cell.Name = name;

        //    string value = TextBoxValue.Text;
        //    if (string.IsNullOrEmpty(value)) return;
        //    cell.Value = TextBoxValue.Text;

        //    if (!int.TryParse(TextBoxRow.Text, out int row)) return;
        //    cell.Row = row;

        //    if (!int.TryParse(TextBoxColumn.Text, out int col)) return;
        //    cell.Column = col;

        //    cell.Address = TextBoxAddres.Text;
        //    cell.Check = ChkBoxCheck.Checked;
        //    cell.Obligatory = ChkBoxObligatory.Checked;
        //    ColumnMapping findcell = _mappingColumns.Find(c => c.Address == cell.Address);
        //    if (findcell != null)
        //    {
        //        _mappingColumns.Remove(findcell);
        //    }
        //    _mappingColumns.Add(cell);
        //    UpdateTableColumns();
        //}

        private void UpdateTableColumns()
        {
            BindingSource Source = new BindingSource();
            for (int i = 0; i < _mappingColumns.Count; i++)
            {
                Source.Add(_mappingColumns[i]);
            };
            TableColumns.DataSource = Source;            
        }

        private void BtnAccept_Click(object sender, EventArgs e)
        {
            _projectManager.ActiveProject.Columns = _mappingColumns;
            _projectManager.ActiveProject.Save();
        }
        
        private void BtnUpdateColumns_Click(object sender, EventArgs e)
        {          
            Excel.Application app = Globals.ThisAddIn.Application;         
            Excel.Range rng = app.Selection;
            if ((rng?.Cells?.Count ?? 0) == 0) return;            
            foreach (Excel.Range cell in rng.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Value))
                {
                    ColumnMapping mapping = new ColumnMapping(cell);
                    ColumnMapping findMapping = _mappingColumns.Find(m => m.Address == mapping.Address);
                    if (findMapping == null)  _mappingColumns.Add(mapping);
                }
            }
            UpdateTableColumns();
        }

        private void TableColumns_SelectionChanged(object sender, EventArgs e)
        {
            if (TableColumns.SelectedRows.Count > 0)
            {
                DataGridViewRow row = TableColumns.SelectedRows[0];             
                string address= row.Cells[3].Value?.ToString() ?? "";
                ColumnMapping cell = _mappingColumns.Find(c => c.Address == address);
                if (cell != null)
                {
                    TextBoxValue.Text = cell.Value;
                    TextBoxAddres.Text = cell.Address;
                    ChkBoxCheck.Checked = cell.Check;
                    ChkBoxObligatory.Checked = cell.Obligatory;
                    //TextBoxCellName.Text = cell.Name;
                    //TextBoxRow.Text = cell.Row.ToString();
                    //TextBoxColumn.Text = cell.Column.ToString();
                }
            }
        }

        private void BtnCheckCells_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            for (int i = 0; i < TableColumns.Rows.Count; i++)
            {
                DataGridViewRow row = TableColumns.Rows[i];
                string address = row.Cells[3].Value?.ToString() ?? "";
                ColumnMapping cell = _mappingColumns.Find(c => c.Address == address);
                row.Cells[1].Style.BackColor = cell.CheckSheet(sheet) ? Color.White : Color.Red;
            }
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            ColumnMapping findcell = _mappingColumns.Find(c => c.Address == TextBoxAddres.Text);
            if (findcell != null)
            {
                _mappingColumns.Remove(findcell);
            }          
            UpdateTableColumns();
        }

        private void BtnActiveCell_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            Hide();
            sheet. SelectionChange += Sheet_SelectionChange;
        }

        private void Sheet_SelectionChange(Excel.Range Target)
        {
            Show();
           
            _selectedCell = new ColumnMapping(Target);

            if (_selectedCell != null)
            {               
                TextBoxAddres.Text = _selectedCell.Address;
                ChkBoxCheck.Checked = false;
                ChkBoxObligatory.Checked = false;
                TextBoxValue.Text = _selectedCell.Value?.ToString() ?? "";
            }
        }

    

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            string address = TextBoxAddres.Text;
            if (!string.IsNullOrEmpty(address))
            {
                Excel.Range xlCell = sheet.Range[address];
                _selectedCell = new ColumnMapping(xlCell);
                string value = TextBoxValue.Text;
                    if (!string.IsNullOrEmpty(value)) _selectedCell.Value = value;
                _selectedCell.Check = ChkBoxCheck.Checked;
                _selectedCell.Obligatory = ChkBoxObligatory.Checked;
                ColumnMapping findcell = _mappingColumns.Find(c => c.Address == _selectedCell.Address);
                if (findcell != null)
                {
                    _mappingColumns.Remove(findcell);
                }
                _mappingColumns.Add(_selectedCell);
                UpdateTableColumns();
            }
        }
    }
}
