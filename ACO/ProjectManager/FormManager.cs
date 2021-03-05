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
                if ((_mappingColumns?.Count??0) >0)
                {
                    //  TableColumns.DataSource = _mappingColumns;
                    UpdateTableColumns();
                    TableColumns.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    TableColumns.Columns[TableColumns.Columns.Count-1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    TableColumns.Columns[0].HeaderText = "Ячейка";
                    TableColumns.Columns[1].HeaderText = "Проверять";
                    TableColumns.Columns[2].HeaderText = "Обязательный";
                    TableColumns.Columns[3].HeaderText = "Адрес";
                    TableColumns.Columns[4].HeaderText = "Значение";
                }
            }
        }

        private void BtnAddProject_Click(object sender, EventArgs e)
        {
            string name = TboxProjectName.Text;
            new ProjectManager().CreateProject(name);
            LoadProjects();
        }

        private void BtnActiveCell_Click(object sender, EventArgs e)
        {
            Excel.Range cell = Globals.ThisAddIn.Application.ActiveCell;
            if (cell != null)
            {
                TextBoxRow.Text = cell.Row.ToString();
                TextBoxColumn.Text = cell.Column.ToString();
                TextBoxAddres.Text = cell.Address;
                ChkBoxCheck.Checked = false;
                ChkBoxObligatory.Checked = false;
                try
                {                 
                    TextBoxCellName.Text = cell.Value?.ToString() ?? "";
                    TextBoxCellName.Text = cell.Name?.Range?.Name ?? "";
                }
                catch (Exception) { }
                TextBoxValue.Text = cell.Value?.ToString() ?? "";
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            ColumnMapping cell = new ColumnMapping();
            cell.Name = TextBoxCellName.Text;
            cell.Value = TextBoxValue.Text;
            cell.Row = int.Parse(TextBoxRow.Text);
            cell.Column = int.Parse(TextBoxColumn.Text);
            cell.Address = TextBoxAddres.Text;
            cell.Check = ChkBoxCheck.Checked ;
            cell.Obligatory = ChkBoxObligatory.Checked ;
            ColumnMapping findcell = _mappingColumns.Find(c => c.Name == cell.Name);
            if (findcell != null)
            {
              _mappingColumns.Remove(findcell);
            }
            _mappingColumns.Add( cell );
            UpdateTableColumns();
        }

        private void UpdateTableColumns()
        {            
            BindingSource Source = new BindingSource();
            for (int i = 0; i < _mappingColumns.Count; i++)
            {
                Source.Add(_mappingColumns[i]);
            };
            TableColumns.DataSource = Source;
           // TableColumns.Update();
        }

        private void BtnAccept_Click(object sender, EventArgs e)
        {
            _projectManager.ActiveProject.Columns = _mappingColumns;
            _projectManager.ActiveProject.Save();
        }

        private void BtnSelect_Click(object sender, EventArgs e)
        {
        }

        private void BtnUpdateColumns_Click(object sender, EventArgs e)
        {

        }

        private void TableColumns_SelectionChanged(object sender, EventArgs e)
        {
            if (TableColumns.SelectedRows.Count > 0)
            {
                //string selectedName = TableProjects.SelectedRows[0].Cells[0].Value?.ToString() ;
                //TextBoxCellName.Text = row.Cells[0].Value?.ToString() ??"" ;
                //TextBoxValue.Text = row.Cells[2].Value?.ToString() ?? "";
                //TextBoxRow.Text = row.Cells[3].Value?.ToString() ?? "";
                //TextBoxColumn.Text = row.Cells[4].Value?.ToString() ?? "" ;
                //TextBoxAddres.Text = row.Cells[5].Value?.ToString() ?? "";

                DataGridViewRow row = TableColumns.SelectedRows[0];
                string selectedName = row.Cells[0].Value?.ToString() ?? "";

                ColumnMapping cell = _mappingColumns.First(c => c.Name == selectedName); 
                if (cell != null)
                {
                    TextBoxCellName.Text = cell.Name;
                    TextBoxValue.Text = cell.Value;
                    TextBoxRow.Text = cell.Row.ToString();
                    TextBoxColumn.Text = cell.Column.ToString();
                    TextBoxAddres.Text = cell.Address;
                    ChkBoxCheck.Checked = cell.Check;
                    ChkBoxObligatory.Checked = cell.Obligatory;
                }
            }
        }

        private void BtnCheckCells_Click(object sender, EventArgs e)
        {
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

            for(int i =0; i < TableColumns.Rows.Count; i++)
            {
            DataGridViewRow row = TableColumns.SelectedRows[i];
            string selectedName = row.Cells[0].Value?.ToString() ?? "";
            ColumnMapping cell = _mappingColumns.Find(c => c.Name == selectedName);
            row.Cells[1].Style.BackColor = cell.CheckSheet(sheet) ?  Color.White : Color.Red ;
                //if (cell.CheckSheet(sheet))
                //{
                //}
                //else
                //{ 
                //    row.Cells[1].Style.BackColor =;
                //}
            }

            //foreach (ColumnMapping cell in _mappingColumns)
            //{
            //    cell.CheckSheet(sheet);
            //}
        }
    }
}
