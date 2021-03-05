using ACO.ExcelHelpers;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.ProjectManager
{
    public partial class FormManager : Form
    {
        private ProjectManager _projectManager;
        private List<Cell> _mappingColumns;
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
                ProjectsTable.DataSource = _projectManager.Projects;
                ProjectsTable.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ProjectsTable.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ProjectsTable.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
          if (_projectManager.ActiveProject != null)
            {
                _mappingColumns = _projectManager.ActiveProject.Columns;
                if ((_mappingColumns?.Count??0) >0)
                {
                    TableColumns.DataSource = _mappingColumns;
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
                try
                {
                    Excel.Application app= Globals.ThisAddIn.Application;
                    //app.ActiveWorkbook.Names.
                    TextBoxCellName.Text = cell.Value?.ToString() ?? "";
                    TextBoxCellName.Text = cell.Name?.Range?.Name ?? "";
                }
                catch (Exception) { }
                TextBoxValue.Text = cell.Value?.ToString() ?? "";
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            Cell cell = new Cell();
            cell.Name = TextBoxCellName.Text;
            cell.Value = TextBoxValue.Text;
            cell.Row = int.Parse(TextBoxRow.Text);
            cell.Column = int.Parse(TextBoxColumn.Text);
            cell.Address = TextBoxAddres.Text;
            _mappingColumns.Add( cell );
            UpdateTableColumns();
        }

        private void UpdateTableColumns()
        {
            TableColumns.DataSource = _mappingColumns;
            TableColumns.Update();
        }

        private void BtnAccept_Click(object sender, EventArgs e)
        {
            _projectManager.ActiveProject.Columns = _mappingColumns;
            _projectManager.ActiveProject.Save();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void TextBoxCellName_TextChanged(object sender, EventArgs e)
        {

        }

        private void TextBoxColumn_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void TextBoxRow_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
