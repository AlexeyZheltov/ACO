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
                ProjectsTable.DataSource = _projectManager.Projects;
                ProjectsTable.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ProjectsTable.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                ProjectsTable.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
          if (_projectManager.ActiveProject != null)
            {
                _mappingColumns = _projectManager.ActiveProject.Columns;
                TableColumns.DataSource = _mappingColumns;

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
                TextBoxCellName.Text = cell.Value?.ToString() ?? "";
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            Cell cell = new Cell();
            cell.Value = TextBoxCellName.Text;
            cell.Row = int.Parse(TextBoxRow.Text);
            cell.Column = int.Parse(TextBoxColumn.Text);
            cell.Address = TextBoxAddres.Text;

            _mappingColumns.Add(new ColumnMapping() { Name = "", Cell = cell });
            UpdateTableColumns();
        }

        private void UpdateTableColumns()
        {
            TableColumns.DataSource = _mappingColumns;
        }

        private void BtnAccept_Click(object sender, EventArgs e)
        {

        }
    }
}
