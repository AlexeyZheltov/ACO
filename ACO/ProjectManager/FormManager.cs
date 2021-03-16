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
        private Excel.Application _app = Globals.ThisAddIn.Application;
        private ProjectManager _projectManager;
        private List<ColumnMapping> _mappingColumns;
        private ColumnMapping _selectedCell;
        public FormManager()
        {
            InitializeComponent();
            TableColumns.ReadOnly = false;
            TableProjects.ReadOnly = false;
            LoadData();
        }
        private void LoadData()
        {
            LoadProjects();
            LoadColumns();
            LoadRangeValues();
        }

        private void LoadProjects()
        {
            _projectManager = new ProjectManager();
            if (_projectManager.Projects.Count > 0)
            {
                // TableProjects.DataSource = _projectManager.Projects;
                UpdateTableProject();
                TableProjects.Columns[0].HeaderText = "Текущий";
                TableProjects.Columns[1].HeaderText = "Проект";
                TableProjects.Columns[2].HeaderText = "Путь";
                TableProjects.Columns[3].Visible = false;
                TableProjects.Columns[4].Visible = false;

                TableProjects.Columns[0].Width = 70;
                TableProjects.Columns[1].Width = 120;

                TableProjects.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
            else
            {
                TableProjects.Rows.Clear();
                TableProjects.ColumnHeadersVisible = false;
            }
        }

        private void LoadColumns()
        {
            if (_projectManager.ActiveProject != null)
            {
                _mappingColumns = _projectManager.ActiveProject.Columns;
                if ((_mappingColumns?.Count ?? 0) > 0)
                {
                    UpdateTableColumns();
                    TableColumns.Columns[0].Width = 70;
                    TableColumns.Columns[1].Width = 90;
                    TableColumns.Columns[3].Width = 60;
                    TableColumns.Columns[2].ReadOnly = true;
                    TableColumns.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    TableColumns.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    TableColumns.Columns[0].HeaderText = "Проверять";
                    TableColumns.Columns[1].HeaderText = "Обязательный";
                    TableColumns.Columns[2].HeaderText = "Название";
                    TableColumns.Columns[3].HeaderText = "Столбец";
                }
                else
                {
                    TableColumns.Rows.Clear();
                    TableColumns.ColumnHeadersVisible = false;
                }
            }
        }
        private void LoadRangeValues()
        {
            if (_projectManager.ActiveProject != null)
            {
                TBoxFirstRowRangeValues.Text =
                    _projectManager.ActiveProject.RowStart.ToString();
                TBoxSheetName.Text = _projectManager.ActiveProject.AnalysisSheetName;
            }
            else
            {
                TBoxFirstRowRangeValues.Text = "";
                TBoxSheetName.Text = "";
            }
        }

        private void BtnAddProject_Click(object sender, EventArgs e)
        {
            string name = TboxProjectName.Text;
            if (!string.IsNullOrWhiteSpace(name))
            {
                ProjectManager projectManager = new ProjectManager();
                projectManager.CreateProject(name);
                LoadData();
            }
        }

        private void UpdateTableProject()
        {
            BindingSource Source = new BindingSource();
            for (int i = 0; i < _projectManager.Projects.Count; i++)
            {
                Source.Add(_projectManager.Projects[i]);
            };

            TableProjects.DataSource = Source;
        }

        private void UpdateTableColumns()
        {
            BindingSource Source = new BindingSource();
            for (int i = 0; i < _mappingColumns.Count; i++)
            {
                Source.Add(_mappingColumns[i]);
            };
            TableColumns.DataSource = Source;
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (_selectedCell != null)
            {
                _mappingColumns.Remove(_selectedCell);
            }
            UpdateTableColumns();
        }
        /// <summary>
        ///  Удалить строку
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TableColumns_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (e.RowIndex > 0)
            {
                DataGridViewRow row = TableColumns.Rows[e.RowIndex];
                string name = row.Cells[2].Value?.ToString() ?? "";
                ColumnMapping columnMapping = _mappingColumns.Find(x => x.Name == name);
                if (columnMapping != null)
                {
                    _mappingColumns.Remove(columnMapping);
                }
            }
        }

        private void BtnSelect_Click(object sender, EventArgs e)
        {
            if (TableProjects.SelectedRows.Count > 0)
            {
                string name = TableProjects.SelectedRows[0].Cells[0].Value.ToString() ?? "";
                Project newActiveProject = _projectManager.Projects.Find(p => p.Name == name);
                if (newActiveProject != null)
                {
                    _projectManager.ActiveProject = newActiveProject;
                    LoadData();
                }
            }
        }

        private void TableColumns_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                _projectManager.ActiveProject.Save();
            }
        }

        private void TableProjects_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            if (e.ColumnIndex == 1 && e.RowIndex >= 0)
            {
                string name = TableProjects.Rows[e.RowIndex].Cells[1].Value?.ToString() ?? "";
                string filePath = TableProjects.Rows[e.RowIndex].Cells[2].Value?.ToString() ?? "";
                Project project = _projectManager.Projects.Find(x => x.FileName == filePath);

                if (!string.IsNullOrEmpty(name) && project != null)
                {
                    project.Name = name;
                    project.Save();
                }
            }
        }

        private void TableProjects_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;

            if (row >= 0)
            {
                string name = TableProjects.Rows[e.RowIndex].Cells[1].Value?.ToString() ?? "";
                string filePath = TableProjects.Rows[e.RowIndex].Cells[2].Value?.ToString() ?? "";
                Project project = _projectManager.Projects.Find(x => x.FileName == filePath);
                if (e.ColumnIndex == 0)
                {
                    var active = TableProjects.Rows[e.RowIndex].Cells[0].Value;
                    if ((bool)active || string.IsNullOrWhiteSpace(name)) return;
                    TableProjects.Rows[e.RowIndex].Cells[0].Value = true;
                    Properties.Settings.Default.ActiveProjectName = name;
                    _projectManager.SetActiveProject();
                    LoadData();
                }
            }
        }

        private void BtnOpenFolserSettings_Click(object sender, EventArgs e)
        {
            string folder = ProjectManager.GetFolderProjects();
            System.Diagnostics.Process.Start(folder);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            
        }

        /// <summary>
        ///  Кнопка Выделенный диапазон \ вкладка 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSetSelectedRangeValues_Click(object sender, EventArgs e)
        {
            Excel.Range rng = _app.Selection;
            if (rng is null) return;
            TBoxSheetName.Text = rng.Parent.name;
            int rowStart = rng.Row + rng.Rows.Count;
            TBoxFirstRowRangeValues.Text = rowStart.ToString();
        }

        /// <summary>
        ///  Сохранение активного проекта
        /// </summary>
        private void Save()
        {
            _projectManager.ActiveProject.Columns = _mappingColumns;
            _projectManager.ActiveProject.AnalysisSheetName = TBoxSheetName.Text;
            _projectManager.ActiveProject.RowStart = int.TryParse(TBoxFirstRowRangeValues.Text, out int fr) ? fr : 0;
            _projectManager.ActiveProject.Save();
        }

        /// <summary>
        ///  Кнопка сохранить
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAccept_Click(object sender, EventArgs e)
        {
            Save();
            Close();
        }

        private void FormManager_Load(object sender, EventArgs e)
        {

        }
    }
}