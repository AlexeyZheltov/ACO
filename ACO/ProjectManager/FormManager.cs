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
                TableProjects.Columns[5].Visible = false;
                TableProjects.Columns[6].Visible = false;

                TableProjects.Columns[0].Width = 60;
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
                    TableColumns.Columns[0].Width = 80;
                    TableColumns.Columns[1].Width = 80;
                    TableColumns.Columns[3].Width = 70;
                    TableColumns.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    TableColumns.Columns[0].HeaderText = "Проверять";
                    TableColumns.Columns[1].HeaderText = "Обязательный";
                    TableColumns.Columns[2].HeaderText = "Значение";
                    TableColumns.Columns[3].HeaderText = "Адрес";
                    TableColumns.Columns[4].Visible = false;
                    TableColumns.Columns[5].Visible = false;
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
                TBoxFirstColumnRangeValues.Text =
                    _projectManager.ActiveProject.RangeValuesStart.ToString();
                TBoxLastColumnRangeValues.Text =
                    _projectManager.ActiveProject.RangeValuesEnd.ToString();
                TBoxFirstRowRangeValues.Text =
                    _projectManager.ActiveProject.RowStart.ToString();
                TBoxSheetName.Text = _projectManager.ActiveProject.AnalysisSheetName;
            }
            else
            {
                TBoxFirstColumnRangeValues.Text = "";
                TBoxLastColumnRangeValues.Text = "";
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



        private void BtnUpdateColumns_Click(object sender, EventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.Selection;
            if ((rng?.Cells?.Count ?? 0) == 0) return;
            foreach (Excel.Range cell in rng.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Text))
                {
                    ColumnMapping mapping = new ColumnMapping(cell);
                    ColumnMapping findMapping = _mappingColumns.Find(m => m.Address == mapping.Address);
                    if (findMapping == null) _mappingColumns.Add(mapping);
                }
            }
            UpdateTableColumns();
        }

        private void TableColumns_SelectionChanged(object sender, EventArgs e)
        {
            if (TableColumns.SelectedRows.Count > 0)
            {
                DataGridViewRow row = TableColumns.SelectedRows[0];
                string address = row.Cells[3].Value?.ToString() ?? "";
                ColumnMapping cell = _mappingColumns.Find(c => c.Address == address);
                if (cell != null)
                {
                    TextBoxValue.Text = cell.Value;
                    TextBoxAddres.Text = cell.Address;
                    ChkBoxCheck.Checked = cell.Check;
                    ChkBoxObligatory.Checked = cell.Obligatory;
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
                string address = row.Cells[3].Value?.ToString() ?? "";
                ColumnMapping columnMapping = _mappingColumns.Find(x => x.Address == address);
                if (columnMapping != null)
                {
                    _mappingColumns.Remove(columnMapping);
                }
            }
        }

        private void BtnActiveCell_Click(object sender, EventArgs e)
        {

            Excel.Range activeCell = _app.Selection;
            Show();
            _selectedCell = new ColumnMapping(activeCell);
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
            try
            {
                Excel.Range xlCell = sheet.Range[address];


                if (!string.IsNullOrEmpty(address))
                {
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
            catch (Exception ex)
            {
                return;
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
            TBoxFirstColumnRangeValues.Text = rng.Column.ToString();
            int lastCol = rng.Column + rng.Columns.Count - 1;
            TBoxLastColumnRangeValues.Text = lastCol.ToString();
            int rowStart = rng.Row + rng.Rows.Count;
            TBoxFirstRowRangeValues.Text = rowStart.ToString();
        }

        /// <summary>
        ///  Сохранение активного проекта
        /// </summary>
        private void Save()
        {
            _projectManager.ActiveProject.Columns = _mappingColumns;
            _projectManager.ActiveProject.Save();

            _projectManager.ActiveProject.AnalysisSheetName = TBoxSheetName.Text;
            _projectManager.ActiveProject.RowStart = int.TryParse(TBoxFirstRowRangeValues.Text, out int fr) ? fr : 0;
            _projectManager.ActiveProject.RangeValuesStart = int.TryParse(TBoxFirstColumnRangeValues.Text, out int fc) ? fc : 0;
            _projectManager.ActiveProject.RangeValuesEnd = int.TryParse(TBoxLastColumnRangeValues.Text, out int lc) ? lc : 0;
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

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}